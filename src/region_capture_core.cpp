#include "region_capture_core.h"

#include <windows.h>
#include <dwmapi.h>

#include <algorithm>
#include <atomic>
#include <climits>
#include <cstdio>
#include <cstring>
#include <cwchar>
#include <memory>
#include <string>
#include <thread>
#include <vector>

// GDI+ 需要 min/max
namespace Gdiplus {
using std::max;
using std::min;
}
#include <gdiplus.h>

namespace screenshot::windows {
namespace {

constexpr int SC_PANEL_WIDTH = 140;
constexpr int SC_PANEL_HEIGHT = 140;
constexpr int SC_MAGNIFIER_HEIGHT = 74;
constexpr int SC_PANEL_MARGIN = 15;
constexpr int SC_PANEL_CORNER_RADIUS = 8;
constexpr int SC_ZOOM_FACTOR = 4;
constexpr UINT WM_SC_STOP_CAPTURE = WM_APP + 1;

enum CaptureState { CS_Idle, CS_Selecting, CS_Done, CS_Cancelled };

struct SCWindowInfo {
    HWND hwnd;
    RECT rect;
    std::wstring title;
};

struct SCGdiResources {
    HBRUSH bgBrush = NULL;
    HPEN borderPen = NULL;
    HPEN crosshairPen = NULL;
    HPEN selectionPen = NULL;
    HPEN highlightPen = NULL;
    HFONT smallFont = NULL;

    void Init() {
        bgBrush = CreateSolidBrush(RGB(52, 52, 53));
        borderPen = CreatePen(PS_SOLID, 0, RGB(102, 102, 102));
        crosshairPen = CreatePen(PS_SOLID, 1, RGB(0, 136, 255));
        selectionPen = CreatePen(PS_SOLID, 1, RGB(0, 136, 255));
        highlightPen = CreatePen(PS_SOLID, 3, RGB(0, 136, 255));

        LOGFONTW lf = {};
        lf.lfHeight = -12;
        lf.lfCharSet = DEFAULT_CHARSET;
        wcscpy_s(lf.lfFaceName, L"微软雅黑");
        smallFont = CreateFontIndirectW(&lf);
    }

    void Cleanup() {
        if (bgBrush) {
            DeleteObject(bgBrush);
            bgBrush = NULL;
        }
        if (borderPen) {
            DeleteObject(borderPen);
            borderPen = NULL;
        }
        if (crosshairPen) {
            DeleteObject(crosshairPen);
            crosshairPen = NULL;
        }
        if (selectionPen) {
            DeleteObject(selectionPen);
            selectionPen = NULL;
        }
        if (highlightPen) {
            DeleteObject(highlightPen);
            highlightPen = NULL;
        }
        if (smallFont) {
            DeleteObject(smallFont);
            smallFont = NULL;
        }
    }
};

struct CaptureContext {
    CaptureState state = CS_Idle;
    int virtualX = 0;
    int virtualY = 0;
    int virtualW = 0;
    int virtualH = 0;
    int startX = 0;
    int startY = 0;
    int endX = 0;
    int endY = 0;
    int mouseX = 0;
    int mouseY = 0;
    COLORREF currentColor = RGB(0, 0, 0);
    std::vector<SCWindowInfo> windows;
    int hoveredWindow = -1;
    HBITMAP screenBitmap = NULL;
    HDC memDC = NULL;
    HDC backDC = NULL;
    HBITMAP backBitmap = NULL;
    RECT lastPanelRect = {0, 0, 0, 0};
    RECT lastSelectionRect = {0, 0, 0, 0};
    RECT lastLabelRect = {0, 0, 0, 0};
    RECT lastHighlightRect = {0, 0, 0, 0};
    bool needFullRedraw = true;
    double dpiScale = 1.0;
    SCGdiResources gdi;
};

static HWND g_overlayWindow = NULL;
static DWORD g_captureThreadId = 0;
static std::thread g_captureThread;
static std::atomic<bool> g_isCapturing(false);
static std::atomic<bool> g_stopRequested(false);
static std::atomic<bool> g_terminalDispatched(false);
static CaptureContext* g_captureCtx = nullptr;
static CaptureCallbacks g_callbacks;

static const char base64Chars[] =
    "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789+/";

static std::string Base64Encode(const BYTE* data, size_t len) {
    std::string result;
    result.reserve(((len + 2) / 3) * 4);
    for (size_t i = 0; i < len; i += 3) {
        const unsigned int bits =
            (data[i] << 16) |
            ((i + 1 < len ? data[i + 1] : 0) << 8) |
            (i + 2 < len ? data[i + 2] : 0);
        result.push_back(base64Chars[(bits >> 18) & 0x3F]);
        result.push_back(base64Chars[(bits >> 12) & 0x3F]);
        result.push_back(i + 1 < len ? base64Chars[(bits >> 6) & 0x3F] : '=');
        result.push_back(i + 2 < len ? base64Chars[bits & 0x3F] : '=');
    }
    return result;
}

static int GetPngEncoderClsid(CLSID* pClsid) {
    UINT num = 0u;
    UINT size = 0u;
    Gdiplus::GetImageEncodersSize(std::addressof(num), std::addressof(size));
    if (size == 0u) {
        return -1;
    }

    std::unique_ptr<Gdiplus::ImageCodecInfo> codecs(
        static_cast<Gdiplus::ImageCodecInfo*>(static_cast<void*>(new BYTE[size])));
    if (codecs == nullptr) {
        return -1;
    }

    Gdiplus::GetImageEncoders(num, size, codecs.get());
    for (UINT i = 0; i < num; ++i) {
        if (std::wcscmp(codecs.get()[i].MimeType, L"image/png") == 0) {
            *pClsid = codecs.get()[i].Clsid;
            return static_cast<int>(i);
        }
    }

    return -1;
}

static bool EmitTerminalOnce() {
    bool expected = false;
    return g_terminalDispatched.compare_exchange_strong(expected, true);
}

static void EmitSelected(const SelectionPayload& payload) {
    if (g_terminalDispatched.load()) {
        return;
    }
    if (g_callbacks.onSelected) {
        g_callbacks.onSelected(payload);
    }
}

static void EmitComplete(const CompletePayload& payload) {
    if (!EmitTerminalOnce()) {
        return;
    }
    if (g_callbacks.onComplete) {
        g_callbacks.onComplete(payload);
    }
}

static void EmitCancel() {
    if (!EmitTerminalOnce()) {
        return;
    }
    if (g_callbacks.onCancel) {
        g_callbacks.onCancel();
    }
}

static void EmitError(const std::string& message) {
    if (!EmitTerminalOnce()) {
        return;
    }
    if (g_callbacks.onError) {
        g_callbacks.onError(message);
    }
}

static double GetDpiScaleFactor() {
    typedef UINT(WINAPI* GetDpiForSystemProc)();
    HMODULE user32 = GetModuleHandleW(L"user32.dll");
    if (user32 != NULL) {
        auto proc = reinterpret_cast<GetDpiForSystemProc>(
            GetProcAddress(user32, "GetDpiForSystem"));
        if (proc != nullptr) {
            UINT dpi = proc();
            double scale = dpi / 96.0;
            if (scale < 0.5) {
                scale = 0.5;
            }
            if (scale > 4.0) {
                scale = 4.0;
            }
            return scale;
        }
    }
    return 1.0;
}

struct MonitorEnumData {
    LONG minLeft;
    LONG minTop;
    LONG maxRight;
    LONG maxBottom;
    double totalDpiScale;
    int monitorCount;
};

static BOOL CALLBACK MonitorEnumProc(
    HMONITOR hMonitor, HDC, LPRECT, LPARAM dwData) {
    auto* data = reinterpret_cast<MonitorEnumData*>(dwData);

    MONITORINFOEXW mi;
    mi.cbSize = sizeof(MONITORINFOEXW);
    if (!GetMonitorInfoW(hMonitor, &mi)) {
        return TRUE;
    }

    data->minLeft = (std::min)(data->minLeft, mi.rcMonitor.left);
    data->minTop = (std::min)(data->minTop, mi.rcMonitor.top);
    data->maxRight = (std::max)(data->maxRight, mi.rcMonitor.right);
    data->maxBottom = (std::max)(data->maxBottom, mi.rcMonitor.bottom);
    data->monitorCount++;

    typedef HRESULT(WINAPI* GetDpiForMonitorProc)(HMONITOR, int, UINT*, UINT*);
    HMODULE shcore = LoadLibraryW(L"shcore.dll");
    if (shcore != NULL) {
        auto getDpiForMonitor = reinterpret_cast<GetDpiForMonitorProc>(
            GetProcAddress(shcore, "GetDpiForMonitor"));
        if (getDpiForMonitor != nullptr) {
            UINT dpiX = 0;
            UINT dpiY = 0;
            if (SUCCEEDED(getDpiForMonitor(hMonitor, 0, &dpiX, &dpiY))) {
                data->totalDpiScale = (std::max)(data->totalDpiScale, dpiX / 96.0);
            }
        }
        FreeLibrary(shcore);
    }

    return TRUE;
}

static bool CaptureVirtualScreen(
    HDC& outMemDC, HBITMAP& outBitmap, int& vx, int& vy, int& vw, int& vh,
    double& dpiScale) {
    vx = GetSystemMetrics(SM_XVIRTUALSCREEN);
    vy = GetSystemMetrics(SM_YVIRTUALSCREEN);
    vw = GetSystemMetrics(SM_CXVIRTUALSCREEN);
    vh = GetSystemMetrics(SM_CYVIRTUALSCREEN);

    MonitorEnumData enumData = {INT_MAX, INT_MAX, INT_MIN, INT_MIN, 1.0, 0};
    EnumDisplayMonitors(NULL, NULL, MonitorEnumProc, reinterpret_cast<LPARAM>(&enumData));

    int physVx = enumData.minLeft;
    int physVy = enumData.minTop;
    int physVw = enumData.maxRight - enumData.minLeft;
    int physVh = enumData.maxBottom - enumData.minTop;

    if (physVw <= 0 || physVh <= 0 || enumData.monitorCount == 0) {
        physVx = static_cast<int>(vx * dpiScale);
        physVy = static_cast<int>(vy * dpiScale);
        physVw = static_cast<int>(vw * dpiScale + 0.5);
        physVh = static_cast<int>(vh * dpiScale + 0.5);
    }

    HDC screenDC = GetDC(NULL);
    if (!screenDC) {
        return false;
    }

    outMemDC = CreateCompatibleDC(screenDC);
    if (!outMemDC) {
        ReleaseDC(NULL, screenDC);
        return false;
    }

    outBitmap = CreateCompatibleBitmap(screenDC, physVw, physVh);
    if (!outBitmap) {
        DeleteDC(outMemDC);
        ReleaseDC(NULL, screenDC);
        return false;
    }

    SelectObject(outMemDC, outBitmap);
    SetStretchBltMode(outMemDC, HALFTONE);
    SetBrushOrgEx(outMemDC, 0, 0, NULL);
    BitBlt(outMemDC, 0, 0, physVw, physVh, screenDC, physVx, physVy, SRCCOPY);

    if (vw > 0 && vh > 0) {
        dpiScale = static_cast<double>(physVw) / vw;
    }

    ReleaseDC(NULL, screenDC);
    return true;
}

static bool CreateBackBuffer(HDC& outDC, HBITMAP& outBmp, int w, int h) {
    HDC screenDC = GetDC(NULL);
    if (!screenDC) {
        return false;
    }

    outDC = CreateCompatibleDC(screenDC);
    if (!outDC) {
        ReleaseDC(NULL, screenDC);
        return false;
    }

    outBmp = CreateCompatibleBitmap(screenDC, w, h);
    if (!outBmp) {
        DeleteDC(outDC);
        ReleaseDC(NULL, screenDC);
        return false;
    }

    SelectObject(outDC, outBmp);
    ReleaseDC(NULL, screenDC);
    return true;
}

static COLORREF GetPixelColorFromBitmap(
    HDC memDC, int x, int y, int vx, int vy, double dpiScale) {
    int lx = x - vx;
    int ly = y - vy;
    int px = static_cast<int>(lx * dpiScale + 0.5);
    int py = static_cast<int>(ly * dpiScale + 0.5);
    return GetPixel(memDC, px, py);
}

static void ColorrefToStrings(COLORREF color, char* hexBuf, char* rgbBuf) {
    const int r = color & 0xFF;
    const int g = (color >> 8) & 0xFF;
    const int b = (color >> 16) & 0xFF;
    sprintf_s(hexBuf, 32, "#%02X%02X%02X", r, g, b);
    sprintf_s(rgbBuf, 32, "%d, %d, %d", r, g, b);
}

static BOOL CALLBACK SCEnumWindowsProc(HWND hwnd, LPARAM lParam) {
    auto* windows = reinterpret_cast<std::vector<SCWindowInfo>*>(lParam);

    if (!IsWindowVisible(hwnd)) {
        return TRUE;
    }

    LONG_PTR exStyle = GetWindowLongPtrW(hwnd, GWL_EXSTYLE);
    if (exStyle & WS_EX_TOOLWINDOW) {
        return TRUE;
    }

    LONG_PTR style = GetWindowLongPtrW(hwnd, GWL_STYLE);
    if (style == 0) {
        return TRUE;
    }

    BOOL isCloaked = FALSE;
    HRESULT cloakedResult =
        DwmGetWindowAttribute(hwnd, DWMWA_CLOAKED, &isCloaked, sizeof(isCloaked));
    if (SUCCEEDED(cloakedResult) && isCloaked) {
        return TRUE;
    }

    const int maxClassName = 256;
    WCHAR className[maxClassName] = {0};
    int classNameLen = GetClassNameW(hwnd, className, maxClassName);
    if (classNameLen > 0) {
        if (wcscmp(className, L"Windows.UI.Core.CoreWindow") == 0) {
            RECT clientRect;
            if (!GetClientRect(hwnd, &clientRect)) {
                return TRUE;
            }

            int clientW = clientRect.right - clientRect.left;
            int clientH = clientRect.bottom - clientRect.top;
            if (clientW < 100 || clientH < 100) {
                return TRUE;
            }
        }

        if (wcscmp(className, L"ApplicationFrameWindow") == 0) {
            if (IsIconic(hwnd)) {
                return TRUE;
            }

            DWORD cloakedReason = 0;
            DwmGetWindowAttribute(hwnd, DWMWA_CLOAKED, &cloakedReason, sizeof(cloakedReason));
            if (cloakedReason != 0) {
                return TRUE;
            }
        }
    }

    int titleLen = GetWindowTextLengthW(hwnd);
    if (titleLen == 0 || hwnd == GetDesktopWindow()) {
        return TRUE;
    }

    std::wstring title(titleLen + 1, L'\0');
    GetWindowTextW(hwnd, &title[0], titleLen + 1);
    title.resize(titleLen);

    RECT rect = {};
    HRESULT frameResult =
        DwmGetWindowAttribute(hwnd, DWMWA_EXTENDED_FRAME_BOUNDS, &rect, sizeof(rect));
    if (FAILED(frameResult) && !GetWindowRect(hwnd, &rect)) {
        return TRUE;
    }

    int w = rect.right - rect.left;
    int h = rect.bottom - rect.top;
    if (w < 50 || h < 50) {
        return TRUE;
    }

    SCWindowInfo info;
    info.hwnd = hwnd;
    info.rect = rect;
    info.title = title;
    windows->push_back(info);
    return TRUE;
}

static std::vector<SCWindowInfo> EnumWindowsForCapture() {
    std::vector<SCWindowInfo> windows;
    EnumWindows(SCEnumWindowsProc, reinterpret_cast<LPARAM>(&windows));
    return windows;
}

static int FindWindowAtPoint(const std::vector<SCWindowInfo>& windows, int x, int y) {
    for (size_t i = 0; i < windows.size(); ++i) {
        const RECT& rect = windows[i].rect;
        if (x >= rect.left && x < rect.right && y >= rect.top && y < rect.bottom) {
            return static_cast<int>(i);
        }
    }
    return -1;
}

static void CalcPanelPosition(
    int mx, int my, int vx, int vy, int vw, int vh, int& px, int& py) {
    int sr = vx + vw;
    int sb = vy + vh;
    px = mx + SC_PANEL_MARGIN;
    py = my + SC_PANEL_MARGIN;
    if (px + SC_PANEL_WIDTH > sr) {
        px = mx - SC_PANEL_WIDTH - SC_PANEL_MARGIN;
    }
    if (py + SC_PANEL_HEIGHT > sb) {
        py = my - SC_PANEL_HEIGHT - SC_PANEL_MARGIN;
    }
    if (px < vx) {
        px = vx + SC_PANEL_MARGIN;
    }
    if (py < vy) {
        py = vy + SC_PANEL_MARGIN;
    }
}

static void RestoreDirtyRegion(HDC backDC, HDC memDC, const RECT& dirty, double dpiScale) {
    int w = dirty.right - dirty.left;
    int h = dirty.bottom - dirty.top;
    if (w <= 0 || h <= 0) {
        return;
    }

    int x = (std::max)(dirty.left, 0);
    int y = (std::max)(dirty.top, 0);
    w = dirty.right - x;
    h = dirty.bottom - y;

    if (dpiScale > 1.01 || dpiScale < 0.99) {
        int px = static_cast<int>(x * dpiScale + 0.5);
        int py = static_cast<int>(y * dpiScale + 0.5);
        int pw = static_cast<int>(w * dpiScale + 0.5);
        int ph = static_cast<int>(h * dpiScale + 0.5);
        StretchBlt(backDC, x, y, w, h, memDC, px, py, pw, ph, SRCCOPY);
    } else {
        BitBlt(backDC, x, y, w, h, memDC, x, y, SRCCOPY);
    }
}

static RECT InflateRectBy(const RECT& rect, int margin) {
    RECT result = {
        rect.left - margin,
        rect.top - margin,
        rect.right + margin,
        rect.bottom + margin,
    };
    return result;
}

static void DrawInfoPanel(
    HDC hdc, int panelX, int panelY, COLORREF color, HDC memDC, int vx, int vy,
    int mx, int my, double dpiScale, const SCGdiResources& gdi) {
    HGDIOBJ oldBrush = SelectObject(hdc, gdi.bgBrush);
    HGDIOBJ oldPen = SelectObject(hdc, gdi.borderPen);

    RoundRect(
        hdc, panelX, panelY, panelX + SC_PANEL_WIDTH, panelY + SC_PANEL_HEIGHT,
        SC_PANEL_CORNER_RADIUS, SC_PANEL_CORNER_RADIUS);

    int srcW = SC_PANEL_WIDTH / SC_ZOOM_FACTOR;
    int srcH = SC_MAGNIFIER_HEIGHT / SC_ZOOM_FACTOR;
    int mxLogical = mx - vx;
    int myLogical = my - vy;
    int mxPhysical = static_cast<int>(mxLogical * dpiScale + 0.5);
    int myPhysical = static_cast<int>(myLogical * dpiScale + 0.5);
    int srcWPhysical = static_cast<int>(srcW * dpiScale + 0.5);
    int srcHPhysical = static_cast<int>(srcH * dpiScale + 0.5);
    int srcXPhysical = mxPhysical - srcWPhysical / 2;
    int srcYPhysical = myPhysical - srcHPhysical / 2;

    int magX = panelX + 2;
    int magY = panelY + 2;
    int magW = SC_PANEL_WIDTH - 4;
    int magH = SC_MAGNIFIER_HEIGHT - 2;

    StretchBlt(
        hdc, magX, magY, magW, magH, memDC, (std::max)(srcXPhysical, 0),
        (std::max)(srcYPhysical, 0), srcWPhysical, srcHPhysical, SRCCOPY);

    SelectObject(hdc, gdi.crosshairPen);
    int cx = magX + magW / 2;
    int cy = magY + magH / 2;
    MoveToEx(hdc, magX, cy, NULL);
    LineTo(hdc, magX + magW, cy);
    MoveToEx(hdc, cx, magY, NULL);
    LineTo(hdc, cx, magY + magH);

    SetBkMode(hdc, TRANSPARENT);
    SetTextColor(hdc, RGB(255, 255, 255));
    HGDIOBJ oldFont = SelectObject(hdc, gdi.smallFont);

    char hexBuf[32];
    char rgbBuf[32];
    char posBuf[64];
    ColorrefToStrings(color, hexBuf, rgbBuf);
    sprintf_s(posBuf, "%d, %d", mx, my);

    const int labelPad = 6;
    int labelX = panelX + labelPad;
    int valueRightX = panelX + SC_PANEL_WIDTH - labelPad;

    SIZE textSize;
    GetTextExtentPoint32W(hdc, L"测试", 2, &textSize);
    int lineH = textSize.cy;
    int infoY = panelY + SC_PANEL_HEIGHT - labelPad - lineH * 3;

    auto drawRightAligned = [&](const wchar_t* text, int len, int rx, int ry) {
        SIZE size;
        GetTextExtentPoint32W(hdc, text, len, &size);
        TextOutW(hdc, rx - size.cx, ry, text, len);
    };

    TextOutW(hdc, labelX, infoY, L"坐标", 2);
    std::wstring posW(posBuf, posBuf + strlen(posBuf));
    drawRightAligned(posW.c_str(), static_cast<int>(posW.size()), valueRightX, infoY);

    TextOutW(hdc, labelX, infoY + lineH, L"HEX", 3);
    std::wstring hexW(hexBuf, hexBuf + strlen(hexBuf));
    drawRightAligned(hexW.c_str(), static_cast<int>(hexW.size()), valueRightX, infoY + lineH);

    TextOutW(hdc, labelX, infoY + lineH * 2, L"RGB", 3);
    std::wstring rgbW(rgbBuf, rgbBuf + strlen(rgbBuf));
    drawRightAligned(
        rgbW.c_str(), static_cast<int>(rgbW.size()), valueRightX, infoY + lineH * 2);

    SelectObject(hdc, oldFont);
    SelectObject(hdc, oldBrush);
    SelectObject(hdc, oldPen);
}

static RECT DrawSizeLabel(
    HDC hdc, int width, int height, int refLeft, int refTop, int refRight,
    int refBottom, int virtualW, int virtualH, const SCGdiResources& gdi) {
    RECT empty = {0, 0, 0, 0};
    if (width < 0 || height < 0) {
        return empty;
    }

    wchar_t sizeBuf[64];
    swprintf_s(sizeBuf, L"%d × %d", width, height);
    int sizeLen = static_cast<int>(wcslen(sizeBuf));

    HGDIOBJ oldFont = SelectObject(hdc, gdi.smallFont);
    SIZE textSize;
    GetTextExtentPoint32W(hdc, sizeBuf, sizeLen, &textSize);

    const int labelPad = 12;
    const int labelSpacing = 5;
    int labelW = textSize.cx + labelPad * 2;
    int labelH = textSize.cy + 4;

    int lx = refLeft;
    int ly = refTop - labelH - labelSpacing;
    if (ly < 0) {
        lx = refLeft + labelSpacing;
        ly = refTop + labelSpacing;
        if (lx + labelW > virtualW) {
            lx = virtualW - labelW - labelSpacing;
        }
        if (ly + labelH > virtualH) {
            ly = virtualH - labelH - labelSpacing;
        }
        if (lx + labelW > refRight) {
            lx = refRight - labelW - labelSpacing;
        }
        if (ly + labelH > refBottom) {
            ly = refBottom - labelH - labelSpacing;
        }
    }

    if (lx < 0) {
        lx = 0;
    }
    if (ly < 0) {
        ly = 0;
    }
    if (lx + labelW > virtualW) {
        lx = virtualW - labelW;
    }
    if (ly + labelH > virtualH) {
        ly = virtualH - labelH;
    }

    HGDIOBJ oldBrush = SelectObject(hdc, gdi.bgBrush);
    HGDIOBJ oldPen = SelectObject(hdc, gdi.borderPen);
    RoundRect(
        hdc, lx, ly, lx + labelW, ly + labelH, SC_PANEL_CORNER_RADIUS,
        SC_PANEL_CORNER_RADIUS);

    SetBkMode(hdc, TRANSPARENT);
    SetTextColor(hdc, RGB(255, 255, 255));
    TextOutW(hdc, lx + labelPad, ly + 2, sizeBuf, sizeLen);

    SelectObject(hdc, oldFont);
    SelectObject(hdc, oldBrush);
    SelectObject(hdc, oldPen);

    RECT result = {lx, ly, lx + labelW, ly + labelH};
    return result;
}

static RECT DrawSelection(
    HDC hdc, int x1, int y1, int x2, int y2, int vx, int vy, int vw, int vh,
    const SCGdiResources& gdi) {
    int left = (std::min)(x1, x2) - vx;
    int top = (std::min)(y1, y2) - vy;
    int right = (std::max)(x1, x2) - vx;
    int bottom = (std::max)(y1, y2) - vy;

    HGDIOBJ oldPen = SelectObject(hdc, gdi.selectionPen);
    HGDIOBJ oldBrush = SelectObject(hdc, GetStockObject(NULL_BRUSH));
    Rectangle(hdc, left, top, right, bottom);

    RECT labelRect = DrawSizeLabel(
        hdc, right - left, bottom - top, left, top, right, bottom, vw, vh, gdi);

    SelectObject(hdc, oldPen);
    SelectObject(hdc, oldBrush);
    return labelRect;
}

static void DrawWindowHighlight(
    HDC hdc, const RECT& rect, int vx, int vy, const SCGdiResources& gdi) {
    int left = rect.left - vx;
    int top = rect.top - vy;
    int right = rect.right - vx;
    int bottom = rect.bottom - vy;

    HGDIOBJ oldPen = SelectObject(hdc, gdi.highlightPen);
    HGDIOBJ oldBrush = SelectObject(hdc, GetStockObject(NULL_BRUSH));
    Rectangle(hdc, left, top, right, bottom);
    SelectObject(hdc, oldPen);
    SelectObject(hdc, oldBrush);
}

static std::string BitmapToBase64Png(HBITMAP hBitmap) {
    Gdiplus::GdiplusStartupInput startupInput;
    ULONG_PTR gdiplusToken = 0;
    Gdiplus::GdiplusStartup(&gdiplusToken, &startupInput, NULL);

    std::string result;
    Gdiplus::Bitmap* bitmap = Gdiplus::Bitmap::FromHBITMAP(hBitmap, NULL);
    if (bitmap != nullptr) {
        CLSID pngClsid;
        if (GetPngEncoderClsid(&pngClsid) >= 0) {
            IStream* stream = NULL;
            CreateStreamOnHGlobal(NULL, TRUE, &stream);
            if (stream != NULL && bitmap->Save(stream, &pngClsid, NULL) == Gdiplus::Ok) {
                HGLOBAL hMem = NULL;
                GetHGlobalFromStream(stream, &hMem);
                size_t len = GlobalSize(hMem);
                BYTE* ptr = static_cast<BYTE*>(GlobalLock(hMem));
                if (ptr != nullptr && len > 0) {
                    result = Base64Encode(ptr, len);
                }
                GlobalUnlock(hMem);
            }
            if (stream != NULL) {
                stream->Release();
            }
        }
        delete bitmap;
    }

    Gdiplus::GdiplusShutdown(gdiplusToken);
    return result;
}

static bool SaveBitmapToClipboard(HBITMAP hBitmap) {
    if (!OpenClipboard(NULL)) {
        return false;
    }

    EmptyClipboard();
    BITMAP bitmap = {};
    GetObject(hBitmap, sizeof(BITMAP), &bitmap);
    HBITMAP hCopy = static_cast<HBITMAP>(
        CopyImage(hBitmap, IMAGE_BITMAP, bitmap.bmWidth, bitmap.bmHeight, LR_COPYRETURNORG));
    HANDLE hResult = SetClipboardData(CF_BITMAP, hCopy);
    CloseClipboard();
    return hResult != NULL;
}

static HBITMAP CreateRegionBitmap(
    HDC memDC, const RECT& rect, int vx, int vy, double dpiScale) {
    int width = rect.right - rect.left;
    int height = rect.bottom - rect.top;
    if (width <= 0 || height <= 0) {
        return NULL;
    }

    int lx = rect.left - vx;
    int ly = rect.top - vy;
    int px = static_cast<int>(lx * dpiScale + 0.5);
    int py = static_cast<int>(ly * dpiScale + 0.5);
    int pw = static_cast<int>(width * dpiScale + 0.5);
    int ph = static_cast<int>(height * dpiScale + 0.5);

    HDC screenDC = GetDC(NULL);
    if (!screenDC) {
        return NULL;
    }

    HDC regionDC = CreateCompatibleDC(screenDC);
    if (!regionDC) {
        ReleaseDC(NULL, screenDC);
        return NULL;
    }

    HBITMAP regionBmp = CreateCompatibleBitmap(screenDC, pw, ph);
    if (!regionBmp) {
        DeleteDC(regionDC);
        ReleaseDC(NULL, screenDC);
        return NULL;
    }

    SelectObject(regionDC, regionBmp);
    BitBlt(regionDC, 0, 0, pw, ph, memDC, px, py, SRCCOPY);

    HBITMAP finalBmp = regionBmp;
    if (dpiScale > 1.01 || dpiScale < 0.99) {
        HDC scaledDC = CreateCompatibleDC(screenDC);
        HBITMAP scaledBmp = CreateCompatibleBitmap(screenDC, width, height);
        if (!scaledDC || !scaledBmp) {
            if (scaledDC) {
                DeleteDC(scaledDC);
            }
            DeleteDC(regionDC);
            DeleteObject(regionBmp);
            ReleaseDC(NULL, screenDC);
            return NULL;
        }

        SelectObject(scaledDC, scaledBmp);
        SetStretchBltMode(scaledDC, HALFTONE);
        SetBrushOrgEx(scaledDC, 0, 0, NULL);
        StretchBlt(scaledDC, 0, 0, width, height, regionDC, 0, 0, pw, ph, SRCCOPY);
        DeleteDC(regionDC);
        DeleteObject(regionBmp);
        DeleteDC(scaledDC);
        finalBmp = scaledBmp;
    } else {
        DeleteDC(regionDC);
    }

    ReleaseDC(NULL, screenDC);
    return finalBmp;
}

static LRESULT CALLBACK ScreenshotOverlayWndProc(
    HWND hwnd, UINT msg, WPARAM wParam, LPARAM lParam) {
    CaptureContext* ctx = g_captureCtx;
    if (ctx == nullptr) {
        return DefWindowProc(hwnd, msg, wParam, lParam);
    }

    switch (msg) {
        case WM_PAINT: {
            PAINTSTRUCT ps;
            HDC hdc = BeginPaint(hwnd, &ps);
            HDC backDC = ctx->backDC;

            int panelX = 0;
            int panelY = 0;
            CalcPanelPosition(
                ctx->mouseX, ctx->mouseY, ctx->virtualX, ctx->virtualY, ctx->virtualW,
                ctx->virtualH, panelX, panelY);

            int panelXRel = panelX - ctx->virtualX;
            int panelYRel = panelY - ctx->virtualY;

            RECT curPanelRect = {
                panelXRel,
                panelYRel,
                panelXRel + SC_PANEL_WIDTH,
                panelYRel + SC_PANEL_HEIGHT,
            };

            RECT curSelRect = {0, 0, 0, 0};
            if (ctx->state == CS_Selecting) {
                curSelRect.left = (std::min)(ctx->startX, ctx->endX) - ctx->virtualX;
                curSelRect.top = (std::min)(ctx->startY, ctx->endY) - ctx->virtualY;
                curSelRect.right = (std::max)(ctx->startX, ctx->endX) - ctx->virtualX;
                curSelRect.bottom = (std::max)(ctx->startY, ctx->endY) - ctx->virtualY;
            }

            RECT curHlRect = {0, 0, 0, 0};
            if (ctx->state == CS_Idle && ctx->hoveredWindow >= 0 &&
                ctx->hoveredWindow < static_cast<int>(ctx->windows.size())) {
                const RECT& wr = ctx->windows[ctx->hoveredWindow].rect;
                curHlRect = {
                    wr.left - ctx->virtualX,
                    wr.top - ctx->virtualY,
                    wr.right - ctx->virtualX,
                    wr.bottom - ctx->virtualY,
                };
            }

            double dpiScale = ctx->dpiScale;
            int physW = static_cast<int>(ctx->virtualW * dpiScale + 0.5);
            int physH = static_cast<int>(ctx->virtualH * dpiScale + 0.5);

            if (ctx->needFullRedraw) {
                if (dpiScale > 1.01 || dpiScale < 0.99) {
                    StretchBlt(
                        backDC, 0, 0, ctx->virtualW, ctx->virtualH, ctx->memDC, 0, 0,
                        physW, physH, SRCCOPY);
                } else {
                    BitBlt(
                        backDC, 0, 0, ctx->virtualW, ctx->virtualH, ctx->memDC, 0, 0,
                        SRCCOPY);
                }
                ctx->needFullRedraw = false;
            } else {
                RestoreDirtyRegion(
                    backDC, ctx->memDC, InflateRectBy(ctx->lastPanelRect, 2), dpiScale);
                if (ctx->lastSelectionRect.right > ctx->lastSelectionRect.left) {
                    RestoreDirtyRegion(
                        backDC, ctx->memDC, InflateRectBy(ctx->lastSelectionRect, 5),
                        dpiScale);
                }
                if (ctx->lastLabelRect.right > ctx->lastLabelRect.left) {
                    RestoreDirtyRegion(
                        backDC, ctx->memDC, InflateRectBy(ctx->lastLabelRect, 2),
                        dpiScale);
                }
                if (ctx->lastHighlightRect.right > ctx->lastHighlightRect.left) {
                    RestoreDirtyRegion(
                        backDC, ctx->memDC, InflateRectBy(ctx->lastHighlightRect, 5),
                        dpiScale);
                }
            }

            if (ctx->state == CS_Idle) {
                if (ctx->hoveredWindow >= 0 &&
                    ctx->hoveredWindow < static_cast<int>(ctx->windows.size())) {
                    DrawWindowHighlight(
                        backDC, ctx->windows[ctx->hoveredWindow].rect, ctx->virtualX,
                        ctx->virtualY, ctx->gdi);
                } else {
                    POINT pt = {ctx->mouseX, ctx->mouseY};
                    HMONITOR hMonitor = MonitorFromPoint(pt, MONITOR_DEFAULTTONEAREST);
                    if (hMonitor) {
                        MONITORINFO monitorInfo;
                        monitorInfo.cbSize = sizeof(MONITORINFO);
                        if (GetMonitorInfo(hMonitor, &monitorInfo)) {
                            DrawWindowHighlight(
                                backDC, monitorInfo.rcMonitor, ctx->virtualX,
                                ctx->virtualY, ctx->gdi);
                        }
                    }
                }
            }

            RECT curLabelRect = {0, 0, 0, 0};
            if (ctx->state == CS_Selecting) {
                curLabelRect = DrawSelection(
                    backDC, ctx->startX, ctx->startY, ctx->endX, ctx->endY,
                    ctx->virtualX, ctx->virtualY, ctx->virtualW, ctx->virtualH,
                    ctx->gdi);
            } else if (ctx->state == CS_Idle) {
                RECT screenRect = {0, 0, 0, 0};
                int ww = 0;
                int wh = 0;

                if (ctx->hoveredWindow >= 0 &&
                    ctx->hoveredWindow < static_cast<int>(ctx->windows.size())) {
                    const RECT& wr = ctx->windows[ctx->hoveredWindow].rect;
                    ww = wr.right - wr.left;
                    wh = wr.bottom - wr.top;
                    screenRect = wr;
                } else {
                    POINT pt = {ctx->mouseX, ctx->mouseY};
                    HMONITOR hMonitor = MonitorFromPoint(pt, MONITOR_DEFAULTTONEAREST);
                    if (hMonitor) {
                        MONITORINFO monitorInfo;
                        monitorInfo.cbSize = sizeof(MONITORINFO);
                        if (GetMonitorInfo(hMonitor, &monitorInfo)) {
                            screenRect = monitorInfo.rcMonitor;
                            ww = screenRect.right - screenRect.left;
                            wh = screenRect.bottom - screenRect.top;
                        }
                    }
                }

                curLabelRect = DrawSizeLabel(
                    backDC, ww, wh, screenRect.left - ctx->virtualX,
                    screenRect.top - ctx->virtualY, screenRect.right - ctx->virtualX,
                    screenRect.bottom - ctx->virtualY, ctx->virtualW, ctx->virtualH,
                    ctx->gdi);
            }

            DrawInfoPanel(
                backDC, panelXRel, panelYRel, ctx->currentColor, ctx->memDC,
                ctx->virtualX, ctx->virtualY, ctx->mouseX, ctx->mouseY,
                ctx->dpiScale, ctx->gdi);

            ctx->lastPanelRect = curPanelRect;
            ctx->lastSelectionRect = curSelRect;
            ctx->lastLabelRect = curLabelRect;
            ctx->lastHighlightRect = curHlRect;

            BitBlt(hdc, 0, 0, ctx->virtualW, ctx->virtualH, backDC, 0, 0, SRCCOPY);
            EndPaint(hwnd, &ps);
            return 0;
        }

        case WM_LBUTTONDOWN: {
            if (ctx->state == CS_Idle) {
                ctx->startX = ctx->mouseX;
                ctx->startY = ctx->mouseY;
                ctx->endX = ctx->mouseX;
                ctx->endY = ctx->mouseY;
                ctx->state = CS_Selecting;
                ctx->needFullRedraw = true;
            }
            return 0;
        }

        case WM_MOUSEMOVE: {
            POINT pt;
            GetCursorPos(&pt);
            if (pt.x != ctx->mouseX || pt.y != ctx->mouseY) {
                ctx->mouseX = pt.x;
                ctx->mouseY = pt.y;
                ctx->currentColor = GetPixelColorFromBitmap(
                    ctx->memDC, ctx->mouseX, ctx->mouseY, ctx->virtualX,
                    ctx->virtualY, ctx->dpiScale);

                if (ctx->state == CS_Selecting) {
                    ctx->endX = ctx->mouseX;
                    ctx->endY = ctx->mouseY;
                } else if (ctx->state == CS_Idle) {
                    int newHovered =
                        FindWindowAtPoint(ctx->windows, ctx->mouseX, ctx->mouseY);
                    if (newHovered != ctx->hoveredWindow) {
                        ctx->hoveredWindow = newHovered;
                    }
                }
                InvalidateRect(hwnd, NULL, FALSE);
            }
            return 0;
        }

        case WM_LBUTTONUP: {
            if (ctx->state == CS_Selecting) {
                int w = abs(ctx->endX - ctx->startX);
                int h = abs(ctx->endY - ctx->startY);

                RECT finalRect = {0, 0, 0, 0};
                if (w <= 1 && h <= 1) {
                    int idx =
                        FindWindowAtPoint(ctx->windows, ctx->mouseX, ctx->mouseY);
                    if (idx >= 0) {
                        finalRect = ctx->windows[idx].rect;
                    } else {
                        POINT pt = {ctx->mouseX, ctx->mouseY};
                        HMONITOR hMonitor =
                            MonitorFromPoint(pt, MONITOR_DEFAULTTONEAREST);
                        if (hMonitor) {
                            MONITORINFO monitorInfo;
                            monitorInfo.cbSize = sizeof(MONITORINFO);
                            if (GetMonitorInfo(hMonitor, &monitorInfo)) {
                                finalRect = monitorInfo.rcMonitor;
                            } else {
                                finalRect = {
                                    ctx->virtualX,
                                    ctx->virtualY,
                                    ctx->virtualX + ctx->virtualW,
                                    ctx->virtualY + ctx->virtualH,
                                };
                            }
                        } else {
                            finalRect = {
                                ctx->virtualX,
                                ctx->virtualY,
                                ctx->virtualX + ctx->virtualW,
                                ctx->virtualY + ctx->virtualH,
                            };
                        }
                    }
                } else {
                    finalRect.left = (std::min)(ctx->startX, ctx->endX);
                    finalRect.top = (std::min)(ctx->startY, ctx->endY);
                    finalRect.right = (std::max)(ctx->startX, ctx->endX);
                    finalRect.bottom = (std::max)(ctx->startY, ctx->endY);
                }

                int width = finalRect.right - finalRect.left;
                int height = finalRect.bottom - finalRect.top;
                HBITMAP captureBitmap = CreateRegionBitmap(
                    ctx->memDC, finalRect, ctx->virtualX, ctx->virtualY, ctx->dpiScale);
                if (!captureBitmap) {
                    ctx->state = CS_Done;
                    EmitError("Failed to capture screenshot region");
                    DestroyWindow(hwnd);
                    return 0;
                }

                std::string base64 = BitmapToBase64Png(captureBitmap);
                if (base64.empty()) {
                    DeleteObject(captureBitmap);
                    ctx->state = CS_Done;
                    EmitError("Failed to encode screenshot PNG");
                    DestroyWindow(hwnd);
                    return 0;
                }

                SelectionPayload selectedPayload = {
                    finalRect.left,
                    finalRect.top,
                    width,
                    height,
                    base64,
                };
                EmitSelected(selectedPayload);

                if (!SaveBitmapToClipboard(captureBitmap)) {
                    DeleteObject(captureBitmap);
                    ctx->state = CS_Done;
                    EmitError("Failed to write screenshot to clipboard");
                    DestroyWindow(hwnd);
                    return 0;
                }

                CompletePayload completePayload = {
                    finalRect.left,
                    finalRect.top,
                    finalRect.right,
                    finalRect.bottom,
                    width,
                    height,
                    "copy",
                    base64,
                };
                DeleteObject(captureBitmap);

                ctx->state = CS_Done;
                EmitComplete(completePayload);
                DestroyWindow(hwnd);
            }
            return 0;
        }

        case WM_RBUTTONDOWN:
        case WM_CLOSE: {
            ctx->state = CS_Cancelled;
            EmitCancel();
            DestroyWindow(hwnd);
            return 0;
        }

        case WM_KEYDOWN: {
            if (wParam == VK_ESCAPE) {
                ctx->state = CS_Cancelled;
                EmitCancel();
                DestroyWindow(hwnd);
            }
            return 0;
        }

        case WM_DESTROY: {
            g_overlayWindow = NULL;
            PostQuitMessage(0);
            return 0;
        }
    }

    return DefWindowProc(hwnd, msg, wParam, lParam);
}

static void CleanupCaptureResources(CaptureContext& ctx) {
    ctx.gdi.Cleanup();
    if (ctx.backDC) {
        DeleteDC(ctx.backDC);
        ctx.backDC = NULL;
    }
    if (ctx.backBitmap) {
        DeleteObject(ctx.backBitmap);
        ctx.backBitmap = NULL;
    }
    if (ctx.memDC) {
        DeleteDC(ctx.memDC);
        ctx.memDC = NULL;
    }
    if (ctx.screenBitmap) {
        DeleteObject(ctx.screenBitmap);
        ctx.screenBitmap = NULL;
    }
}

static void CaptureThread() {
    typedef DPI_AWARENESS_CONTEXT(WINAPI * SetThreadDpiAwarenessContextProc)(
        DPI_AWARENESS_CONTEXT);
    HMODULE user32 = GetModuleHandleW(L"user32.dll");
    if (user32) {
        auto setDpiProc = reinterpret_cast<SetThreadDpiAwarenessContextProc>(
            GetProcAddress(user32, "SetThreadDpiAwarenessContext"));
        if (setDpiProc) {
            setDpiProc(DPI_AWARENESS_CONTEXT_PER_MONITOR_AWARE_V2);
        }
    }

    MSG queueInitMsg;
    PeekMessage(&queueInitMsg, NULL, WM_USER, WM_USER, PM_NOREMOVE);
    g_captureThreadId = GetCurrentThreadId();

    double dpiScale = GetDpiScaleFactor();
    HDC memDC = NULL;
    HBITMAP screenBitmap = NULL;
    int vx = 0;
    int vy = 0;
    int vw = 0;
    int vh = 0;
    if (!CaptureVirtualScreen(memDC, screenBitmap, vx, vy, vw, vh, dpiScale)) {
        EmitError("Failed to capture virtual screen");
        g_captureThreadId = 0;
        g_isCapturing = false;
        g_callbacks = CaptureCallbacks{};
        g_stopRequested = false;
        g_terminalDispatched = false;
        return;
    }

    HDC backDC = NULL;
    HBITMAP backBmp = NULL;
    if (!CreateBackBuffer(backDC, backBmp, vw, vh)) {
        DeleteDC(memDC);
        DeleteObject(screenBitmap);
        EmitError("Failed to create screenshot back buffer");
        g_captureThreadId = 0;
        g_isCapturing = false;
        g_callbacks = CaptureCallbacks{};
        g_stopRequested = false;
        g_terminalDispatched = false;
        return;
    }

    std::vector<SCWindowInfo> windows = EnumWindowsForCapture();

    SCGdiResources gdi;
    gdi.Init();

    CaptureContext ctx = {};
    ctx.state = CS_Idle;
    ctx.virtualX = vx;
    ctx.virtualY = vy;
    ctx.virtualW = vw;
    ctx.virtualH = vh;
    ctx.memDC = memDC;
    ctx.screenBitmap = screenBitmap;
    ctx.backDC = backDC;
    ctx.backBitmap = backBmp;
    ctx.needFullRedraw = true;
    ctx.dpiScale = dpiScale;
    ctx.gdi = gdi;
    ctx.windows = std::move(windows);

    POINT pt;
    GetCursorPos(&pt);
    ctx.mouseX = pt.x;
    ctx.mouseY = pt.y;
    ctx.currentColor = GetPixelColorFromBitmap(memDC, pt.x, pt.y, vx, vy, dpiScale);

    g_captureCtx = &ctx;

    WNDCLASSEXW wc = {};
    wc.cbSize = sizeof(WNDCLASSEXW);
    wc.lpfnWndProc = ScreenshotOverlayWndProc;
    wc.hInstance = GetModuleHandle(NULL);
    wc.hCursor = LoadCursor(NULL, IDC_ARROW);
    wc.lpszClassName = L"ZToolsSharedRegionCaptureOverlay";

    if (!RegisterClassExW(&wc)) {
        CleanupCaptureResources(ctx);
        g_captureCtx = nullptr;
        EmitError("Failed to register screenshot overlay window class");
        g_captureThreadId = 0;
        g_isCapturing = false;
        g_callbacks = CaptureCallbacks{};
        g_stopRequested = false;
        g_terminalDispatched = false;
        return;
    }

    g_overlayWindow = CreateWindowExW(
        WS_EX_TOPMOST | WS_EX_TOOLWINDOW,
        L"ZToolsSharedRegionCaptureOverlay",
        L"Screenshot Overlay",
        WS_POPUP,
        vx,
        vy,
        vw,
        vh,
        NULL,
        NULL,
        GetModuleHandle(NULL),
        NULL);

    if (g_overlayWindow == NULL) {
        UnregisterClassW(L"ZToolsSharedRegionCaptureOverlay", GetModuleHandle(NULL));
        CleanupCaptureResources(ctx);
        g_captureCtx = nullptr;
        EmitError("Failed to create screenshot overlay window");
        g_captureThreadId = 0;
        g_isCapturing = false;
        g_callbacks = CaptureCallbacks{};
        g_stopRequested = false;
        g_terminalDispatched = false;
        return;
    }

    ShowWindow(g_overlayWindow, SW_SHOW);
    SetForegroundWindow(g_overlayWindow);

    MSG msg;
    while (true) {
        if (ctx.state == CS_Done || ctx.state == CS_Cancelled) {
            break;
        }

        if (PeekMessage(&msg, NULL, 0, 0, PM_REMOVE)) {
            if (msg.message == WM_QUIT) {
                break;
            }
            if (msg.message == WM_SC_STOP_CAPTURE) {
                if (g_overlayWindow != NULL) {
                    PostMessageW(g_overlayWindow, WM_CLOSE, 0, 0);
                } else {
                    ctx.state = CS_Cancelled;
                    EmitCancel();
                    break;
                }
                continue;
            }
            TranslateMessage(&msg);
            DispatchMessage(&msg);
        } else {
            if (g_stopRequested.load() && g_overlayWindow != NULL) {
                PostMessageW(g_overlayWindow, WM_CLOSE, 0, 0);
            }
            if ((GetAsyncKeyState(VK_ESCAPE) & 0x8000) && !g_terminalDispatched.load()) {
                ctx.state = CS_Cancelled;
                EmitCancel();
                if (g_overlayWindow != NULL) {
                    DestroyWindow(g_overlayWindow);
                }
                break;
            }
            Sleep(1);
        }
    }

    CleanupCaptureResources(ctx);
    UnregisterClassW(L"ZToolsSharedRegionCaptureOverlay", GetModuleHandle(NULL));
    g_captureCtx = nullptr;
    g_overlayWindow = NULL;
    g_captureThreadId = 0;
    g_isCapturing = false;
    g_callbacks = CaptureCallbacks{};
    g_stopRequested = false;
    g_terminalDispatched = false;
}

}  // namespace

bool StartCaptureSession(CaptureCallbacks callbacks) {
    bool expected = false;
    if (!g_isCapturing.compare_exchange_strong(expected, true)) {
        return false;
    }

    g_stopRequested = false;
    g_terminalDispatched = false;
    g_callbacks = std::move(callbacks);

    try {
        g_captureThread = std::thread(CaptureThread);
        g_captureThread.detach();
        return true;
    } catch (...) {
        g_isCapturing = false;
        g_callbacks = CaptureCallbacks{};
        g_stopRequested = false;
        g_terminalDispatched = false;
        return false;
    }
}

void StopCaptureSession() {
    if (!g_isCapturing.load()) {
        return;
    }

    g_stopRequested = true;

    if (g_overlayWindow != NULL) {
        PostMessageW(g_overlayWindow, WM_CLOSE, 0, 0);
        return;
    }

    if (g_captureThreadId != 0) {
        PostThreadMessage(g_captureThreadId, WM_SC_STOP_CAPTURE, 0, 0);
    }
}

bool IsCaptureSessionActive() {
    return g_isCapturing.load();
}

}  // namespace screenshot::windows
