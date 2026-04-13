#pragma once

#include <functional>
#include <string>

namespace screenshot::windows {

struct SelectionPayload {
    int x;
    int y;
    int width;
    int height;
    std::string base64;
};

struct CompletePayload {
    int x;
    int y;
    int x2;
    int y2;
    int width;
    int height;
    std::string action;
    std::string base64;
};

struct CaptureCallbacks {
    std::function<void(const SelectionPayload& payload)> onSelected;
    std::function<void(const CompletePayload& payload)> onComplete;
    std::function<void()> onCancel;
    std::function<void(const std::string& message)> onError;
};

bool StartCaptureSession(CaptureCallbacks callbacks);
void StopCaptureSession();
bool IsCaptureSessionActive();

}  // namespace screenshot::windows
