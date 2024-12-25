#include "pch.h"
#include "common.h"

extern "C" __declspec(dllexport) int MsEdge_getThreeDotBtnStatus(HWND hWnd) {
    IUIAutomationElement* el = nullptr;
    getWindowEl(hWnd, &el);
    getFirstChildElement(&el); // Intermediate D3D Window

    while (el) {
        if (getClassName(el) == L"BrowserRootView") {
            break;
        }

        getNextSiblingElement(&el);
    }

    auto hr = findFirstElementByClassName(&el, L"BrowserAppMenuButton");
    BOOL isFocused;

    if (SUCCEEDED(hr)) {
        el->get_CurrentHasKeyboardFocus(&isFocused);
    }
    else {
        isFocused = false;
    }

    el->Release();
    return isFocused ? 1 : 0;
}
