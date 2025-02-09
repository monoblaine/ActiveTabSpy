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

extern "C" __declspec(dllexport) BSTR MsEdgeDevTools_getLastMessage(HWND hWnd) {
    IUIAutomationElement* el = nullptr;
    getWindowEl(hWnd, &el);
    auto hr = findFirstElementByAutomationId(&el, L"console-messages");
    if (FAILED(hr)) {
        if (el) {
            el->Release();
        }
        return 0;
    }
    getFirstChildElement(&el);
    getLastChildElement(&el);
    BSTR bstr = nullptr;
    while (el) {
        if (bstr) {
            SysFreeString(bstr);
        }
        el->get_CurrentName(&bstr);
        getLastChildElement(&el);
    }
    //SysFreeString(bstr);
    return bstr;
}
