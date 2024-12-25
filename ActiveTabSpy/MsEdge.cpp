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

    getFirstChildElement(&el);  // NonClientView
    getFirstChildElement(&el);  // GlassBrowserFrameView
    getFirstChildElement(&el);  // GlassBrowserCaptionButtonContainer
    getNextSiblingElement(&el); // BrowserView
    getFirstChildElement(&el);  // TopContainerView
    getFirstChildElement(&el);  // TabStripRegionView
    getNextSiblingElement(&el); // ToolbarView
    getLastChildElement(&el);   // BrowserAppMenuButton

    BOOL isFocused;
    el->get_CurrentHasKeyboardFocus(&isFocused);
    el->Release();
    return isFocused ? 1 : 0;
}
