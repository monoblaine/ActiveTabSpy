#include "pch.h"
#include "common.h"

class MsEdge : public Inspectable {
    public:
    virtual IUIAutomationElement* findActiveTab(IUIAutomationElement* windowEl, bool isHorizontal) {
        IUIAutomationElement* el = windowEl;

        getFirstChildElement(&el, false); // Intermediate D3D Window

        while (el) {
            if (getClassName(el) == L"BrowserRootView") {
                break;
            }

            getNextSiblingElement(&el);
        }

        if (!el) {
            return nullptr;
        }

        getFirstChildElement(&el);        // NonClientView
        getFirstChildElement(&el);        // GlassBrowserFrameView
        getFirstChildElement(&el);        // GlassBrowserCaptionButtonContainer
        getNextSiblingElement(&el);       // BrowserView
        getFirstChildElement(&el);        // TopContainerView
        getFirstChildElement(&el);        // TabStripRegionView

        IUIAutomationSelectionPattern* selectionPattern = nullptr;
        IUIAutomationElementArray* selections = nullptr;
        IUIAutomationElement* selection = nullptr;
        el->GetCurrentPatternAs(
            UIA_SelectionPatternId,
            __uuidof(IUIAutomationSelectionPattern),
            (void**) &selectionPattern
        );
        el->Release();
        selectionPattern->GetCurrentSelection(&selections);
        selectionPattern->Release();
        selections->GetElement(0, &selection);
        selections->Release();

        return selection;
    }
};

MsEdge inspectable;

extern "C" __declspec(dllexport) void MsEdge_inspectActiveTab(
    HWND hWnd, int isHorizontal,
    int* pointX, int* pointY,
    int* left, int* right,
    int* top, int* bottom,
    int* prevPointX, int* prevPointY,
    int* nextPointX, int* nextPointY
) {
    inspectActiveTab(
        hWnd, isHorizontal,
        pointX, pointY,
        left, right,
        top, bottom,
        prevPointX, prevPointY,
        nextPointX, nextPointY,
        &inspectable
    );
}

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
