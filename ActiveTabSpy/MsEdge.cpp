#include "pch.h"
#include "common.h"

class MsEdge : public Inspectable {
    public:
    virtual IUIAutomationElement* findActiveTab(IUIAutomationElement* windowEl, bool isHorizontal) {
        IUIAutomationElement* el = windowEl;

        getFirstChildElement(&el, false); // Intermediate D3D Window

        while (el) {
            if (getClassName(el).compare(L"BrowserRootView") == 0) {
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

extern "C" __declspec(dllexport) void inspectActiveTabOnMsEdge(HWND hWnd, int isHorizontal, int* pointX, int* pointY, int* left, int* right, int* top, int* bottom) {
    inspectActiveTab(hWnd, isHorizontal, pointX, pointY, left, right, top, bottom, &inspectable);
}
