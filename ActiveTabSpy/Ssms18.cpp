#include "pch.h"
#include "common.h"

class Ssms18 : public Inspectable {
    public:
    virtual IUIAutomationElement* findActiveTab(IUIAutomationElement* windowEl, bool isHorizontal) {
        IUIAutomationElement* el = windowEl;
        getFirstChildElement(&el, false);

        if (!el) {
            return nullptr;
        }

        do getNextSiblingElement(&el);
        while (el && !isClsMatch(el, L"DockRoot"));

        if (!el) {
            return nullptr;
        }

        getFirstChildElement(&el); // Document group
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

Ssms18 inspectable;

extern "C" __declspec(dllexport) void inspectActiveTabOnSsms18(HWND hWnd, int isHorizontal, int* pointX, int* pointY, int* left, int* right, int* top, int* bottom) {
    inspectActiveTab(hWnd, isHorizontal, pointX, pointY, left, right, top, bottom, &inspectable);
}