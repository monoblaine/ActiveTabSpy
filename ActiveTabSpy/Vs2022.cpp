#include "pch.h"
#include "common.h"

const COLORREF activeTabColor = RGB(0x00, 0x6C, 0xBE);

class Vs2022 : public Inspectable {
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

        if (isHorizontal) {
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
        else {
            return nullptr; // not tested.
        }
    }
};

Vs2022 inspectable;

extern "C" __declspec(dllexport) void inspectActiveTabOnVs2022(
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
