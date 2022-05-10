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

extern "C" __declspec(dllexport) void inspectActiveTabOnSsms18(
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

extern "C" __declspec(dllexport) void getSsms18ResultsGridActiveColumnCoords(
    HWND hWnd, int* success, int* left, int* right, int* top, int* bottom
) {
    *success = 0;

    IUIAutomationElement* el = nullptr;
    getWindowEl(hWnd, &el);
    getFirstChildElement(&el);

    while (el) {
        if (getClassName(el).compare(L"DockRoot") == 0) {
            break;
        }

        getNextSiblingElement(&el);
    }

    getFirstChildElement(&el); // DocumentGroup
    IUIAutomationSelectionPattern* selectionPattern = nullptr;
    IUIAutomationElementArray* selections = nullptr;
    el->GetCurrentPatternAs(
        UIA_SelectionPatternId,
        __uuidof(IUIAutomationSelectionPattern),
        (void**) &selectionPattern
    );
    el->Release();
    selectionPattern->GetCurrentSelection(&selections);
    selectionPattern->Release();
    selections->GetElement(0, &el);
    selections->Release();

    getFirstChildElement(&el);  // Close (Ctrl+F4)
    getNextSiblingElement(&el); // Toggle pin status (Ctrl+P)
    getNextSiblingElement(&el); // title bar
    getNextSiblingElement(&el); // pane

    getFirstChildElement(&el);
    getFirstChildElement(&el);
    getFirstChildElement(&el);
    getFirstChildElement(&el);
    getFirstChildElement(&el);

    getNextSiblingElement(&el);
    getNextSiblingElement(&el);

    getFirstChildElement(&el);
    IUIAutomationElement* selection = nullptr;
    el->GetCurrentPatternAs(
        UIA_SelectionPatternId,
        __uuidof(IUIAutomationSelectionPattern),
        (void**) &selectionPattern
    );
    selectionPattern->GetCurrentSelection(&selections);
    selectionPattern->Release();
    selections->GetElement(0, &selection);
    selections->Release();
    auto isResultsTabActive = getElName(selection).compare(L"Results") == 0;
    selection->Release();

    if (isResultsTabActive) {
        getFirstChildElement(&el); // Results pane
        getFirstChildElement(&el); // pane
        getFirstChildElement(&el); // pane
        getFirstChildElement(&el); // Results table
        getFirstChildElement(&el); // First column

        BOOL hasKeyboardFocus = FALSE;

        while (el) {
            el->get_CurrentHasKeyboardFocus(&hasKeyboardFocus);

            if (hasKeyboardFocus) {
                break;
            }

            getNextSiblingElement(&el);
        }

        if (hasKeyboardFocus) {
            int pointX;
            int pointY;

            getFirstChildElement(&el); // header
            collectPointInfo(el, &pointX, &pointY, left, right, top, bottom);
            *success = 1;
        }
    }

    if (el) {
        el->Release();
    }
}

extern "C" __declspec(dllexport) void isSsms18ResultsTabActiveAndFocused(HWND hWnd, int* result) {
    IUIAutomationElement* windowEl = nullptr;
    IUIAutomationElement* el = nullptr;
    getWindowEl(hWnd, &windowEl);
    el = inspectable.findActiveTab(windowEl, true);
    getLastChildElement(&el); // pane
    getFirstChildElement(&el);
    getFirstChildElement(&el);
    getFirstChildElement(&el);
    getFirstChildElement(&el);
    getFirstChildElement(&el);
    getFirstChildElement(&el);
    getFirstChildElement(&el);
    getFirstChildElement(&el); // Text editor

    BOOL hasKeyboardFocus = FALSE;

    el->get_CurrentHasKeyboardFocus(&hasKeyboardFocus);
    el->Release();

    if (hasKeyboardFocus) {
        windowEl->Release();
        *result = 0;
        return;
    }

    el = windowEl;
    getFirstChildElement(&el);
    windowEl->Release();

    while (el) {
        if (getClassName(el).compare(L"DockRoot") == 0) {
            break;
        }

        getNextSiblingElement(&el);
    }

    getFirstChildElement(&el); // DocumentGroup
    IUIAutomationSelectionPattern* selectionPattern = nullptr;
    IUIAutomationElementArray* selections = nullptr;
    el->GetCurrentPatternAs(
        UIA_SelectionPatternId,
        __uuidof(IUIAutomationSelectionPattern),
        (void**) &selectionPattern
    );
    el->Release();
    selectionPattern->GetCurrentSelection(&selections);
    selectionPattern->Release();
    selections->GetElement(0, &el);
    selections->Release();

    getFirstChildElement(&el);  // Close (Ctrl+F4)
    getNextSiblingElement(&el); // Toggle pin status (Ctrl+P)
    getNextSiblingElement(&el); // title bar
    getNextSiblingElement(&el); // pane

    getFirstChildElement(&el);
    getFirstChildElement(&el);
    getFirstChildElement(&el);
    getFirstChildElement(&el);
    getFirstChildElement(&el);

    getNextSiblingElement(&el);
    getNextSiblingElement(&el);

    getFirstChildElement(&el);
    IUIAutomationElement* selection = nullptr;
    el->GetCurrentPatternAs(
        UIA_SelectionPatternId,
        __uuidof(IUIAutomationSelectionPattern),
        (void**) &selectionPattern
    );
    selectionPattern->GetCurrentSelection(&selections);
    selectionPattern->Release();
    selections->GetElement(0, &selection);
    selections->Release();
    auto isResultsTabActive = getElName(selection).compare(L"Results") == 0;
    selection->Release();
    el->Release();
    *result = isResultsTabActive ? 1 : 0;
}
