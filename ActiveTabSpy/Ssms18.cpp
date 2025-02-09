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

extern "C" __declspec(dllexport) void Ssms18_inspectActiveTab(
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

extern "C" __declspec(dllexport) int Ssms18_getResultsGridActiveColumnCoords(HWND hWnd, int* left, int* right, int* top, int* bottom) {
    IUIAutomationElement* el = nullptr;
    getWindowEl(hWnd, &el);
    getFirstChildElement(&el);

    while (el) {
        if (getClassName(el) == L"DockRoot") {
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
    auto isResultsTabActive = getElName(selection) == L"Results";
    selection->Release();

    int result = 0;

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
            result = 1;
        }
    }

    if (el) {
        el->Release();
    }

    return result;
}

/// <summary>
/// result,
/// 0: Other
/// 1: Text editor
/// 2: Results tab
/// </summary>
/// <param name="hWnd"></param>
/// <returns></returns>
extern "C" __declspec(dllexport) int Ssms18_getActiveArea(HWND hWnd) {
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
        return 1;
    }

    el = windowEl;
    getFirstChildElement(&el);
    windowEl->Release();

    while (el) {
        if (getClassName(el) == L"DockRoot") {
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
    auto isResultsTabActive = getElName(selection) == L"Results";
    selection->Release();
    el->Release();
    return isResultsTabActive ? 2 : 0;
}

extern "C" __declspec(dllexport) BSTR Ssms18_getCellContent(int* result) {
    IUIAutomationElement* el = nullptr;
    IUIAutomationElement* tmp = nullptr;
    BSTR cellContent;

    auto hr = getFocusedElement(&el);

    if (FAILED(hr) || !isOfType(el, UIA_EditControlTypeId)) {
        goto fail;
    }

    tmp = el;
    getParentElement(&tmp, false);

    if (!tmp) {
        goto fail;
    }

    getParentElement(&tmp);

    if (getElName(tmp) != L"Results") {
        goto fail;
    }

    tmp->Release();
    tmp = nullptr;
    cellContent = getElBstrValue(el);
    el->Release();
    el = nullptr;
    *result = 1;

    return cellContent;

fail:
    *result = 0;

    if (tmp) {
        tmp->Release();
        tmp = nullptr;
    }

    if (el) {
        el->Release();
        el = nullptr;
    }

    return nullptr;
}

extern "C" __declspec(dllexport) int Ssms18_getObjectExplorerNodeType() {
    IUIAutomationElement* el = nullptr;
    auto hr = getFocusedElement(&el);
    if (FAILED(hr)) {
        return 0;
    }
    getParentElement(&el);
    auto elementName = getElName(el);
    el->Release();
    el = nullptr;

    if (elementName.find(L"Tables") == 0) {
        return 1;
    }

    if (elementName.find(L"Views") == 0) {
        return 2;
    }

    if (elementName.find(L"Stored Procedures") == 0) {
        return 3;
    }

    if (elementName.find(L"Table-valued Functions") == 0) {
        return 4;
    }

    if (elementName.find(L"Scalar-valued Functions") == 0) {
        return 5;
    }

    if (elementName.find(L"Keys") == 0) {
        return 6;
    }

    if (elementName.find(L"Constraints") == 0) {
        return 7;
    }

    if (elementName.find(L"Triggers") == 0) {
        return 8;
    }

    if (elementName.find(L"Indexes") == 0) {
        return 9;
    }

    return 0;
}
