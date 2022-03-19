#include "pch.h"
#include "common.h"

const COLORREF activeTabColor = RGB(0x00, 0x6C, 0xBE);

static bool isActiveTabByColor(IUIAutomationElement* tabItem) {
    RECT rect;
    tabItem->get_CurrentBoundingRectangle(&rect);
    int x = rect.left + 4 + 2;
    int y = rect.top + 2 + 2;

    return getPixel(x, y) == activeTabColor;
}

static IUIAutomationElement* findActiveTabByColorAndName(IUIAutomationElement* tabControl, std::wstring* tabName) {
    IUIAutomationElement* tabItem = tabControl;
    getFirstChildElement(&tabItem, false);

    while (tabItem) {
        if (tabName->compare(getElName(tabItem)) == 0 && isActiveTabByColor(tabItem)) {
            return tabItem;
        }

        getNextSiblingElement(&tabItem);
    }

    return nullptr;
}

static IUIAutomationElement* getSelection(IUIAutomationElement* parent) {
    IUIAutomationSelectionPattern* selectionPattern = nullptr;
    IUIAutomationElementArray* selections = nullptr;
    IUIAutomationElement* selection = nullptr;
    parent->GetCurrentPatternAs(
        UIA_SelectionPatternId,
        __uuidof(IUIAutomationSelectionPattern),
        (void**) &selectionPattern
    );
    selectionPattern->GetCurrentSelection(&selections);
    selectionPattern->Release();

    if (selections) {
        int numOfSelections = 0;
        selections->get_Length(&numOfSelections);

        if (numOfSelections) {
            selections->GetElement(0, &selection);
        }

        selections->Release();
    }

    return selection;
}

static void unselectElement(IUIAutomationElement* el) {
    IUIAutomationSelectionItemPattern* selectionItemPattern = nullptr;
    el->GetCurrentPatternAs(
        UIA_SelectionItemPatternId,
        __uuidof(IUIAutomationSelectionItemPattern),
        (void**) &selectionItemPattern
    );
    selectionItemPattern->RemoveFromSelection();
    selectionItemPattern->Release();
}

static void selectElement(IUIAutomationElement* el) {
    IUIAutomationSelectionItemPattern* selectionItemPattern = nullptr;
    el->GetCurrentPatternAs(
        UIA_SelectionItemPatternId,
        __uuidof(IUIAutomationSelectionItemPattern),
        (void**) &selectionItemPattern
    );
    selectionItemPattern->Select();
    selectionItemPattern->Release();
}

static void unSelectPrevAll(IUIAutomationElement* el) {
    if (!el) {
        return;
    }

    if (isOfType(el, UIA_TabControlTypeId)) {
        auto selection = getSelection(el);

        if (selection) {
            unselectElement(selection);
            selection->Release();
        }
    }

    getPrevSiblingElement(&el);
    unSelectPrevAll(el);
}

static IUIAutomationElement* selectInPrevAllByName(IUIAutomationElement* el, std::wstring* selectionName) {
    if (!el) {
        return nullptr;
    }

    if (isOfType(el, UIA_TabControlTypeId)) {
        auto selection = findActiveTabByColorAndName(el, selectionName);

        if (selection) {
            selectElement(selection);
            el->Release();
            return selection;
        }
    }

    getPrevSiblingElement(&el);
    return selectInPrevAllByName(el, selectionName);
}

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
            getLastChildElement(&el); // Document group
            auto selection = getSelection(el);
            auto selectionName = getElName(selection);
            selection->Release();
            //IUIAutomationElement* documentGroup = el;
            //getPrevSiblingElement(&el, false);
            //unSelectPrevAll(el);
            auto resultEl = selectInPrevAllByName(el, &selectionName);

            return resultEl;
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
