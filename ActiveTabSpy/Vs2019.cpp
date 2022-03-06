#include "pch.h"
#include "common.h"

const COLORREF activeTabColor = RGB(0x00, 0x7A, 0xCC);

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

class Vs2019 : public Inspectable {
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
            IUIAutomationSelectionPattern* selectionPattern = nullptr;
            IUIAutomationElementArray* selections = nullptr;
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
            auto selectionName = getElName(selection);
            selection->Release();
            getPrevSiblingElement(&el); // Unpinned tabs
            auto activeTab = findActiveTabByColorAndName(el, &selectionName);

            if (!activeTab) {
                getPrevSiblingElement(&el); // Pinned tabs
                activeTab = findActiveTabByColorAndName(el, &selectionName);
            }

            el->Release();

            if (activeTab) {
                IUIAutomationSelectionItemPattern* selectionItemPattern = nullptr;
                activeTab->GetCurrentPatternAs(
                    UIA_SelectionItemPatternId,
                    __uuidof(IUIAutomationSelectionItemPattern),
                    (void**) &selectionItemPattern
                );
                selectionItemPattern->Select();
                selectionItemPattern->Release();
            }

            return activeTab;
        }
    }
};

Vs2019 inspectable;

extern "C" __declspec(dllexport) void inspectActiveTabOnVs2019(HWND hWnd, int isHorizontal, int* pointX, int* pointY, int* left, int* right, int* top, int* bottom) {
    inspectActiveTab(hWnd, isHorizontal, pointX, pointY, left, right, top, bottom, &inspectable);
}
