#include "pch.h"
#include "common.h"

extern "C" __declspec(dllexport) void WinMerge2011_switchTab(HWND hWnd, int isNext) {
    IUIAutomationElement* tabControl = nullptr;
    getWindowEl(hWnd, &tabControl);
    getFirstChildElement(&tabControl);
    while (tabControl) {
        if (isOfType(tabControl, UIA_TabControlTypeId)) {
            break;
        }
        getNextSiblingElement(&tabControl);
    }
    IUIAutomationSelectionPattern* selectionPattern = nullptr;
    IUIAutomationElementArray* selections = nullptr;
    IUIAutomationElement* selection = nullptr;
    tabControl->GetCurrentPatternAs(
        UIA_SelectionPatternId,
        __uuidof(IUIAutomationSelectionPattern),
        (void**) &selectionPattern
    );
    selectionPattern->GetCurrentSelection(&selections);
    selectionPattern->Release();
    selections->GetElement(0, &selection);
    selections->Release();
    if (isNext) {
        getNextSiblingElement(&selection);
        if (!selection) {
            selection = tabControl;
            getFirstChildElement(&selection, false);
            if (!isOfType(selection, UIA_TabItemControlTypeId)) {
                getNextSiblingElement(&selection);
            }
        }
    }
    else {
        getPrevSiblingElement(&selection);
        if (!selection || !isOfType(selection, UIA_TabItemControlTypeId)) {
            selection = tabControl;
            getLastChildElement(&selection, false);
        }
    }
    tabControl->Release();
    if (selection) {
        IUIAutomationSelectionItemPattern* selectionItemPattern = nullptr;
        selection->GetCurrentPatternAs(
            UIA_SelectionItemPatternId,
            __uuidof(IUIAutomationSelectionItemPattern),
            (void**) &selectionItemPattern
        );
        selectionItemPattern->Select();
        selectionItemPattern->Release();
        selection->Release();
    }
}
