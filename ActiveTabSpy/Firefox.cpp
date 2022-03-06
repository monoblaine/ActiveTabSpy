#include "pch.h"
#include "common.h"

static inline bool isTabsToolbar(IUIAutomationElement* el) {
    auto automationId = getAutomationId(el);

    return automationId == L"TabsToolbar" || automationId == L"tabs-toolbar";
}

class Firefox : public Inspectable {
    public:
    virtual IUIAutomationElement* findActiveTab(IUIAutomationElement* windowEl, bool isHorizontal) {
        IUIAutomationElement* el = windowEl;

        getFirstChildElement(&el, false); // MozillaCompositorWindowClass
        do getNextSiblingElement(&el);
        while (!isTabsToolbar(el));
        getFirstChildElement(&el); // tabbrowser-tabs or tabmail-tabs

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

Firefox inspectable;

extern "C" __declspec(dllexport) void inspectActiveTabOnFirefox(HWND hWnd, int isHorizontal, int* pointX, int* pointY, int* left, int* right, int* top, int* bottom) {
    inspectActiveTab(hWnd, isHorizontal, pointX, pointY, left, right, top, bottom, &inspectable);
}
