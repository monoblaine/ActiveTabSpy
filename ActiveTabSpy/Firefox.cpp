#include "pch.h"
#include "common.h"

static inline bool isTabsToolbar(IUIAutomationElement* el) {
    auto automationId = getAutomationId(el);
    return
        // Firefox
        automationId == L"TabsToolbar" ||
        // Thunderbird
        automationId == L"tabs-toolbar";
}

static inline bool isTabBrowserTabs(IUIAutomationElement* el) {
    auto automationId = getAutomationId(el);
    return
        // Firefox
        automationId == L"tabbrowser-tabs" ||
        // Thunderbird
        automationId == L"tabmail-tabs";
}

class Firefox : public Inspectable {
    public:
    virtual IUIAutomationElement* findActiveTab(IUIAutomationElement* windowEl, bool isHorizontal) {
        IUIAutomationElement* el = windowEl;
        getFirstChildElement(&el, false); // MozillaCompositorWindowClass
        do {
            getNextSiblingElement(&el);
        } while (!isTabsToolbar(el));
        getFirstChildElement(&el);
        while (!isTabBrowserTabs(el)) {
            getNextSiblingElement(&el);
        }
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

extern "C" __declspec(dllexport) void Firefox_inspectActiveTab(
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

extern "C" __declspec(dllexport) void FirefoxDevTools_inspectTabOn(
    HWND hWnd, int tabNumber,
    int* pointX, int* pointY
) {
    IUIAutomationElement* el = nullptr;
    getWindowEl(hWnd, &el);
    getLastChildElement(&el, true);
    getFirstChildElement(&el);
    getFirstChildElement(&el);
    getFirstChildElement(&el);
    getChildElementAt(&el, tabNumber, true);
    int left, right, top, bottom;
    collectPointInfo(el, pointX, pointY, &left, &right, &top, &bottom);
    el->Release();
    el = nullptr;
}
