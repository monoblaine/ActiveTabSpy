#include "pch.h"
#include "common.h"

const COLORREF activeTabColor            = RGB(0x00, 0x6C, 0xBE);
const COLORREF methodImageColor          = RGB(0xC2, 0xB1, 0xD6);
const COLORREF internalMethodImageColor  = RGB(0x27, 0x27, 0x27);
const COLORREF extensionMethodImageColor = RGB(0xA3, 0x87, 0xC3);
const COLORREF classImageColor           = RGB(0x98, 0x6D, 0x00);

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

extern "C" __declspec(dllexport) void Vs2022_inspectActiveTab(
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

extern "C" __declspec(dllexport) int Vs2022_isTextEditorFocused(HWND hWnd) {
    IUIAutomationElement* el = nullptr;
    getFocusedElement(&el);
    int result = 0;
    if (
        getElName(el) == L"Text Editor" ||
        // Possibly a bug in VS2022 prevents "Text Editor" from being the focused el
        getAutomationId(el) == L"VisualStudioMainWindow"
    ) {
        el->Release();
        el = nullptr;
        getWindowEl(hWnd, &el);
        getFirstChildElement(&el);
        result = isOfType(el, UIA_WindowControlTypeId) && getClassName(el) == L"Popup" ? 0 : 1;
    }
    if (el) {
        el->Release();
        el = nullptr;
    }
    return result;
}

static bool endsWith (const std::wstring& str, const std::wstring& suffix) {
    return str.size() >= suffix.size()
        && 0 == str.compare(str.size() - suffix.size(), suffix.size(), suffix);
}

extern "C" __declspec(dllexport) int Vs2022_selectedIntelliSenseItemIsAMethod(HWND hWnd, int intelliSensePopupIsEnough) {
    IUIAutomationElement* el = nullptr;
    IUIAutomationElement* menuItemOrImage = nullptr;
    bool intelliSensePopupIsOpen;
    COLORREF pixelColor;
    getWindowEl(hWnd, &el);
    int result = 0;
    auto elementName = getElName(el);
    elementName.erase(elementName.find_last_not_of(L" \n\r\t") + 1);
    if (endsWith(elementName, L".js") || endsWith(elementName, L".JS")) {
        goto cleanup;
    }
    getFirstChildElement(&el);
    intelliSensePopupIsOpen = el && isOfType(el, UIA_WindowControlTypeId) && getClassName(el) == L"Popup";
    if (!intelliSensePopupIsOpen) {
        goto cleanup;
    }
    getLastChildElement(&el);
    if (!el) {
        goto cleanup;
    }
    getFirstChildElement(&el);
    if (el && isOfType(el, UIA_CustomControlTypeId)) {
        getFirstChildElement(&el);
    }
    if (!el || !isOfType(el, UIA_ListControlTypeId)) {
        goto cleanup;
    }
    menuItemOrImage = getSelection(el);
    if (!menuItemOrImage || !isOfType(menuItemOrImage, UIA_MenuItemControlTypeId)) {
        goto cleanup;
    }
    getFirstChildElement(&menuItemOrImage); // image
    if (!menuItemOrImage || !isOfType(menuItemOrImage, UIA_ImageControlTypeId)) {
        goto cleanup;
    }
    int pointX, pointY, left, right, top, bottom;
    collectPointInfo(menuItemOrImage, &pointX, &pointY, &left, &right, &top, &bottom);
    pixelColor = getPixel(pointX, pointY + 3);
    switch (pixelColor) {
        case methodImageColor:
        case internalMethodImageColor:
        case extensionMethodImageColor:
            result = 1;
            break;
        default:
            if (intelliSensePopupIsEnough) {
                result = 2;
            }
            break;
    }
cleanup:
    if (el) {
        el->Release();
        el = nullptr;
    }
    if (menuItemOrImage) {
        menuItemOrImage->Release();
        menuItemOrImage = nullptr;
    }
    return result;
}

extern "C" __declspec(dllexport) int Vs2022_popupWinExists(HWND hWnd) {
    IUIAutomationElement* el = nullptr;
    getWindowEl(hWnd, &el);
    getFirstChildElement(&el);
    int result = el && isOfType(el, UIA_WindowControlTypeId) && getClassName(el) == L"Popup";
    if (el) {
        el->Release();
        el = nullptr;
    }
    return result;
}
