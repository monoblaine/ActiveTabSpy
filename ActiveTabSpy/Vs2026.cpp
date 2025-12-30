#include "pch.h"
#include "common.h"
#ifdef _DEBUG
#include <fstream>
#include <iostream>
#include <sstream>
#include <string>
#include <filesystem>
#include <codecvt>
#endif

const COLORREF activeTabColor    = RGB(0x24, 0x27, 0x2B);
const COLORREF methodImageColor1 = RGB(0x5D, 0x2E, 0x92); // .net
const COLORREF methodImageColor2 = RGB(0x5B, 0x2E, 0x91); // cpp

static bool endsWith (const std::wstring& str, const std::wstring& suffix) {
    return str.size() >= suffix.size()
        && 0 == str.compare(str.size() - suffix.size(), suffix.size(), suffix);
}

#ifdef _DEBUG
void log(std::wstringstream& ss) {
    const wchar_t* userpath = _wgetenv(L"USERPROFILE");
    if (!userpath) {
        return;
    }
    std::filesystem::path logPath = std::filesystem::path(userpath) / L"Desktop" / L"log.txt";
    std::wofstream outfile(logPath, std::ios::app);
    if (!outfile) {
        return;
    }
    // Ensure UTF-8 output
    outfile.imbue(std::locale(std::locale::empty(), new std::codecvt_utf8<wchar_t>));
    outfile << ss.str() << L"\n\n";
    ss.str(std::wstring());
}
#endif

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

class Vs2026 : public Inspectable {
    public:
    virtual IUIAutomationElement* findActiveTab(IUIAutomationElement* windowEl, bool isHorizontal) {
        IUIAutomationElement* el = windowEl;
#ifdef _DEBUG
        std::wstringstream ss(L"");
#endif
        getFirstChildElement(&el, false);
        if (!el) {
#ifdef _DEBUG
            ss << L"getFirstChildElement of windowEl failed";
            log(ss);
#endif
            return nullptr;
        }
        do getNextSiblingElement(&el);
        while (el && !isClsMatch(el, L"DockRoot"));
        if (!el) {
#ifdef _DEBUG
            ss << L"DockRoot not found!";
            log(ss);
#endif
            return nullptr;
        }
        if (isHorizontal) {
#ifdef _DEBUG
            ss << L"isHorizontal case is not implemented.";
            log(ss);
            el->Release();
#endif
        }
        else {
            getLastChildElement(&el);
            if (!el || !isClsMatch(el, L"DocumentGroup")) {
#ifdef _DEBUG
                ss << L"DocumentGroup not found!";
                log(ss);
#endif
                if (el) {
                    el->Release();
                }
                return nullptr;
            }
            auto selection = getSelection(el);
            auto selectionName = getElName(selection);
            selection->Release();
#ifdef _DEBUG
            ss << L"selectionName: " << selectionName;
            log(ss);
#endif
            getPrevSiblingElement(&el); // TabControl/UnpinnedTabs
            if (!el || !(isClsMatch(el, L"TabControl") && getAutomationId(el) == L"UnpinnedTabs")) {
#ifdef _DEBUG
                ss << L"TabControl/UnpinnedTabs not found!";
                log(ss);
#endif
                if (el) {
                    el->Release();
                }
                return nullptr;
            }
            auto activeTab = findActiveTabByColorAndName(el, &selectionName);
            el->Release();
            if (activeTab) {
                return activeTab;
            }
            else {
#ifdef _DEBUG
                ss << L"activeTab not found";
                log(ss);
#endif
            }
        }
        return nullptr;
    }
};

Vs2026 inspectable;

extern "C" __declspec(dllexport) void Vs2026_inspectActiveTab(
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

extern "C" __declspec(dllexport) int Vs2026_isTextEditorFocused(HWND hWnd) {
    IUIAutomationElement* el = nullptr;
    auto hr = getFocusedElement(&el);
    int result = 0;
    if (FAILED(hr)) {
        return result;
    }
    if (
        getElName(el) == L"Text Editor" ||
        // Possibly a bug in VS2022 prevents "Text Editor" from being the focused el ([2025-12-30 16.21] Not sure if this is still the case in VS2026)
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

extern "C" __declspec(dllexport) int Vs2026_selectedIntelliSenseItemIsAMethod(HWND hWnd, int intelliSensePopupIsEnough) {
    IUIAutomationElement* el = nullptr;
    IUIAutomationElement* menuItemOrImage = nullptr;
    bool intelliSensePopupIsOpen;
    COLORREF pixelColor;
#ifdef _DEBUG
    std::wstringstream ss(L"");
#endif
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
#ifdef _DEBUG
    ss << L"pointX: " << pointX << L"\n";
    ss << L"pointY: " << pointY << L"\n";
    ss << L"left: "   << left   << L"\n";
    ss << L"right: "  << right  << L"\n";
    ss << L"top: "    << top    << L"\n";
    ss << L"bottom: " << bottom;
    log(ss);
    for (int y = top; y <= bottom; y++) {
        for (int x = left; x <= right; x++) {
            pixelColor = getPixel(x, y);
            ss << L"(x=" << x << L" y=" << y << L")" << L"\n";
            ss << L"(r=" << GetRValue(pixelColor) << L" g=" << GetGValue(pixelColor) << L" b=" << GetBValue(pixelColor) << L")" << L"\n";
        }
    }
    log(ss);
#endif
    pixelColor = getPixel(left + 1, top + 4);
    switch (pixelColor) {
        case methodImageColor1:
        case methodImageColor2:
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

// 1: is an IntelliSense item and is a method
// 2: NOT(1) and text editor is focused
// 0: None of the above
extern "C" __declspec(dllexport) int Vs2026_getFocusedElementType(HWND hWnd) {
    if (Vs2026_selectedIntelliSenseItemIsAMethod(hWnd, 1)) {
        return 1;
    }
    if (Vs2026_isTextEditorFocused(hWnd)) {
        return 2;
    }
    return 0;
}

extern "C" __declspec(dllexport) int Vs2026_popupWinExists(HWND hWnd) {
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
