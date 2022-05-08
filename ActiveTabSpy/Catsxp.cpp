#include "pch.h"
#include "common.h"

extern "C" __declspec(dllexport) void getCatsxpThreeDotBtnStatus(HWND hWnd, int* result) {
    IUIAutomationElement* el = nullptr;
    getWindowEl(hWnd, &el);
    getLastChildElement(&el);   // Catsxp
    getFirstChildElement(&el);
    getLastChildElement(&el);
    getFirstChildElement(&el);
    getFirstChildElement(&el);
    getNextSiblingElement(&el); // toolbar
    getLastChildElement(&el);

    BOOL isFocused;
    el->get_CurrentHasKeyboardFocus(&isFocused);
    el->Release();
    *result = isFocused ? 1 : 0;
}

extern "C" __declspec(dllexport) void isCatsxpAddressBarFocused(HWND hWnd, int* result) {
    IUIAutomationElement* el = nullptr;
    getWindowEl(hWnd, &el);
    getLastChildElement(&el);   // Catsxp
    getFirstChildElement(&el);
    getLastChildElement(&el);
    getFirstChildElement(&el);
    getFirstChildElement(&el);
    getNextSiblingElement(&el); // toolbar
    getFirstChildElement(&el);

    while (!isOfType(el, UIA_EditControlTypeId)) {
        getNextSiblingElement(&el);
    }

    BOOL isFocused;
    el->get_CurrentHasKeyboardFocus(&isFocused);
    el->Release();
    *result = isFocused ? 1 : 0;
}
