#include "pch.h"
#include "common.h"

class Chromium : public Inspectable {
    public:
    virtual IUIAutomationElement* findActiveTab(IUIAutomationElement* windowEl, bool isHorizontal) {
        IUIAutomationElement* el = windowEl;

        getLastChildElement(&el, false); // Chromium
        getFirstChildElement(&el);
        getLastChildElement(&el);
        getFirstChildElement(&el);
        getFirstChildElement(&el);
        getFirstChildElement(&el);
        getFirstChildElement(&el); // First tab

        while (el) {
            if (isActiveTab(el)) {
                break;
            }

            getNextSiblingElement(&el);
        }

        if (!el) {
            return nullptr;
        }

        return el;
    }
};

Chromium inspectable;

extern "C" __declspec(dllexport) void Chromium_inspectActiveTab(
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

extern "C" __declspec(dllexport) int Chromium_getThreeDotBtnStatus(HWND hWnd) {
    IUIAutomationElement* el = nullptr;
    getWindowEl(hWnd, &el);
    getLastChildElement(&el); // Chromium
    getFirstChildElement(&el);
    getLastChildElement(&el);
    getFirstChildElement(&el);
    getFirstChildElement(&el);
    getNextSiblingElement(&el); // toolbar
    getLastChildElement(&el);

    BOOL isFocused;
    el->get_CurrentHasKeyboardFocus(&isFocused);
    el->Release();
    return isFocused ? 1 : 0;
}
