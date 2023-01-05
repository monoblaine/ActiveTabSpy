#include "pch.h"
#include "common.h"

class CopyQ : public Inspectable {
    public:
    virtual IUIAutomationElement* findActiveTab(IUIAutomationElement* windowEl, bool isHorizontal) {
        IUIAutomationElement* el = windowEl;
        getLastChildElement(&el, false);
        getFirstChildElement(&el);
        getFirstChildElement(&el);
        BOOL hasKeyboardFocus = false;
        do {
            el->get_CurrentHasKeyboardFocus(&hasKeyboardFocus);
            if (hasKeyboardFocus) {
                break;
            }
            getNextSiblingElement(&el);
        }
        while (el);
        if (!hasKeyboardFocus) {
            if (el) {
                el->Release();
            }
            return nullptr;
        }
        return el;
    }
};

CopyQ inspectable;

extern "C" __declspec(dllexport) void CopyQ_inspectActiveTab(
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
