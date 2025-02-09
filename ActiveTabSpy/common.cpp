#include "pch.h"
#include "common.h"

static IUIAutomationTreeWalker* treeWalker = nullptr;
static IUIAutomation* uiAutomation = nullptr;
static HDC dc = GetDC(NULL);

static bool init() {
    auto hr = CoInitialize(NULL);

    if (SUCCEEDED(hr)) {
        hr = CoCreateInstance(
            __uuidof(CUIAutomation),
            NULL,
            CLSCTX_INPROC_SERVER,
            __uuidof(IUIAutomation),
            (void**) &uiAutomation
        );

        if (SUCCEEDED(hr)) {
            uiAutomation->get_ContentViewWalker(&treeWalker);
        }
    }

    return (SUCCEEDED(hr));
}

COLORREF getPixel(int x, int y) {
    BYTE* bitPointer = nullptr;
    HDC dcTmp = CreateCompatibleDC(dc);
    BITMAPINFO bitmap {};
    bitmap.bmiHeader.biSize = sizeof(bitmap.bmiHeader);
    bitmap.bmiHeader.biWidth = 1;
    bitmap.bmiHeader.biHeight = 1;
    bitmap.bmiHeader.biPlanes = 1;
    bitmap.bmiHeader.biBitCount = 24;
    bitmap.bmiHeader.biCompression = BI_RGB;
    bitmap.bmiHeader.biSizeImage = 3;
    bitmap.bmiHeader.biClrUsed = 0;
    bitmap.bmiHeader.biClrImportant = 0;
    HBITMAP hBitmap2 = CreateDIBSection(dcTmp, &bitmap, DIB_RGB_COLORS, (void**) (&bitPointer), NULL, NULL);
    SelectObject(dcTmp, hBitmap2);
    BitBlt(dcTmp, 0, 0, 1, 1, dc, x, y, SRCCOPY);
    COLORREF color = RGB(bitPointer[2], bitPointer[1], bitPointer[0]);
    DeleteObject(hBitmap2);
    DeleteDC(dcTmp);

    return color;
}

bool isOfType(IUIAutomationElement* el, CONTROLTYPEID typeId) {
    CONTROLTYPEID currentControlTypeId;
    auto hr = el->get_CurrentControlType(&currentControlTypeId);

    return SUCCEEDED(hr) && currentControlTypeId == typeId;
}

bool isClsMatch(IUIAutomationElement* el, const wchar_t* clsName) {
    BSTR currentClsName;
    auto hr = el->get_CurrentClassName(&currentClsName);
    bool isMatch;

    if (SUCCEEDED(hr)) {
        std::wstring currentClsNameWStr(currentClsName);
        isMatch = currentClsNameWStr.compare(clsName) == 0;
    }
    else {
        isMatch = false;
    }

    SysFreeString(currentClsName);

    return isMatch;
}

std::wstring getElName(IUIAutomationElement* el) {
    BSTR bstr;
    el->get_CurrentName(&bstr);
    std::wstring name(bstr);
    SysFreeString(bstr);

    return name;
}

BSTR getElBstrValue(IUIAutomationElement* el) {
    VARIANT variant;
    el->GetCurrentPropertyValue(UIA_LegacyIAccessibleValuePropertyId, &variant);
    BSTR value = SysAllocString(variant.bstrVal);
    VariantClear(&variant);
    return value;
}

std::wstring getAutomationId(IUIAutomationElement* el) {
    BSTR bstr;
    el->get_CurrentAutomationId(&bstr);
    std::wstring name(bstr);
    SysFreeString(bstr);

    return name;
}

std::wstring getClassName(IUIAutomationElement* el) {
    BSTR bstr;
    el->get_CurrentClassName(&bstr);
    std::wstring name(bstr);
    SysFreeString(bstr);

    return name;
}

static inline void updateEl(HRESULT hr, IUIAutomationElement** el, IUIAutomationElement** tmp, bool releaseOriginalEl) {
    if (SUCCEEDED(hr)) {
        if (releaseOriginalEl) {
            (*el)->Release();
        }

        *el = *tmp;
    }
    else {
        if (releaseOriginalEl) {
            (*el)->Release();
        }
        *el = nullptr;
    }
}

void getNextSiblingElement(IUIAutomationElement** el, bool releaseOriginalEl) {
    IUIAutomationElement* tmp;
    auto hr = treeWalker->GetNextSiblingElement(*el, &tmp);
    updateEl(hr, el, &tmp, releaseOriginalEl);
}

void getPrevSiblingElement(IUIAutomationElement** el, bool releaseOriginalEl) {
    IUIAutomationElement* tmp;
    auto hr = treeWalker->GetPreviousSiblingElement(*el, &tmp);
    updateEl(hr, el, &tmp, releaseOriginalEl);
}

void getFirstChildElement(IUIAutomationElement** el, bool releaseOriginalEl) {
    IUIAutomationElement* tmp;
    auto hr = treeWalker->GetFirstChildElement(*el, &tmp);
    updateEl(hr, el, &tmp, releaseOriginalEl);
}

void getChildElementAt(IUIAutomationElement** el, int index, bool releaseOriginalEl) {
    getFirstChildElement(el, false);

    for (int i = 1; i < index; i++) {
        getNextSiblingElement(el, true);
    }
}

void getParentElement(IUIAutomationElement** el, bool releaseOriginalEl) {
    IUIAutomationElement* tmp;
    auto hr = treeWalker->GetParentElement(*el, &tmp);
    updateEl(hr, el, &tmp, releaseOriginalEl);
}

void getLastChildElement(IUIAutomationElement** el, bool releaseOriginalEl) {
    IUIAutomationElement* tmp;
    auto hr = treeWalker->GetLastChildElement(*el, &tmp);
    updateEl(hr, el, &tmp, releaseOriginalEl);
}

HRESULT findFirstElementByClassName(IUIAutomationElement** el, std::wstring className, bool releaseOriginalEl) {
    VARIANT variant {};
    variant.vt = VT_BSTR;
    variant.bstrVal = SysAllocStringLen(className.c_str(), className.size());
    IUIAutomationCondition* condition = nullptr;
    auto hr = uiAutomation->CreatePropertyCondition(UIA_ClassNamePropertyId, variant, &condition);
    if (FAILED(hr)) {
        SysFreeString(variant.bstrVal);
        return hr;
    }
    IUIAutomationElement* tmp;
    hr = (*el)->FindFirst(TreeScope_Descendants, condition, &tmp);
    if (FAILED(hr)) {
        condition->Release();
        SysFreeString(variant.bstrVal);
        return hr;
    }
    condition->Release();
    SysFreeString(variant.bstrVal);
    updateEl(hr, el, &tmp, releaseOriginalEl);
    return hr;
}

HRESULT findFirstElementByAutomationId(IUIAutomationElement** el, std::wstring automationId, bool releaseOriginalEl) {
    VARIANT variant {};
    variant.vt = VT_BSTR;
    variant.bstrVal = SysAllocStringLen(automationId.c_str(), automationId.size());
    IUIAutomationCondition* condition = nullptr;
    auto hr = uiAutomation->CreatePropertyCondition(UIA_AutomationIdPropertyId, variant, &condition);
    if (FAILED(hr)) {
        SysFreeString(variant.bstrVal);
        return hr;
    }
    IUIAutomationElement* tmp;
    hr = (*el)->FindFirst(TreeScope_Descendants, condition, &tmp);
    if (FAILED(hr)) {
        condition->Release();
        SysFreeString(variant.bstrVal);
        return hr;
    }
    condition->Release();
    SysFreeString(variant.bstrVal);
    updateEl(hr, el, &tmp, releaseOriginalEl);
    return hr;
}

static bool containsState(IUIAutomationElement* el, INT state) {
    VARIANT variant;
    el->GetCurrentPropertyValue(UIA_LegacyIAccessibleStatePropertyId, &variant);
    auto currentState = variant.intVal;
    VariantClear(&variant);
    return (currentState & state) != 0;
}

bool isActiveTab(IUIAutomationElement* el) {
    return containsState(el, STATE_SYSTEM_SELECTED);
}

bool isButtonPressed(IUIAutomationElement* el) {
    return containsState(el, STATE_SYSTEM_PRESSED);
}

void collectPointInfo(IUIAutomationElement* el, int* pointX, int* pointY, int* left, int* right, int* top, int* bottom) {
    VARIANT variant;

    el->GetCurrentPropertyValue(UIA_BoundingRectanglePropertyId, &variant);
    double* rectData = nullptr;
    SafeArrayAccessData(variant.parray, (void**) &rectData);
    *left = (int) rectData[0];
    *top = (int) rectData[1];
    *right = *left + (int) rectData[2];
    *bottom = *top + (int) rectData[3];
    SafeArrayUnaccessData(variant.parray);
    VariantClear(&variant);

    auto hr = el->GetCurrentPropertyValue(UIA_ClickablePointPropertyId, &variant);

    if (SUCCEEDED(hr) && variant.vt != VT_EMPTY) {
        double* pointData = nullptr;
        SafeArrayAccessData(variant.parray, (void**) &pointData);
        *pointX = (int) pointData[0];
        *pointY = (int) pointData[1];
        SafeArrayUnaccessData(variant.parray);
    }
    else {
        *pointX = *left  + (int) ((*right - *left) / 2);
        *pointY = *top + (int) ((*bottom - *top) / 2);
    }

    VariantClear(&variant);
}

void getWindowEl(HWND hWnd, IUIAutomationElement** el) {
    if (!uiAutomation) {
        init();
    }
    uiAutomation->ElementFromHandle(hWnd, el);
}

void getFocusedElement(IUIAutomationElement** el) {
    if (!uiAutomation) {
        init();
    }
    uiAutomation->GetFocusedElement(el);
}

void inspectActiveTab(
    HWND hWnd, int isHorizontal,
    int* pointX, int* pointY,
    int* left, int* right,
    int* top, int* bottom,
    int* prevPointX, int* prevPointY,
    int* nextPointX, int* nextPointY,
    Inspectable* inspectable
) {
    IUIAutomationElement* windowEl = nullptr;
    getWindowEl(hWnd, &windowEl);
    auto activeTab = inspectable->findActiveTab(windowEl, isHorizontal != 0);
    windowEl->Release();

    if (activeTab) {
        IUIAutomationElement* parentEl = activeTab;
        getParentElement(&parentEl, false);
        IUIAutomationElement* prevTab = parentEl;
        IUIAutomationElement* nextTab = parentEl;
        VARIANT variant;
        activeTab->GetCurrentPropertyValue(UIA_PositionInSetPropertyId, &variant);
        auto positionInSet = variant.intVal;
        VariantClear(&variant);
        activeTab->GetCurrentPropertyValue(UIA_SizeOfSetPropertyId, &variant);
        auto sizeOfSet = variant.intVal;
        auto sizeOfSetFound = sizeOfSet > 0;
        VariantClear(&variant);
        if (sizeOfSetFound) {
            auto previousTabIndex = positionInSet - 1;
            auto nextTabIndex = (positionInSet + 1) % sizeOfSet;
            if (previousTabIndex == 0) {
                previousTabIndex = sizeOfSet;
            }
            if (nextTabIndex == 0) {
                nextTabIndex = sizeOfSet;
            }
            getChildElementAt(&prevTab, previousTabIndex, false);
            getChildElementAt(&nextTab, nextTabIndex, false);
        }
        collectPointInfo(activeTab, pointX, pointY, left, right, top, bottom);
        if (sizeOfSetFound) {
            int tmpLeft;
            int tmpRight;
            int tmpTop;
            int tmpBottom;
            collectPointInfo(prevTab, prevPointX, prevPointY, &tmpLeft, &tmpRight, &tmpTop, &tmpBottom);
            collectPointInfo(nextTab, nextPointX, nextPointY, &tmpLeft, &tmpRight, &tmpTop, &tmpBottom);
        }
        activeTab->Release();
        if (sizeOfSetFound) {
            prevTab->Release();
            nextTab->Release();
        }
        parentEl->Release();
    }
}

extern "C" __declspec(dllexport) void cleanup() {
    ReleaseDC(NULL, dc);

    if (uiAutomation) {
        if (treeWalker) {
            treeWalker->Release();
        }

        uiAutomation->Release();
    }

    CoUninitialize();
}
