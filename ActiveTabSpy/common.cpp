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

static bool bstrCompare(BSTR bstr, const wchar_t* str) {
    if (!bstr) {
        return false;
    }

    auto bstrLen = SysStringLen(bstr);

    if (bstrLen < 1) {
        return false;
    }

    std::wstring wstr(str);
    const size_t maxLen = min(wstr.size(), static_cast<size_t>(bstrLen));

    return wcsncmp(wstr.c_str(), bstr, maxLen) == 0;
}

bool isClsMatch(IUIAutomationElement* el, const wchar_t* clsName) {
    BSTR currentClsName;
    auto hr = el->get_CurrentClassName(&currentClsName);
    bool isMatch;

    if (SUCCEEDED(hr)) {
        isMatch = bstrCompare(currentClsName, clsName);
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

void getLastChildElement(IUIAutomationElement** el, bool releaseOriginalEl) {
    IUIAutomationElement* tmp;
    auto hr = treeWalker->GetLastChildElement(*el, &tmp);
    updateEl(hr, el, &tmp, releaseOriginalEl);
}

bool isActiveTab(IUIAutomationElement* el) {
    VARIANT variant;
    el->GetCurrentPropertyValue(UIA_LegacyIAccessibleStatePropertyId, &variant);
    auto state = variant.intVal;
    VariantClear(&variant);

    return (state & STATE_SYSTEM_SELECTED) != 0;
}

void inspectActiveTab(HWND hWnd, int isHorizontal, int* pointX, int* pointY, int* left, int* right, int* top, int* bottom, Inspectable* inspectable) {
    if (!uiAutomation) {
        init();
    }

    IUIAutomationElement* windowEl = nullptr;
    uiAutomation->ElementFromHandle(hWnd, &windowEl);
    auto activeTab = inspectable->findActiveTab(windowEl, isHorizontal != 0);
    windowEl->Release();

    if (activeTab) {
        VARIANT variant;

        activeTab->GetCurrentPropertyValue(UIA_BoundingRectanglePropertyId, &variant);
        double* rectData = nullptr;
        SafeArrayAccessData(variant.parray, (void**) &rectData);
        *left = (int) rectData[0];
        *top = (int) rectData[1];
        *right = *left + (int) rectData[2];
        *bottom = *top + (int) rectData[3];
        SafeArrayUnaccessData(variant.parray);
        VariantClear(&variant);

        auto hr = activeTab->GetCurrentPropertyValue(UIA_ClickablePointPropertyId, &variant);

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
        activeTab->Release();
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
