#include "pch.h"
#include "common.h"
#include <tchar.h>

static IUIAutomationTreeWalker* treeWalker = nullptr;
static IUIAutomation* uiAutomation = nullptr;
IUIAutomationElement* toolbar = nullptr;
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
            auto hDesktop = GetDesktopWindow();
            auto hTray = FindWindowEx(hDesktop, NULL, _T("Shell_TrayWnd"), NULL);
            auto hReBar = FindWindowEx(hTray, NULL, _T("ReBarWindow32"), NULL);
            auto hTask = FindWindowEx(hReBar, NULL, _T("MSTaskSwWClass"), NULL);
            auto hToolbar = FindWindowEx(hTask, NULL, _T("MSTaskListWClass"), NULL);
            uiAutomation->ElementFromHandle(hToolbar, &toolbar);
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

BSTR getElBstrName(IUIAutomationElement* el) {
    BSTR name;
    auto hr = el->get_CurrentName(&name);
    return FAILED(hr)
        ? nullptr
        : name;
}

BSTR getElBstrValue(IUIAutomationElement* el) {
    VARIANT variant;
    BSTR value;
    auto hr = el->GetCurrentPropertyValue(UIA_LegacyIAccessibleValuePropertyId, &variant);
    if (FAILED(hr)) {
        value = nullptr;
    }
    else {
        value = SysAllocString(variant.bstrVal);
        VariantClear(&variant);
    }
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
        goto cleanup;
    }
    IUIAutomationElement* tmp;
    hr = (*el)->FindFirst(TreeScope_Descendants, condition, &tmp);
    if (FAILED(hr)) {
        goto cleanup;
    }
    updateEl(hr, el, &tmp, releaseOriginalEl);
cleanup:
    if (condition != nullptr) {
        condition->Release();
    }
    SysFreeString(variant.bstrVal);
    VariantClear(&variant);
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

HRESULT getFocusedElement(IUIAutomationElement** el) {
    if (!uiAutomation) {
        init();
    }
    return uiAutomation->GetFocusedElement(el);
}

static bool IsButtonWithPopup(IUIAutomationElement* el, int* buttonState) {
    CONTROLTYPEID typeId;
    el->get_CurrentControlType(&typeId);
    if (typeId != UIA_ButtonControlTypeId) {
        return false;
    }
    VARIANT variant;
    el->GetCurrentPropertyValue(UIA_LegacyIAccessibleStatePropertyId, &variant);
    auto hasPopup = ((*buttonState = variant.intVal) & STATE_SYSTEM_HASPOPUP) != 0;
    VariantClear(&variant);
    return hasPopup;
}

static void SendLeftClickToCoords (int x, int y) {
    const double XSCALEFACTOR = 65535.0 / (GetSystemMetrics(SM_CXSCREEN) - 1);
    const double YSCALEFACTOR = 65535.0 / (GetSystemMetrics(SM_CYSCREEN) - 1);
    POINT cursorPos;
    GetCursorPos(&cursorPos);
    double cx = cursorPos.x * XSCALEFACTOR;
    double cy = cursorPos.y * YSCALEFACTOR;
    double nx = x * XSCALEFACTOR;
    double ny = y * YSCALEFACTOR;
    INPUT input = { 0 };
    input.type = INPUT_MOUSE;
    input.mi.dx = (LONG) nx;
    input.mi.dy = (LONG) ny;
    input.mi.dwFlags = MOUSEEVENTF_MOVE | MOUSEEVENTF_ABSOLUTE | MOUSEEVENTF_LEFTDOWN | MOUSEEVENTF_LEFTUP;
    SendInput(1, &input, sizeof(INPUT));
    input.mi.dx = (LONG) cx;
    input.mi.dy = (LONG) cy;
    input.mi.dwFlags = MOUSEEVENTF_MOVE | MOUSEEVENTF_ABSOLUTE;
    SendInput(1, &input, sizeof(INPUT));
}

static void SendLeftClickToTaskBarButton (IUIAutomationElement* el) {
    if (!el) {
        return;
    }
    VARIANT variant;
    el->GetCurrentPropertyValue(UIA_BoundingRectanglePropertyId, &variant);
    double* rectData = nullptr;
    SafeArrayAccessData(variant.parray, (void**) &rectData);
    auto left = (int) rectData[0];
    auto top = (int) rectData[1];
    auto right = left + (int) rectData[2];
    auto bottom = top + (int) rectData[3];
    SafeArrayUnaccessData(variant.parray);
    VariantClear(&variant);
    auto pointX = left + (int) ((right - left) / 2);
    auto pointY = top + (int) ((bottom - top) / 2);
    SendLeftClickToCoords(pointX, pointY);
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

extern "C" __declspec(dllexport) BSTR getFocusedElValue(int* result) {
    IUIAutomationElement* el = nullptr;
    auto hr = getFocusedElement(&el);
    BSTR value;
    if (FAILED(hr)) {
        value = nullptr;
    }
    else {
        value = getElBstrValue(el);
        el->Release();
        el = nullptr;
    }
    *result = value == nullptr ? 0 : 1;
    return value;
}

extern "C" __declspec(dllexport) BSTR getFocusedElName(int* result) {
    IUIAutomationElement* el = nullptr;
    auto hr = getFocusedElement(&el);
    BSTR value;
    if (FAILED(hr)) {
        value = nullptr;
    }
    else {
        value = getElBstrName(el);
        el->Release();
        el = nullptr;
    }
    *result = value == nullptr ? 0 : 1;
    return value;
}

extern "C" __declspec(dllexport) void getFocusedElCoords(
    int* result,
    int* pointX, int* pointY,
    int* left, int* right,
    int* top, int* bottom
) {
    IUIAutomationElement* el = nullptr;
    auto hr = getFocusedElement(&el);
    if (FAILED(hr)) {
        *result = 0;
        return;
    }
    collectPointInfo(el, pointX, pointY, left, right, top, bottom);
    el->Release();
    el = nullptr;
    *result = 1;
}

extern "C" __declspec(dllexport) void rearrangeFileExplorerWindowsMruStates () {
    if (!uiAutomation) {
        init();
    }
    IUIAutomationElement* el = toolbar;
    int buttonState;
    getLastChildElement(&el, false);
    while (el) {
        if (
            getAutomationId(el) == L"Microsoft.Windows.Explorer" &&
            IsButtonWithPopup(el, &buttonState) &&
            (buttonState & STATE_SYSTEM_PRESSED) == 0
        ) {
            Sleep(50);
            SendLeftClickToTaskBarButton(el);
        }
        getPrevSiblingElement(&el);
    }
}
