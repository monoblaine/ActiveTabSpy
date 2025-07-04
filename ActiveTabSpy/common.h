#pragma once

#include "pch.h"
#include "Inspectable.h"

COLORREF getPixel(int x, int y);
bool isOfType(IUIAutomationElement* el, CONTROLTYPEID typeId);
bool isClsMatch(IUIAutomationElement* el, const wchar_t* clsName);
std::wstring getElName(IUIAutomationElement* el);
BSTR getElBstrName(IUIAutomationElement* el);
BSTR getElBstrValue(IUIAutomationElement* el);
std::wstring getAutomationId(IUIAutomationElement* el);
std::wstring getClassName(IUIAutomationElement* el);
void getNextSiblingElement(IUIAutomationElement** el, bool releaseOriginalEl = true);
void getPrevSiblingElement(IUIAutomationElement** el, bool releaseOriginalEl = true);
void getFirstChildElement(IUIAutomationElement** el, bool releaseOriginalEl = true);
void getChildElementAt(IUIAutomationElement** el, int index, bool releaseOriginalEl);
void getParentElement(IUIAutomationElement** el, bool releaseOriginalEl = true);
void getLastChildElement(IUIAutomationElement** el, bool releaseOriginalEl = true);
HRESULT findFirstElementByClassName(IUIAutomationElement** el, std::wstring className, bool releaseOriginalEl = true);
HRESULT findFirstElementByAutomationId(IUIAutomationElement** el, std::wstring automationId, bool releaseOriginalEl = true);
bool isActiveTab(IUIAutomationElement* el);
bool isButtonPressed(IUIAutomationElement* el);
void collectPointInfo(IUIAutomationElement* el, int* pointX, int* pointY, int* left, int* right, int* top, int* bottom);
void getWindowEl(HWND hWnd, IUIAutomationElement** el);
HRESULT getFocusedElement(IUIAutomationElement** el);
void inspectActiveTab(
    HWND hWnd, int isHorizontal,
    int* pointX, int* pointY,
    int* left, int* right,
    int* top, int* bottom,
    int* prevPointX, int* prevPointY,
    int* nextPointX, int* nextPointY,
    Inspectable* inspectable
);
extern "C" __declspec(dllexport) void cleanup();
extern "C" __declspec(dllexport) BSTR getFocusedElValue(int* result);
extern "C" __declspec(dllexport) BSTR getFocusedElName(int* result);
extern "C" __declspec(dllexport) void getFocusedElCoords(int* result, int* pointX, int* pointY, int* left, int* right, int* top, int* bottom);
