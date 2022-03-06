#pragma once

#include "pch.h"
#include "Inspectable.h"

COLORREF getPixel(int x, int y);
bool isOfType(IUIAutomationElement* el, CONTROLTYPEID typeId);
bool isClsMatch(IUIAutomationElement* el, const wchar_t* clsName);
std::wstring getElName(IUIAutomationElement* el);
std::wstring getAutomationId(IUIAutomationElement* el);
std::wstring getClassName(IUIAutomationElement* el);
void getNextSiblingElement(IUIAutomationElement** el, bool releaseOriginalEl = true);
void getPrevSiblingElement(IUIAutomationElement** el, bool releaseOriginalEl = true);
void getFirstChildElement(IUIAutomationElement** el, bool releaseOriginalEl = true);
void getLastChildElement(IUIAutomationElement** el, bool releaseOriginalEl = true);
bool isActiveTab(IUIAutomationElement* el);
void inspectActiveTab(HWND hWnd, int isHorizontal, int* pointX, int* pointY, int* left, int* right, int* top, int* bottom, Inspectable* inspectable);
extern "C" __declspec(dllexport) void cleanup();
