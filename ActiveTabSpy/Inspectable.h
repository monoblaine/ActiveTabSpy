#pragma once

#include "pch.h"

class Inspectable {
    public:
    virtual ~Inspectable() {};
    virtual IUIAutomationElement* findActiveTab(IUIAutomationElement* windowEl, bool isHorizontal) = 0;
};
