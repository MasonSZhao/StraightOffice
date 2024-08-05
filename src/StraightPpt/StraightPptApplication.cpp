// Copyright (c) 2024-present, Shuai (Mason) Zhao
// Distributed under the MIT License (http://opensource.org/licenses/MIT)

#include "..\..\include\StraightPpt\StraightPptApplication.h"
#include "..\..\include\StraightOle\StraightOleVariant.h"
#include <sstream>

namespace STRAIGHTPPT {
void PptApplication::get_Presentations(PptPresentations& des)
{
    static const int version { 2007 };
    VARIANT result;
    VariantInit(&result);
    this->invokeNameId(DISPATCH_PROPERTYGET, &result, 0x7d1, 0);
    // this->invokeNameStr(DISPATCH_PROPERTYGET, &result, const_cast<LPOLESTR>(L"Presentations"), 0);
    des._dispatch = result.pdispVal;
    if (NULL != des._dispatch) {
    }
    std::stringstream ss;
    des.operator<<(ss);
    des._ostreamStr = ss.str();
}

void PptApplication::put_Visible(bool val)
{
    static const int version { 2007 };
    this->invokeNameId(DISPATCH_PROPERTYPUT, NULL, 0x7ee, 1, STRAIGHTOLE::OleBool(val));
    // this->invokeNameStr(DISPATCH_PROPERTYPUT, NULL, const_cast<LPOLESTR>(L"Visible"), 1, STRAIGHTOLE::OleBool(val));
}

}
