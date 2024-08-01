// Copyright (c) 2024-present, Shuai (Mason) Zhao
// Distributed under the MIT License (http://opensource.org/licenses/MIT)

#include "..\..\include\StraightExcel\StraightExcelApplication.h"
#include "..\..\include\StraightOle\StraightOleVariant.h"
#include <sstream>

namespace STRAIGHTEXCEL {
void ExcelApplication::put_Visible(bool val)
{
    static const int version { 2007 };
    this->invokeNameId(DISPATCH_PROPERTYPUT, NULL, 0x22e, 1, STRAIGHTOLE::OleBool(val));
    // this->invokeNameStr(DISPATCH_PROPERTYPUT, NULL, const_cast<LPOLESTR>(L"Visible"), 1, STRAIGHTOLE::OleBool(val));
}

void ExcelApplication::get_Workbooks(ExcelWorkbooks& des)
{
    static const int version { 2007 };
    VARIANT result;
    VariantInit(&result);
    this->invokeNameId(DISPATCH_PROPERTYGET, &result, 0x23c, 0);
    // this->invokeNameStr(DISPATCH_PROPERTYGET, &result, const_cast<LPOLESTR>(L"Workbooks"), 0);
    des._dispatch = result.pdispVal;
    if (NULL != des._dispatch) {
    }
    std::stringstream ss;
    des.operator<<(ss);
    des._ostreamStr = ss.str();
}

}
