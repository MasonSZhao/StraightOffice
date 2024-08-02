// Copyright (c) 2024-present, Shuai (Mason) Zhao
// Distributed under the MIT License (http://opensource.org/licenses/MIT)

#include "..\..\include\StraightExcel\StraightExcelWorkbooks.h"
#include "..\..\include\StraightOle\StraightOleVariant.h"

namespace STRAIGHTEXCEL {
void ExcelWorkbooks::open(ExcelWorkbook& des, const std::wstring& filePath)
{
    static const int version { 2007 };
    STRAIGHTOLE::OleStrW x { filePath };
    VARIANT result;
    VariantInit(&result);
    this->invokeNameId(DISPATCH_METHOD, &result, 0x783, 1, x);
    // this->invokeNameStr(DISPATCH_METHOD, &result, const_cast<LPOLESTR>(L"Open"), 1, x);
    des._dispatch = result.pdispVal;
}

}
