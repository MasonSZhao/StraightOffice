// Copyright (c) 2024-present, Shuai (Mason) Zhao
// Distributed under the MIT License (http://opensource.org/licenses/MIT)

#include "..\..\include\StraightPpt\StraightPptPresentations.h"
#include "..\..\include\StraightOle\StraightOleVariant.h"

namespace STRAIGHTPPT {
void PptPresentations::open(PptPresentation& des, const std::wstring& filePath)
{
    static const int version { 2007 };
    VARIANT x;
    x.vt = VT_BSTR;
    BSTR bstrFilePath = ::SysAllocString(filePath.c_str());
    x.bstrVal = bstrFilePath;
    STRAIGHTOLE::OleStrW ole(filePath);
    VARIANT result;
    VariantInit(&result);
    this->invokeNameId(DISPATCH_METHOD, &result, 0x7d5, 1, x);

    ::SysFreeString(bstrFilePath);
    des._dispatch = result.pdispVal;
}

}
