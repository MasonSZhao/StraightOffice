// Copyright (c) 2024-present, Shuai (Mason) Zhao
// Distributed under the MIT License (http://opensource.org/licenses/MIT)

#include "..\..\include\StraightWord\StraightWordDocuments.h"
#include "..\..\include\StraightOle\StraightOleVariant.h"

namespace STRAIGHTWORD {
void WordDocuments::open(WordDocument& des, const std::wstring& filePath)
{
    static const int version { 2007 };
    VARIANT x;
    x.vt = VT_BSTR;
    BSTR bstrFilePath = ::SysAllocString(filePath.c_str());
    x.bstrVal = bstrFilePath;
    VARIANT result;
    VariantInit(&result);
    this->invokeNameId(DISPATCH_METHOD, &result, 0x14, 1, x);
    // this->invokeNameId(DISPATCH_METHOD, &result, const_cast<LPOLESTR>(L"Open"), 1, x);
    ::SysFreeString(bstrFilePath);
    des._dispatch = result.pdispVal;
}

}
