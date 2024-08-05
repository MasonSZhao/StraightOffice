// Copyright (c) 2024-present, Shuai (Mason) Zhao
// Distributed under the MIT License (http://opensource.org/licenses/MIT)

#include "..\..\include\StraightPpt\StraightPptPresentations.h"
#include "..\..\include\StraightOle\StraightOleVariant.h"
#include "atlconv.h"
#include <iostream>
#include <vector>

namespace STRAIGHTPPT {

void PptPresentations::open(PptPresentation& des, const std::wstring& filePath)
{
    static const int version { 2007 };
    STRAIGHTOLE::OleStrW x(filePath);
    VARIANT result;
    VariantInit(&result);
    this->invokeNameId(DISPATCH_METHOD, &result, 0x7d5, 1, x);
    // this->invokeNameId(DISPATCH_METHOD, &result, const_cast<LPOLESTR>(L"Open"), 1, x);
    des._dispatch = result.pdispVal;
}
}
