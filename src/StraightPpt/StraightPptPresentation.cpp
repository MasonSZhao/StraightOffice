// Copyright (c) 2024-present, Shuai (Mason) Zhao
// Distributed under the MIT License (http://opensource.org/licenses/MIT)

#include "..\..\include\StraightPpt\StraightPptPresentation.h"

namespace STRAIGHTPPT {
void PptPresentation::Close()
{
    static const int version { 2007 };
    this->invokeNameId(DISPATCH_METHOD, NULL, 0x7f7, 0);
    // this->invokeNameStr(DISPATCH_METHOD, NULL, /*const_cast<LPOLESTR>(L"Close")*/, 0);
}

}
