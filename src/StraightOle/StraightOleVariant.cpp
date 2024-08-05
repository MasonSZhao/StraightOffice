// Copyright (c) 2024-present, Shuai (Mason) Zhao
// Distributed under the MIT License (http://opensource.org/licenses/MIT)

#include "..\..\include\StraightOle\StraightOleVariant.h"
#include <atlconv.h> // USES_CONVERSION, W2A
#include <vector>

namespace STRAIGHTOLE {
OleVariant::~OleVariant()
{
    VariantClear(this);
}

OleBool::OleBool(bool val)
{
    this->vt = VT_BOOL;
    this->boolVal = true == val ? -1 : 0;
}

OleChar::OleChar(char val)
{
    this->vt = VT_I1;
    this->cVal = val;
}

OleUChar::OleUChar(unsigned char val)
{
    this->vt = VT_UI1;
    this->bVal = val;
}

OleShrt::OleShrt(short val)
{
    this->vt = VT_I2;
    this->iVal = val;
}

OleUShrt::OleUShrt(unsigned short val)
{
    this->vt = VT_UI2;
    this->uiVal = val;
}

OleInt::OleInt(int val)
{
    this->vt = VT_INT;
    this->intVal = val;
}

OleUInt::OleUInt(unsigned int val)
{
    this->vt = VT_UINT;
    this->uintVal = val;
}

OleLong::OleLong(long val)
{
    this->vt = VT_I4;
    this->lVal = val;
}

OleULong::OleULong(unsigned long val)
{
    this->vt = VT_UI4;
    this->ulVal = val;
}

OleLLong::OleLLong(long long val)
{
    this->vt = VT_I8;
    this->llVal = val;
}

OleULLong::OleULLong(unsigned long long val)
{
    this->vt = VT_UI8;
    this->ullVal = val;
}

OleFlt::OleFlt(float val)
{
    this->vt = VT_R4;
    this->fltVal = val;
}

OleDbl::OleDbl(double val)
{
    this->vt = VT_R8;
    this->dblVal = val;
}

OleStrA::OleStrA(const std::string& val)
{
    this->vt = VT_BSTR;
    USES_CONVERSION;
    this->bstrVal = A2BSTR(val.c_str());
}

OleStrW::OleStrW(const std::wstring& val)
{
    this->vt = VT_BSTR;
    USES_CONVERSION;
    this->bstrVal = W2BSTR(val.c_str());
}

}
