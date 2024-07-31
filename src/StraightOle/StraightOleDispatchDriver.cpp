// Copyright (c) 2024-present, Shuai (Mason) Zhao
// Distributed under the MIT License (http://opensource.org/licenses/MIT)

#include "..\..\include\StraightOle\StraightOleDispatchDriver.h"
#include <iostream>
#include <typeinfo>

namespace STRAIGHTOLE {
HRESULT OleDispatchDriver::invokeNameId(int autoType, VARIANT* pvResult, DISPID dispID, int cArgs...)
{
    if (NULL == this->_dispatch) {
        return E_ABORT;
    }

    va_list marker;
    va_start(marker, cArgs);
    VARIANT* pArgs = new VARIANT[cArgs + 1];
    for (int i = 0; i < cArgs; i++) {
        pArgs[i] = va_arg(marker, VARIANT);
    }

    DISPPARAMS dp = { NULL, NULL, 0, 0 };
    DISPID dispidNamed = DISPID_PROPERTYPUT;
    dp.cArgs = cArgs;
    dp.rgvarg = pArgs;
    if (autoType & DISPATCH_PROPERTYPUT) {
        dp.cNamedArgs = 1;
        dp.rgdispidNamedArgs = &dispidNamed;
    }

    HRESULT hr;
    hr = this->_dispatch->Invoke(dispID, IID_NULL, LOCALE_SYSTEM_DEFAULT, autoType, &dp, pvResult, NULL, NULL);

    va_end(marker);

    delete[] pArgs;
    return hr;
}

HRESULT OleDispatchDriver::invokeNameStr(int autoType, VARIANT* pvResult, LPOLESTR ptName, int cArgs...)
{
    if (NULL == this->_dispatch) {
        return E_ABORT;
    }

    HRESULT hr;
    DISPID dispID;
    char szName[MAX_PATH];

    WideCharToMultiByte(CP_ACP /*ANSI*/, 0, ptName, -1, szName, MAX_PATH, NULL, NULL);

    hr = this->_dispatch->GetIDsOfNames(IID_NULL, &ptName, 1, LOCALE_USER_DEFAULT, &dispID);
    if (FAILED(hr)) {
        return hr;
    }

    va_list marker;
    va_start(marker, cArgs);
    VARIANT* pArgs = new VARIANT[cArgs + 1];
    for (int i = 0; i < cArgs; i++) {
        pArgs[i] = va_arg(marker, VARIANT);
    }

    DISPPARAMS dp = { NULL, NULL, 0, 0 };
    DISPID dispidNamed = DISPID_PROPERTYPUT;
    dp.cArgs = cArgs;
    dp.rgvarg = pArgs;
    if (autoType & DISPATCH_PROPERTYPUT) {
        dp.cNamedArgs = 1;
        dp.rgdispidNamedArgs = &dispidNamed;
    }

    hr = this->_dispatch->Invoke(dispID, IID_NULL, LOCALE_SYSTEM_DEFAULT, autoType, &dp, pvResult, NULL, NULL);

    va_end(marker);

    delete[] pArgs;
    return hr;
}

OleDispatchDriver::OleDispatchDriver()
{
}

OleDispatchDriver::~OleDispatchDriver()
{
    if (true == this->_autoRelease) {
        ReleaseDispatch();
    }
}

std::ostream& OleDispatchDriver::operator<<(std::ostream& os) const
{
    return os << typeid(this).name();
}

void OleDispatchDriver::ReleaseDispatch()
{
    if (true != this->_autoRelease) {
        this->_autoRelease = true;
    }
    if (!this->_ostreamStr.empty()) {
        this->_ostreamStr.clear();
    }
    if (NULL != this->_dispatch) {
        this->_dispatch->Release();
        this->_dispatch = NULL;
    }
}

}
