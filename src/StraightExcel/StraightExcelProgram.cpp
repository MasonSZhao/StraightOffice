// Copyright (c) 2024-present, Shuai (Mason) Zhao
// Distributed under the MIT License (http://opensource.org/licenses/MIT)

#include "..\..\include\StraightExcel\StraightExcelProgram.h"
#include <sstream>

namespace STRAIGHTEXCEL {
HRESULT ExcelProgram::getCLSIDFromProgID(CLSID& des)
{
    return CLSIDFromProgID(L"Excel.Application", &des);
}

HRESULT ExcelProgram::start(ExcelApplication& des)
{
    HRESULT hr;
    IDispatch* pXlApp {};
    CLSID clsid;
    hr = getCLSIDFromProgID(clsid);
    if (FAILED(hr)) {
        return hr;
    }
    hr = CoCreateInstance(clsid, NULL, CLSCTX_LOCAL_SERVER, IID_IDispatch, (void**)&pXlApp);
    if (FAILED(hr)) {
        return hr;
    }
    des._dispatch = pXlApp;
    if (NULL != des._dispatch) {
    }
    std::stringstream ss;
    des.operator<<(ss);
    des._ostreamStr = ss.str();
    return hr;
}

HRESULT ExcelProgram::conn(ExcelApplication& des)
{
    HRESULT hr;
    IDispatch* pXlApp {};
    CLSID clsid;
    IUnknown* pUnk {};
    hr = getCLSIDFromProgID(clsid);
    if (FAILED(hr)) {
        return hr;
    }
    hr = GetActiveObject(clsid, NULL, (IUnknown**)&pUnk);
    if (FAILED(hr)) {
        return hr;
    }
    hr = pUnk->QueryInterface(IID_IDispatch, (void**)&pXlApp);
    if (FAILED(hr)) {
        return hr;
    }
    des._dispatch = pXlApp;
    if (NULL != des._dispatch) {
    }
    std::stringstream ss;
    des.operator<<(ss);
    des._ostreamStr = ss.str();
    return hr;
}

}
