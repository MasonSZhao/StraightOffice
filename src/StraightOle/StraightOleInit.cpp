// Copyright (c) 2024-present, Shuai (Mason) Zhao
// Distributed under the MIT License (http://opensource.org/licenses/MIT)

#include "..\..\include\StraightOle\StraightOleInit.h"

namespace STRAIGHTOLE {

HRESULT OleInit::initComForAllThreads()
{
    return CoInitializeEx(NULL, COINIT_APARTMENTTHREADED | COINIT_DISABLE_OLE1DDE);
}

HRESULT OleInit::initComForThisThread()
{
    return CoInitialize(NULL);
}

void OleInit::uninitComForThisThread()
{
    CoUninitialize();
}

}
