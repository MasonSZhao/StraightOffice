// Copyright (c) 2024-present, Shuai (Mason) Zhao
// Distributed under the MIT License (http://opensource.org/licenses/MIT)

#include "..\..\include\StraightOle\StraightOleDialogBox.h"
#include <atlconv.h> // USES_CONVERSION, W2A
#include <shobjidl.h> // FileOpenDialogBox
#include <vector>

namespace STRAIGHTOLE {
std::string OleDialogBox::startFileOpenDialogBoxA()
{
    USES_CONVERSION;
    return W2A(OleDialogBox::startFileOpenDialogBoxW().c_str());
}

std::wstring OleDialogBox::startFileOpenDialogBoxW()
{
    std::wstring ret {};
    HRESULT hr;
    IFileOpenDialog* pFileOpen;
    // https://learn.microsoft.com/en-us/windows/win32/api/combaseapi/nf-combaseapi-cocreateinstance
    hr = CoCreateInstance(CLSID_FileOpenDialog, NULL, CLSCTX_ALL, IID_IFileOpenDialog, reinterpret_cast<void**>(&pFileOpen));
    if (SUCCEEDED(hr)) {
        hr = pFileOpen->Show(NULL); // Show the Open Dialog.
        if (SUCCEEDED(hr)) {
            IShellItem* pItem;
            hr = pFileOpen->GetResult(&pItem);
            if (SUCCEEDED(hr)) {
                PWSTR /*Unicode UTF-16 WideChar*/ pszFilePath;
                hr = pItem->GetDisplayName(SIGDN_FILESYSPATH, &pszFilePath);
                if (SUCCEEDED(hr)) {
                    ret = std::wstring { pszFilePath };
                    CoTaskMemFree(pszFilePath);
                }
                pItem->Release();
            }
        }
        pFileOpen->Release();
    }
    return ret;
}

}
