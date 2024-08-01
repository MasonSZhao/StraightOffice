// Copyright (c) 2024-present, Shuai (Mason) Zhao
// Distributed under the MIT License (http://opensource.org/licenses/MIT)

#include "..\..\include\StraightOle\StraightOleDialogBox.h"
#include <shobjidl.h> // FileOpenDialogBox
#include <vector>

namespace STRAIGHTOLE {
std::string OleDialogBox::startFileOpenDialogBoxA()
{
    std::string ret {};
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
                    int strLength = WideCharToMultiByte(CP_ACP /*ANSI*/, 0, pszFilePath, -1, NULL, 0, NULL, FALSE);
                    if (strLength) {
                        std::vector<char> buf(strLength + 1);
                        buf[0] = 0;
                        WideCharToMultiByte(CP_ACP /*ANSI*/, 0, pszFilePath, -1, &buf[0], strLength + 1, NULL, FALSE);
                        ret = std::string { &buf[0] };
                    }
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
