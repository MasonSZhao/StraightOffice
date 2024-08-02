// Copyright (c) 2024-present, Shuai (Mason) Zhao
// Distributed under the MIT License (http://opensource.org/licenses/MIT)

#include "..\..\include\StraightOle\StraightOleClipboard.h"
#include <vector>

namespace STRAIGHTOLE {
std::string OleClipboard::getClipboardTextA()
{

    std::vector<char> buf {};
    HRESULT hr;
    // https://learn.microsoft.com/en-us/windows/win32/api/winuser/nf-winuser-openclipboard
    if (OpenClipboard(NULL)) {
        // https://learn.microsoft.com/en-us/windows/win32/api/winuser/nf-winuser-getclipboarddata
        HGLOBAL hClipboardData = GetClipboardData(CF_TEXT);
        if (NULL != hClipboardData) {
            // https://learn.microsoft.com/en-us/windows/win32/api/winbase/nf-winbase-globallock
            char* pchData = (char*)GlobalLock(hClipboardData);
            // https://learn.microsoft.com/en-us/windows/win32/api/winbase/nf-winbase-globalsize
            auto nLength = GlobalSize(hClipboardData);
            buf.resize(nLength);
            if (NULL != pchData) {
                strcpy_s(&buf[0], nLength, pchData);
            }
            // https://learn.microsoft.com/en-us/windows/win32/api/winbase/nf-winbase-globalunlock
            GlobalUnlock(hClipboardData);
        }
        // https://learn.microsoft.com/en-us/windows/win32/api/winuser/nf-winuser-closeclipboard
        CloseClipboard();
    }
    if (!buf.empty()) {
        return std::string { &buf[0] };
    } else {
        return std::string {};
    }
}

bool OleClipboard::setClipboardTextA(const std::string& str)
{
    bool ret;

    if (OpenClipboard(NULL)) {
        // the text should be placed in "global" memory
        HGLOBAL hMem = GlobalAlloc(GMEM_SHARE | GMEM_MOVEABLE, (str.size() + 1) * sizeof(str[0]));
        LPTSTR ptxt = (LPTSTR)GlobalLock(hMem);
        lstrcpy(ptxt, str.c_str());
        GlobalUnlock(hMem);
        // set data in clipboard; we are no longer responsible for hMem
        ret = (BOOL)SetClipboardData(CF_TEXT, hMem);

        CloseClipboard(); // relinquish it for other windows
    }
    return ret;
}

}
