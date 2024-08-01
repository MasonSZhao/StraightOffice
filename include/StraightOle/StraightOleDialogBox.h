// Copyright (c) 2024-present, Shuai (Mason) Zhao
// Distributed under the MIT License (http://opensource.org/licenses/MIT)

#pragma once

#include <ole2.h>
#include <string>

namespace STRAIGHTOLE {
/**
 * @brief This class helps to read and write the clipboard.
 */
class OleDialogBox {
public:
    /**
     * @brief Display a Open dialog box to the user. If the user selects a file, return the wstring path of the file.
     * @return The wstring path of the file.
     * @details https://learn.microsoft.com/en-us/windows/win32/learnwin32/example--the-open-dialog-box
     * @warning initComForThisThread() must be called before you use this function.
     */
    static std::string startFileOpenDialogBoxA();
};

}
