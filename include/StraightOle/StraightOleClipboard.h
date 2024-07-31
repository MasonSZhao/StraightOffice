// Copyright (c) 2024-present, Shuai (Mason) Zhao
// Distributed under the MIT License (http://opensource.org/licenses/MIT)

#pragma once

#include <ole2.h>
#include <string>

namespace STRAIGHTOLE {
/**
 * @brief This class helps to read and write the clipboard.
 */
class OleClipboard {
private:
    /**
     * @brief Gets the Clipboard Ole data and gets the Text data in it.
     * @return
     *     - Content string if Clipboard data is Text.
     *     - Empty string if Clipboard data is empty or NOT Text.
     */
    [[deprecated("")]] static std::string getOleClipboardTextA();

public:
    /**
     * @brief Gets the Clipboard Text data.
     * @return
     *     - Content string if Clipboard data is Text.
     *     - Empty string if Clipboard data is empty or NOT Text.
     * @warning initOleForThisThread() must be called before you use this function.
     */
    static std::string getClipboardTextA();

    /**
     * @brief Sets the Clipboard Text data
     * @return Success or not.
     */
    static bool setClipboardTextA(const std::string& str);
};

}
