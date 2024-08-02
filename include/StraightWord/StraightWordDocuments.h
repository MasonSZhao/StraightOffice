// Copyright (c) 2024-present, Shuai (Mason) Zhao
// Distributed under the MIT License (http://opensource.org/licenses/MIT)

#pragma once

#include "StraightWordDispatchDriver.h"
#include "StraightWordDocument.h"

namespace STRAIGHTWORD {
/**
 * @brief A collection of all the Document objects that are currently open in Word.
 */
class WordDocuments : public STRAIGHTOLE::OleDispatchDriver {
private:
public:
    friend class WordApplication;

    /**
     * @brief Open a document.
     * @param des Destination.
     * @param filePath The file path.
     */
    void open(WordDocument& des, const std::wstring& filePath);
};
}
