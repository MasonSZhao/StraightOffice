// Copyright (c) 2024-present, Shuai (Mason) Zhao
// Distributed under the MIT License (http://opensource.org/licenses/MIT)

#pragma once

#include "StraightWordDispatchDriver.h"

namespace STRAIGHTWORD {
/**
 * @brief Represents a document. The Document object is a member of the Documents collection. The Documents collection contains all the Document objects that are currently open in Word.
 */
class WordDocument : public STRAIGHTOLE::OleDispatchDriver {
private:
public:
    friend class WordDocuments;

    /**
     * @brief Closes the specified document.
     */
    void Close();
};

}
