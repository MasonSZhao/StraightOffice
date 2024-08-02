// Copyright (c) 2024-present, Shuai (Mason) Zhao
// Distributed under the MIT License (http://opensource.org/licenses/MIT)

#pragma once

#include "StraightWordDispatchDriver.h"
#include "StraightWordDocuments.h"

namespace STRAIGHTWORD {
/**
 * @brief Represents the entire Microsoft Word application.
 */
class WordApplication : public STRAIGHTOLE::OleDispatchDriver {
private:
public:
    friend class WordProgram;

    /**
     * @brief Returns a Documents collection that represents all the open documents.
     * @param des The destination.
     */
    void get_Documents(WordDocuments& des);

    /**
     * @brief Sets a Boolean value that determines whether the object is visible.
     * @param val The value to set to.
     */
    void put_Visible(bool val);
};

}
