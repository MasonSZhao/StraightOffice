// Copyright (c) 2024-present, Shuai (Mason) Zhao
// Distributed under the MIT License (http://opensource.org/licenses/MIT)

#pragma once

#include "StraightPptDispatchDriver.h"
#include "StraightPptPresentations.h"

namespace STRAIGHTPPT {
/**
 * @brief Represents the entire Microsoft Ppt application.
 */
class PptApplication : public STRAIGHTOLE::OleDispatchDriver {
private:
public:
    friend class PptProgram;

    /**
     * @brief Returns a Presentations collection that represents all open presentations.
     * @param des The destination.
     */
    void get_Presentations(PptPresentations& des);

    /**
     * @brief Sets a Boolean value that determines whether the object is visible.
     * @param val The value to set to.
     */
    void put_Visible(bool val);
};
}
