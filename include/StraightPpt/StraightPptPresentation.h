// Copyright (c) 2024-present, Shuai (Mason) Zhao
// Distributed under the MIT License (http://opensource.org/licenses/MIT)

#pragma once

#include "StraightPptDispatchDriver.h"

namespace STRAIGHTPPT {
/**
 * @brief Represents a Microsoft PowerPoint presentation.
 */
class PptPresentation : public STRAIGHTOLE::OleDispatchDriver {
private:
public:
    friend class PptPresentations;
    /**
     * @brief Closes the specified presentation.
     */
    void Close();
};
}
