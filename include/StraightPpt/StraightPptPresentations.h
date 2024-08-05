// Copyright (c) 2024-present, Shuai (Mason) Zhao
// Distributed under the MIT License (http://opensource.org/licenses/MIT)

#pragma once

#include "StraightPptDispatchDriver.h"
#include "StraightPptPresentation.h"

namespace STRAIGHTPPT {
/**
 * @brief A collection of all the Presentation objects in Microsoft PowerPoint. Each Presentation object represents a presentation that's currently open in PowerPoint.
 */
class PptPresentations : public STRAIGHTOLE::OleDispatchDriver {
private:
public:
    friend class PptApplication;

    /**
     * @brief Open a presentation.
     * @param des The destination.
     * @param filePath The file path.
     * @warning Do use the native backslash format rather than the generic forward slash format.
     */
    void open(PptPresentation& des, const std::wstring& filePath);
};
}
