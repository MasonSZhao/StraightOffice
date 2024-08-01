// Copyright (c) 2024-present, Shuai (Mason) Zhao
// Distributed under the MIT License (http://opensource.org/licenses/MIT)

#pragma once

#include "StraightExcelDispatchDriver.h"
#include "StraightExcelWorkbooks.h"

namespace STRAIGHTEXCEL {
/**
 * @brief Represents the entire Microsoft Excel application.
 */
class ExcelApplication : public STRAIGHTOLE::OleDispatchDriver {
private:
public:
    friend class ExcelProgram;

    /**
     * @brief Sets a Boolean value that determines whether the object is visible.
     * @param val The value to set to.
     */
    void put_Visible(bool val);

    /**
     * @brief Returns a Workbooks collection that represents all the open workbooks.
     * @param des The destination.
     */
    void get_Workbooks(ExcelWorkbooks& des);
};

}
