// Copyright (c) 2024-present, Shuai (Mason) Zhao
// Distributed under the MIT License (http://opensource.org/licenses/MIT)

#pragma once

#include "StraightExcelDispatchDriver.h"

namespace STRAIGHTEXCEL {
/**
 * @brief Represents an Excel workbook.
 */
class ExcelWorkbook : public STRAIGHTOLE::OleDispatchDriver {
private:
public:
    friend class ExcelWorkbooks;

    /**
     * @brief Closes the object.
     */
    void Close();
};
}
