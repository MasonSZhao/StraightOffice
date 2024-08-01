// Copyright (c) 2024-present, Shuai (Mason) Zhao
// Distributed under the MIT License (http://opensource.org/licenses/MIT)

#pragma once

#include "StraightExcelDispatchDriver.h"
#include "StraightExcelWorkbook.h"

namespace STRAIGHTEXCEL {
/**
 * @brief A collection of all the Workbook objects that are currently open in the Excel application.
 */
class ExcelWorkbooks : public STRAIGHTOLE::OleDispatchDriver {
private:
public:
    friend class ExcelApplication;

    /**
     * @brief Open a workbook.
     * @param des The destination.
     * @param filePath The file path.
     */
    void open(ExcelWorkbook& des, const std::wstring& filePath);
};

}
