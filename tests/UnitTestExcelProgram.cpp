#include "..\..\include\StraightExcel\StraightExcelApplication.h"
#include "..\..\include\StraightExcel\StraightExcelProgram.h"
#include "..\..\include\StraightExcel\StraightExcelWorkbook.h"
#include "..\..\include\StraightExcel\StraightExcelWorkbooks.h"
#include "..\..\include\StraightOle\StraightOleDialogBox.h"
#include "..\..\include\StraightOle\StraightOleInit.h"

/*
 * @details
 *     https://github.com/catchorg/Catch2/blob/devel/docs/own-main.md
 */
#define CATCH_CONFIG_MAIN
#include "catch.hpp"

SCENARIO("UnitTestExcelProgram", "[ExcelProgram]")
{
    GIVEN("class ExcelProgram")
    {
        GIVEN("function start()")
        {
            WHEN("The user selects.")
            {
                THEN("The file should be opened.")
                {
                    STRAIGHTOLE::OleInit::initComForThisThread();
                    STRAIGHTEXCEL::ExcelApplication pXlApp;
                    STRAIGHTEXCEL::ExcelProgram::start(pXlApp);
                    pXlApp.put_Visible(true);
                    STRAIGHTEXCEL::ExcelWorkbooks pXlWorkbooks;
                    pXlApp.get_Workbooks(pXlWorkbooks);
                    STRAIGHTEXCEL::ExcelWorkbook pXlWorkbook;
                    std::string str = STRAIGHTOLE::OleDialogBox::startFileOpenDialogBoxA();
                    std::wstring wstr;
                    if (true) {
                        int strLength = MultiByteToWideChar(CP_ACP, 0, str.c_str(), -1, NULL, 0);
                        if (strLength) {
                            std::vector<wchar_t> buf(strLength + 1);
                            buf[0] = 0;
                            strLength = MultiByteToWideChar(CP_ACP, 0, str.c_str(), -1, &buf[0], strLength + 1);
                            wstr = std::wstring { &buf[0] };
                        }
                    }
                    pXlWorkbooks.open(pXlWorkbook, wstr);
                    STRAIGHTOLE::OleInit::uninitCom();
                }
            }
        }
    }
}