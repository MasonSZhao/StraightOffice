#include "..\..\include\StraightExcel\StraightExcelApplication.h"
#include "..\..\include\StraightExcel\StraightExcelProgram.h"
#include "..\..\include\StraightExcel\StraightExcelWorkbook.h"
#include "..\..\include\StraightExcel\StraightExcelWorkbooks.h"
#include "..\..\include\StraightOle\StraightOleDialogBox.h"
#include "..\..\include\StraightOle\StraightOleInit.h"
#include <thread>

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
                    if (true) {
                        STRAIGHTOLE::OleInit::initComForThisThread();
                        STRAIGHTEXCEL::ExcelApplication pXlApp;
                        STRAIGHTEXCEL::ExcelProgram::start(pXlApp);
                        pXlApp.put_Visible(true);
                        STRAIGHTEXCEL::ExcelWorkbooks pXlWorkbooks;
                        pXlApp.get_Workbooks(pXlWorkbooks);
                        STRAIGHTEXCEL::ExcelWorkbook pXlWorkbook;
                        pXlWorkbooks.open(pXlWorkbook, STRAIGHTOLE::OleDialogBox::startFileOpenDialogBoxW());
                        STRAIGHTOLE::OleInit::uninitComForThisThread();
                    }
                }
            }
        }

        GIVEN("Multiple Excel files.")
        {
            WHEN("Opens in multiple threads.")
            {
                THEN("All files should be opened.")
                {
                    if (true) {
                        std::vector<std::thread> vecThread;
                        vecThread.emplace_back([]() {
                            STRAIGHTOLE::OleInit::initComForThisThread();
                            STRAIGHTEXCEL::ExcelApplication pXlApp;
                            STRAIGHTEXCEL::ExcelProgram::start(pXlApp);
                            pXlApp.put_Visible(true);
                            STRAIGHTEXCEL::ExcelWorkbooks pXlWorkbooks;
                            pXlApp.get_Workbooks(pXlWorkbooks);
                            if (true) {
                                STRAIGHTEXCEL::ExcelWorkbook pXlWorkbook;
                                std::wstring wstr { L"C:\\Users\\mason\\Documents\\线程1文件1_Thread1File1.xlsx" };
                                pXlWorkbooks.open(pXlWorkbook, wstr);
                            }
                            if (true) {
                                STRAIGHTEXCEL::ExcelWorkbook pXlWorkbook;
                                std::wstring wstr { L"C:\\Users\\mason\\Documents\\线程1文件2_Thread1File2.xlsx" };
                                pXlWorkbooks.open(pXlWorkbook, wstr);
                            }
                            STRAIGHTOLE::OleInit::uninitComForThisThread();
                        });
                        vecThread.emplace_back([]() {
                            STRAIGHTOLE::OleInit::initComForThisThread();
                            STRAIGHTEXCEL::ExcelApplication pXlApp;
                            STRAIGHTEXCEL::ExcelProgram::start(pXlApp);
                            pXlApp.put_Visible(true);
                            STRAIGHTEXCEL::ExcelWorkbooks pXlWorkbooks;
                            pXlApp.get_Workbooks(pXlWorkbooks);
                            if (true) {
                                STRAIGHTEXCEL::ExcelWorkbook pXlWorkbook;
                                std::wstring wstr { L"C:\\Users\\mason\\Documents\\线程2文件1_Thread2File1.xlsx" };
                                pXlWorkbooks.open(pXlWorkbook, wstr);
                            }
                            if (true) {
                                STRAIGHTEXCEL::ExcelWorkbook pXlWorkbook;
                                std::wstring wstr { L"C:\\Users\\mason\\Documents\\线程2文件2_Thread2File2.xlsx" };
                                pXlWorkbooks.open(pXlWorkbook, wstr);
                            }
                            STRAIGHTOLE::OleInit::uninitComForThisThread();
                        });
                        for (auto& e : vecThread)
                            e.join();
                    }
                }
            }
        }
    }
}