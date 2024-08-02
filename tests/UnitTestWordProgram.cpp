#include "..\..\include\StraightOle\StraightOleDialogBox.h"
#include "..\..\include\StraightOle\StraightOleInit.h"
#include "..\include\StraightWord\StraightWordApplication.h"
#include "..\include\StraightWord\StraightWordDocument.h"
#include "..\include\StraightWord\StraightWordDocuments.h"
#include "..\include\StraightWord\StraightWordProgram.h"
#include <chrono>
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
                    STRAIGHTOLE::OleInit::initComForThisThread();
                    STRAIGHTWORD::WordApplication pXlApp;
                    STRAIGHTWORD::WordProgram::start(pXlApp);
                    pXlApp.put_Visible(true);
                    STRAIGHTWORD::WordDocuments pXlDocuments;
                    pXlApp.get_Documents(pXlDocuments);
                    STRAIGHTWORD::WordDocument pXlDocument;
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
                    pXlDocuments.open(pXlDocument, wstr);
                    std::this_thread::sleep_for(std::chrono::seconds(3));
                    pXlDocument.Close();
                    STRAIGHTOLE::OleInit::uninitCom();
                }
            }
        }
    }
}