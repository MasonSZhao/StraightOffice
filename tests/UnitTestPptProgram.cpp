#include "..\..\include\StraightOle\StraightOleDialogBox.h"
#include "..\..\include\StraightOle\StraightOleInit.h"
#include "..\..\include\StraightPpt\StraightPptApplication.h"
#include "..\..\include\StraightPpt\StraightPptPresentation.h"
#include "..\..\include\StraightPpt\StraightPptPresentations.h"
#include "..\..\include\StraightPpt\StraightPptProgram.h"

/*
 * @details
 *     https://github.com/catchorg/Catch2/blob/devel/docs/own-main.md
 */
#define CATCH_CONFIG_MAIN
#include "catch.hpp"

SCENARIO("UnitTestPptProgram", "[PptProgram]")
{
    GIVEN("class PptProgram")
    {
        GIVEN("function start()")
        {
            WHEN("The user selects.")
            {
                THEN("The file should be opened.")
                {
                    STRAIGHTOLE::OleInit::initComForThisThread();
                    STRAIGHTPPT::PptApplication pXlApp;
                    STRAIGHTPPT::PptProgram::start(pXlApp);
                    STRAIGHTPPT::PptPresentations pXlPresentations;
                    pXlApp.get_Presentations(pXlPresentations);

                    STRAIGHTPPT::PptPresentation pXlPresentation;

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
                    pXlPresentations.open(pXlPresentation, wstr);
                    STRAIGHTOLE::OleInit::uninitCom();
                }
            }
        }
    }
}