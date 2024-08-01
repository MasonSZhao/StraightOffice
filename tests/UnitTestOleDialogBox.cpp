#include "..\..\include\StraightOle\StraightOleDialogBox.h"
#include "..\..\include\StraightOle\StraightOleInit.h"

/*
 * @details
 *     https://github.com/catchorg/Catch2/blob/devel/docs/own-main.md
 */
#define CATCH_CONFIG_MAIN
#include "catch.hpp"

SCENARIO("UnitTestOleDialogBox", "[OleDialogBox]")
{
    GIVEN("class OleDialogBox")
    {
        GIVEN("function startFileOpenDialogBoxA()")
        {
            WHEN("A open dialog box displays and the user selects.")
            {
                THEN("The string path should be returned.")
                {
                    STRAIGHTOLE::OleInit::initComForThisThread();
                    std::cout << STRAIGHTOLE::OleDialogBox::startFileOpenDialogBoxA() << std::endl;
                    STRAIGHTOLE::OleInit::uninitCom();
                }
            }
        }
    }
}