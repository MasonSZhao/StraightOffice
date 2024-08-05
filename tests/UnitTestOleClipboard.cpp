#include "..\..\include\StraightOle\StraightOleClipboard.h"
#include "..\..\include\StraightOle\StraightOleInit.h"

/*
 * @details
 *     https://github.com/catchorg/Catch2/blob/devel/docs/own-main.md
 */
#define CATCH_CONFIG_MAIN
#include "catch.hpp"

SCENARIO("UnitTestOleClipboard", "[OleClipboard]")
{
    GIVEN("class OleClipboard")
    {
        WHEN("Write and read.")
        {
            THEN("The data read should match the data written.")
            {
                STRAIGHTOLE::OleInit::initComForThisThread();
                std::string src { "–¥»ÎºÙ«–∞Â°£" };      
                STRAIGHTOLE::OleClipboard::setClipboardTextA(src);
                std::string des = STRAIGHTOLE::OleClipboard::getClipboardTextA();
                STRAIGHTOLE::OleInit::uninitComForThisThread();
                CHECK(0 == src.compare(des));
            }
        }
    }
}