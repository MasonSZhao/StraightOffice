#include "..\..\include\StraightOle\StraightOleVariant.h"

/*
 * @details
 *     https://github.com/catchorg/Catch2/blob/devel/docs/own-main.md
 */
#define CATCH_CONFIG_MAIN
#include "catch.hpp"

SCENARIO("UnitTestOleVariant", "[OleVariant]")
{
    GIVEN("class OleVariant and subclasses")
    {
        WHEN("Construct.")
        {
            THEN("The new instances should be constructed.")
            {
                STRAIGHTOLE::OleBool oleBool { true };
                STRAIGHTOLE::OleChar oleChar { 'a' };
                STRAIGHTOLE::OleUChar oleUChar { 'a' };
                STRAIGHTOLE::OleShrt oleShrt { SHRT_MAX };
                STRAIGHTOLE::OleUShrt oleUShrt { USHRT_MAX };
                STRAIGHTOLE::OleInt oleInt { INT_MAX };
                STRAIGHTOLE::OleUInt oleUInt { UINT_MAX };
                STRAIGHTOLE::OleLong oleLong { LONG_MAX };
                STRAIGHTOLE::OleULong oleULong { ULONG_MAX };
                STRAIGHTOLE::OleLLong oleLLong { LLONG_MAX };
                STRAIGHTOLE::OleULLong oleULLong { ULLONG_MAX };
                STRAIGHTOLE::OleFlt oleFlt { 0.f };
                STRAIGHTOLE::OleDbl oleDbl { 0. };
                STRAIGHTOLE::OleStrA oleStrA { "StrA" };
                STRAIGHTOLE::OleStrW oleStrW { L"StrW" };
            }
        }
    }
}