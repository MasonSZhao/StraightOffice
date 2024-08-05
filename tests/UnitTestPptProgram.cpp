#include "..\..\include\StraightOle\StraightOleDialogBox.h"
#include "..\..\include\StraightOle\StraightOleInit.h"
#include "..\..\include\StraightPpt\StraightPptApplication.h"
#include "..\..\include\StraightPpt\StraightPptPresentation.h"
#include "..\..\include\StraightPpt\StraightPptPresentations.h"
#include "..\..\include\StraightPpt\StraightPptProgram.h"
#include <thread>

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
                    if (true) {
                        STRAIGHTOLE::OleInit::initComForThisThread();
                        STRAIGHTPPT::PptApplication pXlApp;
                        STRAIGHTPPT::PptProgram::start(pXlApp);
                        pXlApp.put_Visible(true);
                        STRAIGHTPPT::PptPresentations pXlPresentations;
                        pXlApp.get_Presentations(pXlPresentations);
                        STRAIGHTPPT::PptPresentation pXlPresentation;
                        pXlPresentations.open(pXlPresentation, STRAIGHTOLE::OleDialogBox::startFileOpenDialogBoxW());
                        STRAIGHTOLE::OleInit::uninitComForThisThread();
                    }
                }
            }
        }

        GIVEN("Multiple Ppt files.")
        {
            WHEN("Opens in multiple threads.")
            {
                THEN("All files should be opened.")
                {
                    if (true) {
                        std::vector<std::thread> vecThread;
                        vecThread.emplace_back([]() {
                            STRAIGHTOLE::OleInit::initComForThisThread();
                            STRAIGHTPPT::PptApplication pXlApp;
                            STRAIGHTPPT::PptProgram::start(pXlApp);
                            pXlApp.put_Visible(true);
                            STRAIGHTPPT::PptPresentations pXlPresentations;
                            pXlApp.get_Presentations(pXlPresentations);
                            if (true) {
                                STRAIGHTPPT::PptPresentation pXlPresentation;
                                std::wstring wstr { L"C:\\Users\\mason\\Documents\\线程1文件1_Thread1File1.pptx" };
                                pXlPresentations.open(pXlPresentation, wstr);
                            }
                            if (true) {
                                STRAIGHTPPT::PptPresentation pXlPresentation;
                                std::wstring wstr { L"C:\\Users\\mason\\Documents\\线程1文件2_Thread1File2.pptx" };
                                pXlPresentations.open(pXlPresentation, wstr);
                            }
                            STRAIGHTOLE::OleInit::uninitComForThisThread();
                        });
                        vecThread.emplace_back([]() {
                            STRAIGHTOLE::OleInit::initComForThisThread();
                            STRAIGHTPPT::PptApplication pXlApp;
                            STRAIGHTPPT::PptProgram::start(pXlApp);
                            pXlApp.put_Visible(true);
                            STRAIGHTPPT::PptPresentations pXlPresentations;
                            pXlApp.get_Presentations(pXlPresentations);
                            if (true) {
                                STRAIGHTPPT::PptPresentation pXlPresentation;
                                std::wstring wstr { L"C:\\Users\\mason\\Documents\\线程2文件1_Thread2File1.pptx" };
                                pXlPresentations.open(pXlPresentation, wstr);
                            }
                            if (true) {
                                STRAIGHTPPT::PptPresentation pXlPresentation;
                                std::wstring wstr { L"C:\\Users\\mason\\Documents\\线程2文件2_Thread2File2.pptx" };
                                pXlPresentations.open(pXlPresentation, wstr);
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