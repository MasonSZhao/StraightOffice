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
                    if (true) {
                        STRAIGHTOLE::OleInit::initComForThisThread();
                        STRAIGHTWORD::WordApplication pXlApp;
                        STRAIGHTWORD::WordProgram::start(pXlApp);
                        pXlApp.put_Visible(true);
                        STRAIGHTWORD::WordDocuments pXlDocuments;
                        pXlApp.get_Documents(pXlDocuments);
                        STRAIGHTWORD::WordDocument pXlDocument;
                        pXlDocuments.open(pXlDocument, STRAIGHTOLE::OleDialogBox::startFileOpenDialogBoxW());
                        std::this_thread::sleep_for(std::chrono::seconds(3));
                        std::chrono::year_month_day ymd;
                        pXlDocument.Close();
                        STRAIGHTOLE::OleInit::uninitComForThisThread();
                    }
                }
            }
        }

        GIVEN("Multiple Word files.")
        {
            WHEN("Opens in multiple threads.")
            {
                THEN("All files should be opened.")
                {
                    if (true) {
                        std::vector<std::thread> vecThread;
                        vecThread.emplace_back([]() {
                            STRAIGHTOLE::OleInit::initComForThisThread();
                            STRAIGHTWORD::WordApplication pXlApp;
                            STRAIGHTWORD::WordProgram::start(pXlApp);
                            pXlApp.put_Visible(true);
                            STRAIGHTWORD::WordDocuments pXlDocuments;
                            pXlApp.get_Documents(pXlDocuments);
                            STRAIGHTWORD::WordDocument pXlDocument1;
                            STRAIGHTWORD::WordDocument pXlDocument2;
                            if (true) {
                                std::wstring wstr { L"C:\\Users\\mason\\Documents\\线程1文件1_Thread1File1.docx" };
                                pXlDocuments.open(pXlDocument1, wstr);
                            }
                            if (true) {
                                std::wstring wstr { L"C:\\Users\\mason\\Documents\\线程1文件2_Thread1File2.docx" };
                                pXlDocuments.open(pXlDocument2, wstr);
                            }
                            std::this_thread::sleep_for(std::chrono::seconds(6));
                            pXlDocument1.Close();
                            pXlDocument2.Close();
                            STRAIGHTOLE::OleInit::uninitComForThisThread();
                        });
                        vecThread.emplace_back([]() {
                            STRAIGHTOLE::OleInit::initComForThisThread();
                            STRAIGHTWORD::WordApplication pXlApp;
                            STRAIGHTWORD::WordProgram::start(pXlApp);
                            pXlApp.put_Visible(true);
                            STRAIGHTWORD::WordDocuments pXlDocuments;
                            pXlApp.get_Documents(pXlDocuments);
                            STRAIGHTWORD::WordDocument pXlDocument1;
                            STRAIGHTWORD::WordDocument pXlDocument2;
                            if (true) {
                                std::wstring wstr { L"C:\\Users\\mason\\Documents\\线程2文件1_Thread2File1.docx" };
                                pXlDocuments.open(pXlDocument1, wstr);
                            }
                            if (true) {
                                std::wstring wstr { L"C:\\Users\\mason\\Documents\\线程2文件2_Thread2File2.docx" };
                                pXlDocuments.open(pXlDocument2, wstr);
                            }
                            std::this_thread::sleep_for(std::chrono::seconds(6));
                            pXlDocument1.Close();
                            pXlDocument2.Close();
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