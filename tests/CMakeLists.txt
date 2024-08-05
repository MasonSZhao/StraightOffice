# Set the C++ standard.
set(CMAKE_CXX_STANDARD 20)

# Add source to the test's executable test.
add_executable(UnitTestOleClipboard "UnitTestOleClipboard.cpp")

# Specify libraries or flags to use when linking a given target and/or its dependents.
target_link_libraries(UnitTestOleClipboard PUBLIC StraightOfficeStatic)

# Add a test to the project to be run by ctest.
add_test(
    NAME UnitTestOleClipboard
    COMMAND UnitTestOleClipboard
)

# Add source to the test's executable test.
add_executable(UnitTestOleDialogBox "UnitTestOleDialogBox.cpp")

# Specify libraries or flags to use when linking a given target and/or its dependents.
target_link_libraries(UnitTestOleDialogBox PUBLIC StraightOfficeStatic)

# Add a test to the project to be run by ctest.
add_test(
    NAME UnitTestOleDialogBox
    COMMAND UnitTestOleDialogBox
)

# Add source to the test's executable test.
add_executable(UnitTestOleVariant "UnitTestOleVariant.cpp")

# Specify libraries or flags to use when linking a given target and/or its dependents.
target_link_libraries(UnitTestOleVariant PUBLIC StraightOfficeStatic)

# Add a test to the project to be run by ctest.
add_test(
    NAME UnitTestOleVariant
    COMMAND UnitTestOleVariant
)

# Add source to the test's executable test.
add_executable(UnitTestPptProgram "UnitTestPptProgram.cpp")

# Specify libraries or flags to use when linking a given target and/or its dependents.
target_link_libraries(UnitTestPptProgram PUBLIC StraightOfficeStatic)

# Add a test to the project to be run by ctest.
add_test(
    NAME UnitTestPptProgram
    COMMAND UnitTestPptProgram
)

# Add source to the test's executable test.
add_executable(UnitTestExcelProgram "UnitTestExcelProgram.cpp")

# Specify libraries or flags to use when linking a given target and/or its dependents.
target_link_libraries(UnitTestExcelProgram PUBLIC StraightOfficeStatic)

# Add a test to the project to be run by ctest.
add_test(
    NAME UnitTestExcelProgram
    COMMAND UnitTestExcelProgram
)

# Add source to the test's executable test.
add_executable(UnitTestWordProgram "UnitTestWordProgram.cpp")

# Specify libraries or flags to use when linking a given target and/or its dependents.
target_link_libraries(UnitTestWordProgram PUBLIC StraightOfficeStatic)

# Add a test to the project to be run by ctest.
add_test(
    NAME UnitTestWordProgram
    COMMAND UnitTestWordProgram
)