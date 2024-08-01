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
