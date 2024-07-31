# Add source to the test's executable test.
add_executable(UnitTestOleClipboard "UnitTestOleClipboard.cpp")

# Specify libraries or flags to use when linking a given target and/or its dependents.
target_link_libraries(UnitTestOleClipboard PUBLIC StraightOfficeStatic)

# Add a test to the project to be run by ctest.
add_test(
    NAME UnitTestOleClipboard
    COMMAND UnitTestOleClipboard
)
