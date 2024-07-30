# CMakeList.txt : CMake project for StraightOffice, include source and define project specific logic here.

# Require a minimum version of cmake.
cmake_minimum_required (VERSION 3.21)

# Set the C++ standard.
set(CMAKE_CXX_STANDARD 14)

# Enable Hot Reload for MSVC compilers if supported.
if (POLICY CMP0141)
  cmake_policy(SET CMP0141 NEW)
  set(CMAKE_MSVC_DEBUG_INFORMATION_FORMAT "$<IF:$<AND:$<C_COMPILER_ID:MSVC>,$<CXX_COMPILER_ID:MSVC>>,$<$<CONFIG:Debug,RelWithDebInfo>:EditAndContinue>,$<$<CONFIG:Debug,RelWithDebInfo>:ProgramDatabase>>")
endif()

# Set the name of the project.
project ("StraightOffice")

# Add source to the project's executable application.
add_executable (
	StraightOfficeExecute
	"apps/StraightOffice.cpp"
	"apps/StraightOffice.h"
)

# Test option
set(STRAIGHTOFFICE_BUILD_TESTS OFF)

if (STRAIGHTOFFICE_BUILD_TESTS)

	# Enable testing for current directory and below.
	enable_testing()

	# Add a subdirectory to the build.
	add_subdirectory("tests")

endif()
