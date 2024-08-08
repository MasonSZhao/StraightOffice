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

# Glob all files ended with .h and .cpp in the folders.
file (GLOB STRAIGHTOLE_SRCS CONFIGURE_DEPENDS "src/StraightOle/StraightOle*.cpp")
file (GLOB STRAIGHTOLE_HEADERS CONFIGURE_DEPENDS "include/StraightOle/StraightOle*.h")

# Glob all files ended with .h and .cpp in the folders.
file (GLOB STRAIGHTWORD_SRCS CONFIGURE_DEPENDS "src/StraightWord/StraightWord*.cpp")
file (GLOB STRAIGHTWORD_HEADERS CONFIGURE_DEPENDS "include/StraightWord/StraightWord*.h")

# Glob all files ended with .h and .cpp in the folders.
file (GLOB STRAIGHTEXCEL_SRCS CONFIGURE_DEPENDS "src/StraightExcel/StraightExcel*.cpp")
file (GLOB STRAIGHTEXCEL_HEADERS CONFIGURE_DEPENDS "include/StraightExcel/StraightExcel*.h")

# Glob all files ended with .h and .cpp in the folders.
file (GLOB STRAIGHTPPT_SRCS CONFIGURE_DEPENDS "src/StraightPpt/StraightPpt*.cpp")
file (GLOB STRAIGHTPPT_HEADERS CONFIGURE_DEPENDS "include/StraightPpt/StraightPpt*.h")

# Add source to the project's executable application.
add_executable(
	StraightOfficeExecute
	"apps/StraightOffice.cpp"
	"apps/StraightOffice.h"
	${STRAIGHTOLE_SRCS}
	${STRAIGHTOLE_HEADERS}
	${STRAIGHTWORD_SRCS}
	${STRAIGHTWORD_HEADERS}
	${STRAIGHTEXCEL_SRCS}
	${STRAIGHTEXCEL_HEADERS}
	${STRAIGHTPPT_SRCS}
	${STRAIGHTPPT_HEADERS}
)

# Add source to the project's static library.
add_library(
	StraightOfficeStatic 
	STATIC
	${STRAIGHTOLE_SRCS}
	${STRAIGHTOLE_HEADERS}
	${STRAIGHTWORD_SRCS}
	${STRAIGHTWORD_HEADERS}
	${STRAIGHTEXCEL_SRCS}
	${STRAIGHTEXCEL_HEADERS}
	${STRAIGHTPPT_SRCS}
	${STRAIGHTPPT_HEADERS}
)

# Add include directories to the project's static library.
target_include_directories(
	StraightOfficeStatic 
	PUBLIC "$<BUILD_INTERFACE:${CMAKE_CURRENT_LIST_DIR}/include/StraightOle>"
	PUBLIC "$<BUILD_INTERFACE:${CMAKE_CURRENT_LIST_DIR}/include/StraightWord>"
	PUBLIC "$<BUILD_INTERFACE:${CMAKE_CURRENT_LIST_DIR}/include/StraightExcel>"
	PUBLIC "$<BUILD_INTERFACE:${CMAKE_CURRENT_LIST_DIR}/include/StraightPpt>"
)

# Test option
set(STRAIGHTOFFICE_BUILD_TESTS OFF)

if (STRAIGHTOFFICE_BUILD_TESTS)

	# Enable testing for current directory and below.
	enable_testing()

	# Add a subdirectory to the build.
	add_subdirectory("tests")

endif()
