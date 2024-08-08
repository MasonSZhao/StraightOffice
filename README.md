# StraightOffice

Straight C++ Office (Word, Excel, Ppt) library.

## Install

### CMake Fetch Content

```CMake
include(FetchContent)

FetchContent_Declare(StraightOffice GIT_REPOSITORY https://github.com/MasonSZhao/StraightOffice)

FetchContent_MakeAvailable(StraightOffice)

target_link_libraries(<target_executable> StraightOfficeStatic)

target_link_libraries(<target_library> StraightOfficeStatic)
```

### CMake Link Library

Copy all folders in the include [folder](https://github.com/MasonSZhao/StraightOffice/tree/master/include) to your include folder.
Copy the static library file to your CMakeLists.txt folder.

```CMake
file (GLOB STRAIGHTOLE_HEADERS CONFIGURE_DEPENDS "include/StraightOle/StraightOle*.h")
file (GLOB STRAIGHTWORD_HEADERS CONFIGURE_DEPENDS "include/StraightWord/StraightWord*.h")
file (GLOB STRAIGHTEXCEL_HEADERS CONFIGURE_DEPENDS "include/StraightExcel/StraightExcel*.h")
file (GLOB STRAIGHTPPT_HEADERS CONFIGURE_DEPENDS "include/StraightPpt/StraightPpt*.h")

add_executable(<target_executable> ${STRAIGHTOLE_HEADERS} ${STRAIGHTWORD_HEADERS} ${STRAIGHTEXCEL_HEADERS} ${STRAIGHTPPT_HEADERS})

target_link_libraries(<target_executable> ${CMAKE_CURRENT_SOURCE_DIR}/StraightOfficeStatic.lib)

add_library(<target_library> STATIC ${STRAIGHTOLE_HEADERS} ${STRAIGHTWORD_HEADERS} ${STRAIGHTEXCEL_HEADERS} ${STRAIGHTPPT_HEADERS})

target_link_libraries(<target_library> ${CMAKE_CURRENT_SOURCE_DIR}/StraightOfficeStatic.lib)
```

## Contribute

This project uses the Trunk-Based Development.
