# Copyright (c) Microsoft Corporation. All rights reserved.
# Licensed under the MIT License.

cmake_minimum_required(VERSION 3.14)

project(InstanaPlugin)

# don't use the PAL for tests
build_init(CPP InstrumentationEngine)

set(src_files
    ./dllmain.cpp
	./InstanaInstrumentation.cpp
    ./InstrumentationEngineApi.cpp
    ./stdafx.cpp
)


add_lib(${PROJECT_NAME}
    SHARED
    false # use_pal
    false # use_redefines
    true # hide_symbols
    ${src_files}
    )

target_include_directories(${PROJECT_NAME} PUBLIC ${REPOSITORY_ROOT}/src/unix/inc/mincor )

# For some reason linking against musl-libc we need libintl and on glibc we need libdl.
# They have nothing to do with one another, each is just a quirk of the environment.
if(CLR_CMAKE_PLATFORM_ALPINE_LINUX)
    target_link_libraries(${PROJECT_NAME} :libintl.so)
elseif(CMAKE_SYSTEM_NAME STREQUAL Linux)
    # NOTE: Explicitly link to the .so version of libdl.so to make sure we don't somehow accidentally
    # start statically linking to these, as these are LGPL components.
    target_link_libraries(${PROJECT_NAME} :libdl.so)
else()
	message("Inside the target link libraries of windows")
    target_link_libraries(${PROJECT_NAME} dl)
endif()

target_link_libraries(
    ${PROJECT_NAME}
    pthread
    stdc++
    Common.Lib
)

if (CMAKE_SYSTEM_NAME STREQUAL Linux)
    target_link_libraries(${PROJECT_NAME} stdc++fs)
elseif(CMAKE_SYSTEM_NAME STREQUAL Darwin)
    target_link_libraries(${PROJECT_NAME}
        iconv
        "-framework CoreFoundation"
        "-framework Security"
    )
endif()