cmake_minimum_required(VERSION 3.20)
project(turboxl LANGUAGES CXX)

set(CMAKE_CXX_STANDARD 20)
set(CMAKE_CXX_STANDARD_REQUIRED ON)

option(BUILD_TESTS "Build tests" OFF)
option(BUILD_PYTHON "Build Python module" ON)
option(BUILD_CLI "Build CLI tool" OFF)
option(BUILD_BENCHMARKS "Build benchmarks" OFF)


if(WIN32)
    find_package(LibXml2 REQUIRED)
    find_package(ZLIB REQUIRED)

    find_package(minizip CONFIG REQUIRED)

    set(LIBXML2_INCLUDE_DIRS ${LIBXML2_INCLUDE_DIRS})
    set(LIBXML2_LIBRARIES ${LIBXML2_LIBRARIES})
    set(ZLIB_INCLUDE_DIRS ${ZLIB_INCLUDE_DIRS})
    set(ZLIB_LIBRARIES ${ZLIB_LIBRARIES})
    set(ZIP_LIBRARIES minizip::minizip)

else()
    find_package(PkgConfig REQUIRED)

    # libxml2
    pkg_check_modules(LIBXML2 REQUIRED libxml-2.0)
    if(NOT LIBXML2_FOUND)
        message(FATAL_ERROR "libxml2 not found. Please install libxml2-dev")
    endif()

    # zlib-ng (preferred) or zlib
    pkg_check_modules(ZLIB_NG zlib-ng)
    if(ZLIB_NG_FOUND)
        set(ZLIB_LIBRARY "zlib-ng")
        set(ZLIB_INCLUDE_DIRS ${ZLIB_NG_INCLUDE_DIRS})
        set(ZLIB_LIBRARIES ${ZLIB_NG_LIBRARIES})
        message(STATUS "Using zlib-ng")
    else()
        find_package(ZLIB REQUIRED)
        set(ZLIB_INCLUDE_DIRS ${ZLIB_INCLUDE_DIRS})
        set(ZLIB_LIBRARIES ${ZLIB_LIBRARIES})
        message(STATUS "Using standard zlib")
    endif()

    # minizip-ng or libzip
    pkg_check_modules(MINIZIP minizip-ng)
    if(MINIZIP_FOUND)
        set(ZIP_LIBRARIES ${MINIZIP_LIBRARIES})
        set(ZIP_INCLUDE_DIRS ${MINIZIP_INCLUDE_DIRS})
    else()
        pkg_check_modules(LIBZIP libzip)
        if(LIBZIP_FOUND)
            set(ZIP_LIBRARIES ${LIBZIP_LIBRARIES})
            set(ZIP_INCLUDE_DIRS ${LIBZIP_INCLUDE_DIRS})
        else()
            message(FATAL_ERROR "Neither minizip-ng nor libzip found.")
        endif()
    endif()
endif()

add_library(turboxl_core
    src/core/zip_reader.cpp
    src/core/opc_package.cpp
    src/core/workbook.cpp
    src/core/styles_registry.cpp
    src/core/shared_strings_provider.cpp
    src/core/cell_data.cpp
    src/core/sheet_stream_reader.cpp
    src/core/data_converter.cpp
    src/csv/csv_encoder.cpp
    src/facade/xlsx_reader.cpp
)

target_include_directories(turboxl_core
    PUBLIC
        $<BUILD_INTERFACE:${CMAKE_CURRENT_SOURCE_DIR}/include>
        $<INSTALL_INTERFACE:include>
    PRIVATE
        ${CMAKE_CURRENT_SOURCE_DIR}/src
        ${LIBXML2_INCLUDE_DIRS}
        ${ZIP_INCLUDE_DIRS}
        ${ZLIB_INCLUDE_DIRS}
)

target_link_libraries(turboxl_core
    PRIVATE
        ${LIBXML2_LIBRARIES}
        ${ZIP_LIBRARIES}
        ${ZLIB_LIBRARIES}
)

target_compile_features(turboxl_core PUBLIC cxx_std_20)

if(MSVC)
    target_compile_options(turboxl_core PRIVATE /W4 /wd4251)
    target_compile_definitions(turboxl_core PRIVATE _CRT_SECURE_NO_WARNINGS)
    if(CMAKE_BUILD_TYPE STREQUAL "Release")
        target_compile_options(turboxl_core PRIVATE /O2 /GL /arch:AVX2)
        target_link_options(turboxl_core PRIVATE /LTCG)
    endif()
else()
    target_compile_options(turboxl_core PRIVATE -Wall -Wextra -Wpedantic -fvisibility=hidden)
    if(CMAKE_BUILD_TYPE STREQUAL "Release")
        target_compile_options(turboxl_core PRIVATE -O3 -march=native -flto)
        target_link_options(turboxl_core PRIVATE -flto)
    endif()
endif()

if(BUILD_PYTHON)
    find_package(pybind11 CONFIG REQUIRED)
    pybind11_add_module(turboxl src/python/module.cpp)
    target_link_libraries(turboxl PRIVATE turboxl_core)

    set_target_properties(turboxl PROPERTIES
        CXX_VISIBILITY_PRESET hidden
        VISIBILITY_INLINES_HIDDEN ON
    )

    install(TARGETS turboxl
        LIBRARY DESTINATION .
        RUNTIME DESTINATION .
        ARCHIVE DESTINATION .
    )
endif()

if(BUILD_TESTS)
    enable_testing()
    find_package(GTest QUIET)
    if(NOT GTest_FOUND)
        include(FetchContent)
        FetchContent_Declare(
            googletest
            GIT_REPOSITORY https://github.com/google/googletest.git
            GIT_TAG release-1.12.1
        )
        FetchContent_MakeAvailable(googletest)
    endif()

    add_executable(turboxl_tests
        tests/main.cpp
        tests/test_zip_reader.cpp
        tests/test_opc_package.cpp
        tests/test_workbook.cpp
        tests/test_styles_registry.cpp
        tests/test_shared_strings_provider.cpp
        tests/test_sheet_stream_reader.cpp
        tests/test_sheet_discovery.cpp
        tests/test_phase5_functionality.cpp
        tests/test_csv_encoder.cpp
        tests/test_integration.cpp
    )

    target_link_libraries(turboxl_tests
        PRIVATE turboxl_core GTest::gtest GTest::gtest_main
    )
    add_test(NAME core_tests COMMAND turboxl_tests)
endif()

if(BUILD_CLI)
    add_executable(turboxl_cli tools/cli.cpp)
    target_link_libraries(turboxl_cli PRIVATE turboxl_core)
endif()

if(BUILD_BENCHMARKS)
    find_package(benchmark QUIET)
    if(NOT benchmark_FOUND)
        include(FetchContent)
        FetchContent_Declare(
            benchmark
            GIT_REPOSITORY https://github.com/google/benchmark.git
            GIT_TAG v1.8.3
        )
        FetchContent_MakeAvailable(benchmark)
    endif()

    add_executable(turboxl_benchmarks benchmarks/main.cpp)
    target_link_libraries(turboxl_benchmarks PRIVATE turboxl_core benchmark::benchmark)
endif()

install(TARGETS turboxl_core
    EXPORT turboxlTargets
    LIBRARY DESTINATION lib
    ARCHIVE DESTINATION lib
    RUNTIME DESTINATION bin
    INCLUDES DESTINATION include
)

install(DIRECTORY include/ DESTINATION include)

install(EXPORT turboxlTargets
    FILE turboxlTargets.cmake
    NAMESPACE turboxl::
    DESTINATION lib/cmake/turboxl
)

include(CMakePackageConfigHelpers)
write_basic_package_version_file(
    "${CMAKE_CURRENT_BINARY_DIR}/turboxlConfigVersion.cmake"
    VERSION 0.1.0
    COMPATIBILITY AnyNewerVersion
)

configure_file(
    "${CMAKE_CURRENT_SOURCE_DIR}/cmake/turboxlConfig.cmake.in"
    "${CMAKE_CURRENT_BINARY_DIR}/turboxlConfig.cmake"
    @ONLY
)

install(FILES
    "${CMAKE_CURRENT_BINARY_DIR}/turboxlConfig.cmake"
    "${CMAKE_CURRENT_BINARY_DIR}/turboxlConfigVersion.cmake"
    DESTINATION lib/cmake/turboxl
)
