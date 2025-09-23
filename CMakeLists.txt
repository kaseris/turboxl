cmake_minimum_required(VERSION 3.20)
project(turboxl LANGUAGES CXX)

# Set C++ standard
set(CMAKE_CXX_STANDARD 20)
set(CMAKE_CXX_STANDARD_REQUIRED ON)

# Project options
option(BUILD_TESTS "Build tests" ON)
option(BUILD_PYTHON "Build Python module" ON)
option(BUILD_CLI "Build CLI tool" OFF)
option(BUILD_BENCHMARKS "Build benchmarks" OFF)

# Find required packages
find_package(PkgConfig REQUIRED)

# --- Windows-specific discovery (minimal guard) ---
if(WIN32)
    # Prefer CMake package configs supplied by vcpkg; keep variable names used below.
    find_package(LibXml2 REQUIRED)
    # The FindLibXml2 module defines LIBXML2_INCLUDE_DIRS/LIBXML2_LIBRARIES

    # zlib: with zlib-ng[compat] in vcpkg, find_package(ZLIB) resolves to the
    # compatibility layer, so consumers still link against ZLIB as usual.
    find_package(ZLIB REQUIRED)
    set(ZLIB_LIBRARY "zlib")
    set(ZLIB_INCLUDE_DIRS ${ZLIB_INCLUDE_DIRS})
    set(ZLIB_LIBRARIES ${ZLIB_LIBRARIES})
    set(ZLIB_LIBRARY_DIRS "")

    # minizip-ng via its CMake config (prefer config packages); fall back to manual override
    find_package(minizip CONFIG QUIET)
    if(minizip_FOUND)
        # Accept common exported target names
        if(TARGET minizip::minizip)
            set(ZIP_LIBRARY "minizip-ng")
            set(ZIP_INCLUDE_DIRS "")
            set(ZIP_LIBRARIES minizip::minizip)
        elseif(TARGET minizip)
            # Create expected namespace alias if only plain target exists
            add_library(minizip::minizip ALIAS minizip)
            set(ZIP_LIBRARY "minizip-ng")
            set(ZIP_INCLUDE_DIRS "")
            set(ZIP_LIBRARIES minizip::minizip)
        else()
            # Allow explicit override if a config is present but no expected targets are exported
            if(DEFINED MINIZIP_LIBRARY AND DEFINED MINIZIP_INCLUDE_DIR)
                add_library(minizip::minizip STATIC IMPORTED)
                set_target_properties(minizip::minizip PROPERTIES
                    IMPORTED_LOCATION "${MINIZIP_LIBRARY}"
                    INTERFACE_INCLUDE_DIRECTORIES "${MINIZIP_INCLUDE_DIR}"
                )
                set(ZIP_LIBRARY "minizip-ng")
                set(ZIP_INCLUDE_DIRS "${MINIZIP_INCLUDE_DIR}")
                set(ZIP_LIBRARIES minizip::minizip)
            else()
                message(FATAL_ERROR "minizip package found, but expected target not exported; set MINIZIP_LIBRARY and MINIZIP_INCLUDE_DIR to override.")
            endif()
        endif()
    else()
        # Final fallback: manual override via variables provided by the build system
        if(DEFINED MINIZIP_LIBRARY AND DEFINED MINIZIP_INCLUDE_DIR)
            add_library(minizip::minizip STATIC IMPORTED)
            set_target_properties(minizip::minizip PROPERTIES
                IMPORTED_LOCATION "${MINIZIP_LIBRARY}"
                INTERFACE_INCLUDE_DIRECTORIES "${MINIZIP_INCLUDE_DIR}"
            )
            set(ZIP_LIBRARY "minizip-ng")
            set(ZIP_INCLUDE_DIRS "${MINIZIP_INCLUDE_DIR}")
            set(ZIP_LIBRARIES minizip::minizip)
        else()
            message(FATAL_ERROR "minizip-ng not found (install via vcpkg) and no manual override provided.")
        endif()
    endif()

else()
    # --- Non-Windows: keep existing pkg-config based discovery ---
    # Find libxml2
    pkg_check_modules(LIBXML2 REQUIRED libxml-2.0)
    if(NOT LIBXML2_FOUND)
        message(FATAL_ERROR "libxml2 not found. Please install libxml2-dev")
    endif()

    # Find zlib-ng with compatibility mode (preferred) or fallback to standard zlib
    find_package(PkgConfig QUIET)
    pkg_check_modules(ZLIB_NG zlib-ng)
    if(ZLIB_NG_FOUND)
        set(ZLIB_LIBRARY "zlib-ng")
        set(ZLIB_INCLUDE_DIRS ${ZLIB_NG_INCLUDE_DIRS})
        set(ZLIB_LIBRARIES ${ZLIB_NG_LIBRARIES})
        set(ZLIB_LIBRARY_DIRS ${ZLIB_NG_LIBRARY_DIRS})
        message(STATUS "Using zlib-ng for optimized compression performance")
    else()
        find_package(ZLIB REQUIRED)
        set(ZLIB_LIBRARY "zlib")
        set(ZLIB_INCLUDE_DIRS ${ZLIB_INCLUDE_DIRS})
        set(ZLIB_LIBRARIES ${ZLIB_LIBRARIES})
        set(ZLIB_LIBRARY_DIRS "")
        message(STATUS "Using standard zlib (consider installing zlib-ng for better performance)")
    endif()

    # Find minizip-ng or libzip
    pkg_check_modules(MINIZIP minizip-ng)
    if(MINIZIP_FOUND)
        set(ZIP_LIBRARY "minizip-ng")
        set(ZIP_INCLUDE_DIRS ${MINIZIP_INCLUDE_DIRS})
        # Also add parent include directory so includes like <minizip-ng/...> resolve
        foreach(_inc IN LISTS MINIZIP_INCLUDE_DIRS)
            get_filename_component(_parent_inc "${_inc}" DIRECTORY)
            list(APPEND ZIP_INCLUDE_DIRS "${_parent_inc}")
        endforeach()
        set(ZIP_LIBRARIES ${MINIZIP_LIBRARIES})
    else()
        pkg_check_modules(LIBZIP libzip)
        if(LIBZIP_FOUND)
            set(ZIP_LIBRARY "libzip")
            set(ZIP_INCLUDE_DIRS ${LIBZIP_INCLUDE_DIRS})
            set(ZIP_LIBRARIES ${LIBZIP_LIBRARIES})
        else()
            message(FATAL_ERROR "Neither minizip-ng nor libzip found. Please install one of them.")
        endif()
    endif()
endif()

# Core library
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
        ${CMAKE_CURRENT_SOURCE_DIR}/include
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

# Add library search directories for zlib-ng if needed
if(ZLIB_LIBRARY_DIRS)
    target_link_directories(turboxl_core PRIVATE ${ZLIB_LIBRARY_DIRS})
endif()

# Add library search directories for minizip/libzip if provided
if(MINIZIP_FOUND AND MINIZIP_LIBRARY_DIRS)
    target_link_directories(turboxl_core PRIVATE ${MINIZIP_LIBRARY_DIRS})
endif()
if(LIBZIP_FOUND AND LIBZIP_LIBRARY_DIRS)
    target_link_directories(turboxl_core PRIVATE ${LIBZIP_LIBRARY_DIRS})
endif()

target_compile_features(turboxl_core PUBLIC cxx_std_20)

# Compiler-specific settings
if(CMAKE_CXX_COMPILER_ID MATCHES "Clang|GNU")
    target_compile_options(turboxl_core PRIVATE
        -Wall -Wextra -Wpedantic
        -fvisibility=hidden
    )
    
    # Performance optimization flags for Release builds
    if(CMAKE_BUILD_TYPE STREQUAL "Release")
        target_compile_options(turboxl_core PRIVATE
            -O3 -march=native -flto
        )
        target_link_options(turboxl_core PRIVATE -flto)
    endif()
endif()

if(MSVC)
    target_compile_options(turboxl_core PRIVATE
        /W4 /wd4251
    )
    target_compile_definitions(turboxl_core PRIVATE
        _CRT_SECURE_NO_WARNINGS
    )
    
    # Performance optimization flags for Release builds
    if(CMAKE_BUILD_TYPE STREQUAL "Release")
        target_compile_options(turboxl_core PRIVATE
            /O2 /GL /arch:AVX2
        )
        target_link_options(turboxl_core PRIVATE /LTCG)
    endif()
endif()

# Python bindings
if(BUILD_PYTHON)
    find_package(pybind11 CONFIG REQUIRED)
    
    pybind11_add_module(turboxl src/python/module.cpp)
    target_link_libraries(turboxl PRIVATE turboxl_core)
    
    # Add library search directories for zlib-ng if needed
    if(ZLIB_LIBRARY_DIRS)
        target_link_directories(turboxl PRIVATE ${ZLIB_LIBRARY_DIRS})
    endif()
    # Add library search directories for minizip/libzip if provided
    if(MINIZIP_FOUND AND MINIZIP_LIBRARY_DIRS)
        target_link_directories(turboxl PRIVATE ${MINIZIP_LIBRARY_DIRS})
    endif()
    if(LIBZIP_FOUND AND LIBZIP_LIBRARY_DIRS)
        target_link_directories(turboxl PRIVATE ${LIBZIP_LIBRARY_DIRS})
    endif()
    
    # Ensure hidden visibility for small ABI surface
    set_target_properties(turboxl PROPERTIES
        CXX_VISIBILITY_PRESET hidden
        VISIBILITY_INLINES_HIDDEN ON
    )

    # Install the Python extension into the wheel's platlib root
    install(TARGETS turboxl
        LIBRARY DESTINATION .
        RUNTIME DESTINATION .
        ARCHIVE DESTINATION .
    )
endif()

# Tests
if(BUILD_TESTS)
    enable_testing()
    
    # Find Google Test
    find_package(GTest QUIET)
    if(NOT GTest_FOUND)
        # Fallback: download and build GTest
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
        PRIVATE
            turboxl_core
            GTest::gtest
            GTest::gtest_main
    )
    
    # Add library search directories for zlib-ng if needed
    if(ZLIB_LIBRARY_DIRS)
        target_link_directories(turboxl_tests PRIVATE ${ZLIB_LIBRARY_DIRS})
    endif()
    
    add_test(NAME core_tests COMMAND turboxl_tests)
endif()

# CLI tool
if(BUILD_CLI)
    add_executable(turboxl_cli tools/cli.cpp)
    target_link_libraries(turboxl_cli PRIVATE turboxl_core)
endif()

# Benchmarks
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
    target_link_libraries(turboxl_benchmarks
        PRIVATE
            turboxl_core
            benchmark::benchmark
    )
endif()

# Installation
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

# Create config file
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
