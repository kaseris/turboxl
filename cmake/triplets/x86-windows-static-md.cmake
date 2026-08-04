# Python extensions must use the same dynamic MSVC runtime as CPython, while
# third-party libraries are static so the wheel has no extra DLLs to bundle.
set(VCPKG_TARGET_ARCHITECTURE x86)
set(VCPKG_CRT_LINKAGE dynamic)
set(VCPKG_LIBRARY_LINKAGE static)
