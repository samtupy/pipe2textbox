cmake_minimum_required(VERSION 3.20)
project(show C)

# Much of the following is to get as small of a binary as possible; CMake loves to insert default flags for us.
foreach(var CMAKE_C_FLAGS CMAKE_C_FLAGS_DEBUG CMAKE_C_FLAGS_RELEASE CMAKE_C_FLAGS_RELWITHDEBINFO CMAKE_C_FLAGS_MINSIZEREL CMAKE_EXE_LINKER_FLAGS CMAKE_EXE_LINKER_FLAGS_DEBUG CMAKE_EXE_LINKER_FLAGS_RELEASE CMAKE_EXE_LINKER_FLAGS_RELWITHDEBINFO CMAKE_EXE_LINKER_FLAGS_MINSIZEREL)
	set(${var} "" CACHE STRING "" FORCE)
endforeach()
add_link_options(/MANIFEST:NO /DYNAMICBASE:NO /NXCOMPAT:NO /SAFESEH:NO /NODEFAULTLIB /ENTRY:main)
add_compile_options(/GS- /GF /Os /O1 /DNDEBUG /TC)

enable_language(RC)
set_source_files_properties(textbox.rc PROPERTIES LANGUAGE RC)
add_executable(show show.c textbox.rc)
target_link_libraries(show advapi32 comdlg32 comctl32 kernel32 user32)
