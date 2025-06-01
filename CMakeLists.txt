cmake_minimum_required(VERSION 3.18)
project(SIMNRA++)

# Set C++ standard
set(CMAKE_CXX_STANDARD 17)
set(CMAKE_CXX_STANDARD_REQUIRED True)

# Handle CMP0148 policy
if (POLICY CMP0148)
  cmake_policy(SET CMP0148 NEW)
endif()

# Compiler flags and definitions
set(CMAKE_CXX_FLAGS "${CMAKE_CXX_FLAGS} /O2")
add_definitions(-D_CRT_SECURE_NO_WARNINGS)

# Libraries to link with COM
set(COM_LIBS ole32 oleaut32)

# --- SIMNRApp Library ---
file(GLOB_RECURSE SIMNRApp_SOURCES src/*.cpp)
add_library(SIMNRApp STATIC ${SIMNRApp_SOURCES})
target_include_directories(SIMNRApp PUBLIC include)
target_link_libraries(SIMNRApp PRIVATE ${COM_LIBS})
set_target_properties(SIMNRApp PROPERTIES ARCHIVE_OUTPUT_DIRECTORY "${CMAKE_BINARY_DIR}/lib")

# --- Main executable ---
add_executable(simnra main/simnra_test.cpp)
target_link_libraries(simnra PRIVATE SIMNRApp ${COM_LIBS})
set_target_properties(simnra PROPERTIES RUNTIME_OUTPUT_DIRECTORY "${CMAKE_BINARY_DIR}/bin")
