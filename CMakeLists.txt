cmake_minimum_required(VERSION 3.27.2)
set(CMAKE_CXX_STANDARD 20)



project(spreadsheet)

SET(CMAKE_RUNTIME_OUTPUT_DIRECTORY_RELEASE ${CMAKE_SOURCE_DIR}/libs)
SET(CMAKE_RUNTIME_OUTPUT_DIRECTORY_DEBUG ${CMAKE_SOURCE_DIR}/libs)
SET(CMAKE_ARCHIVE_OUTPUT_DIRECTORY_RELEASE  ${CMAKE_SOURCE_DIR}/libs)

include(C:/vcpkg/scripts/buildsystems/vcpkg.cmake)

add_subdirectory(grid)
add_subdirectory(testframe)