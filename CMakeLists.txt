cmake_minimum_required(VERSION 3.27.2)
set(CMAKE_CXX_STANDARD 20)


project(spreadsheet)



include(C:/vcpkg/scripts/buildsystems/vcpkg.cmake)

add_subdirectory(grid)
add_subdirectory(testframe)