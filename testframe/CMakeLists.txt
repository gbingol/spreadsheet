project(testframe)



file(GLOB testfiles "*.h" "*.cpp")

add_executable(${PROJECT_NAME} WIN32 ${testfiles})

set_property(TARGET ${PROJECT_NAME} PROPERTY VS_DPI_AWARE "PerMonitor")

target_compile_definitions(${PROJECT_NAME} PRIVATE 
_WINDOWS _USRDLL _UNICODE __WXMSW__ wxUSE_NATIVE_DATAVIEWCTRL
		wxUSE_MENUBAR WXUSINGDLL _CRT_SECURE_NO_WARNINGS _SILENCE_ALL_CXX17_DEPRECATION_WARNINGS)


find_package(wxWidgets REQUIRED core base OPTIONAL_COMPONENTS aui net xml)

include(${wxWidgets_USE_FILE})
include_directories(${CMAKE_SOURCE_DIR}/wxgrid)

target_include_directories(${PROJECT_NAME} PRIVATE ${wxWidgets_INCLUDE_DIRS})
target_link_directories(${PROJECT_NAME} PRIVATE ${wxWidgets_LIBRARY_DIRS} )
target_link_libraries(${PROJECT_NAME} ${wxWidgets_LIBRARIES} grid)

