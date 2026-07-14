cmake_minimum_required(VERSION 3.20)

project(Vendetta VERSION 0.2.0 LANGUAGES CXX)

option(VENDETTA_ENABLE_PYTHON "Enable the experimental Python gameplay scripting layer" OFF)

set(CMAKE_CXX_STANDARD 17)
set(CMAKE_CXX_STANDARD_REQUIRED ON)
set(CMAKE_CXX_EXTENSIONS OFF)

find_package(SFML 2.5 COMPONENTS graphics window system REQUIRED)

file(GLOB_RECURSE VENDETTA_SOURCES CONFIGURE_DEPENDS
    "${CMAKE_CURRENT_SOURCE_DIR}/src/*.cpp"
)

add_executable(vendetta ${VENDETTA_SOURCES})
target_include_directories(vendetta PRIVATE "${CMAKE_CURRENT_SOURCE_DIR}/src")
target_link_libraries(vendetta PRIVATE sfml-graphics sfml-window sfml-system)

if(VENDETTA_ENABLE_PYTHON)
    find_package(Python3 COMPONENTS Development REQUIRED)
    find_package(Boost REQUIRED COMPONENTS "python${Python3_VERSION_MAJOR}${Python3_VERSION_MINOR}")
    target_link_libraries(vendetta PRIVATE Python3::Python "Boost::python${Python3_VERSION_MAJOR}${Python3_VERSION_MINOR}")
    target_compile_definitions(vendetta PRIVATE VENDETTA_ENABLE_PYTHON=1)
else()
    target_compile_definitions(vendetta PRIVATE VENDETTA_ENABLE_PYTHON=0)
endif()

if(MSVC)
    target_compile_options(vendetta PRIVATE /W4 /permissive-)
else()
    target_compile_options(vendetta PRIVATE -Wall -Wextra -Wpedantic)
endif()

add_custom_command(TARGET vendetta POST_BUILD
    COMMAND "${CMAKE_COMMAND}" -E copy_directory
            "${CMAKE_CURRENT_SOURCE_DIR}/Data"
            "$<TARGET_FILE_DIR:vendetta>/Data"
    COMMENT "Copying game data next to the executable"
)

install(TARGETS vendetta RUNTIME DESTINATION bin)
install(DIRECTORY Data DESTINATION bin)
