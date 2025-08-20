cmake_minimum_required(VERSION 3.21)

add_executable(exc3l-matrix-editor
    finaleprova.cpp
)

# xlnt via vcpkg or system install providing a CMake config file
find_package(xlnt CONFIG REQUIRED)

target_link_libraries(exc3l-matrix-editor PRIVATE xlnt::xlnt)
target_compile_features(exc3l-matrix-editor PRIVATE cxx_std_17)


