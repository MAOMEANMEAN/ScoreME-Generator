cmake_minimum_required(VERSION 3.14)
project(ScoreME_Generator)

set(CMAKE_CXX_STANDARD 17)
set(CMAKE_CXX_STANDARD_REQUIRED ON)

include(FetchContent)

# Fetch xlnt
FetchContent_Declare(
    xlnt
    GIT_REPOSITORY https://github.com/tfussell/xlnt.git
    GIT_TAG master
)
FetchContent_MakeAvailable(xlnt)

# Fetch tabulate
FetchContent_Declare(
    tabulate
    GIT_REPOSITORY https://github.com/p-ranav/tabulate.git
    GIT_TAG master
)
FetchContent_MakeAvailable(tabulate)

# Create data directories in both source and build locations
file(MAKE_DIRECTORY ${CMAKE_SOURCE_DIR}/data)
file(MAKE_DIRECTORY ${CMAKE_SOURCE_DIR}/data/backups)
file(MAKE_DIRECTORY ${CMAKE_BINARY_DIR}/data)
file(MAKE_DIRECTORY ${CMAKE_BINARY_DIR}/data/backups)

# Copy data directory to build directory for easy access
add_custom_target(copy_data_structure ALL
    COMMAND ${CMAKE_COMMAND} -E copy_directory
    ${CMAKE_SOURCE_DIR}/data ${CMAKE_BINARY_DIR}/data
    COMMENT "Copying data directory structure to build directory"
)

# Add all source files (REMOVED ExcelUtil.cpp since it contained duplicate functions)
set(SOURCES
    main.cpp
    src/Person.cpp
    src/Student.cpp
    src/Admin.cpp
    src/ExcelUtil.cpp
    src/MenuUtils.cpp
    src/GradeUtil.cpp
)

# Add executable
add_executable(ScoreME_Generator ${SOURCES})

# Include directories
target_include_directories(ScoreME_Generator PRIVATE
    include
    ${CMAKE_SOURCE_DIR}/include
)

# Link libraries
target_link_libraries(ScoreME_Generator PRIVATE
    xlnt
    tabulate
)

# Compiler-specific options
if(CMAKE_CXX_COMPILER_ID STREQUAL "GNU" OR CMAKE_CXX_COMPILER_ID STREQUAL "Clang")
    target_compile_options(ScoreME_Generator PRIVATE -Wall -Wextra -O2)
endif()

# Copy DLL on Windows (if needed)
if(WIN32)
    add_custom_command(TARGET ScoreME_Generator POST_BUILD
        COMMAND ${CMAKE_COMMAND} -E copy_if_different
        "${CMAKE_SOURCE_DIR}/build/_deps/xlnt-build/source/Release/xlnt.dll"
        $<TARGET_FILE_DIR:ScoreME_Generator>
        COMMENT "Copying xlnt.dll to output directory"
    )
endif()

# Create sample Excel files with data
add_custom_target(create_sample_excel_data ALL
    COMMAND ${CMAKE_COMMAND} -E echo "Creating sample Excel data files..."
    DEPENDS ScoreME_Generator copy_data_structure
)

# Create sample data files after build
add_custom_command(TARGET ScoreME_Generator POST_BUILD
    COMMAND $<TARGET_FILE:ScoreME_Generator> --create-sample-data
    WORKING_DIRECTORY ${CMAKE_BINARY_DIR}
    COMMENT "Creating sample Excel data files"
)

# Install target
install(TARGETS ScoreME_Generator
    RUNTIME DESTINATION bin
)

# Print build information
message(STATUS "Building ScoreME_Generator")
message(STATUS "C++ Standard: ${CMAKE_CXX_STANDARD}")
message(STATUS "Build Type: ${CMAKE_BUILD_TYPE}")
message(STATUS "Compiler: ${CMAKE_CXX_COMPILER_ID}")