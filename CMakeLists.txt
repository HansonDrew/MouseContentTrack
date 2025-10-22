cmake_minimum_required(VERSION 3.15)
project(MouseContentTracker)

# 设置 C++ 标准
set(CMAKE_CXX_STANDARD 17)
set(CMAKE_CXX_STANDARD_REQUIRED ON)

# 源文件
set(SOURCES
    main.cpp
    MouseTracker.cpp
    MouseTracker.h
    BrowserContentExtractor.h
    BrowserContentExtractor.cpp
    Logger.cpp
    Logger.h
)

# 创建可执行文件
add_executable(MouseContentTracker ${SOURCES})

# Windows 特定设置
if(WIN32)
    # 链接必要的 Windows 库
    target_link_libraries(MouseContentTracker
        oleacc
        psapi
        ole32
        oleaut32
        uuid
    )
    
    # 设置 Windows 子系统为控制台
    set_target_properties(MouseContentTracker PROPERTIES
        WIN32_EXECUTABLE FALSE
    )
    
    # 添加 Unicode 支持
    target_compile_definitions(MouseContentTracker PRIVATE
        UNICODE
        _UNICODE
    )
endif()

# 设置输出目录
set_target_properties(MouseContentTracker PROPERTIES
    RUNTIME_OUTPUT_DIRECTORY "${CMAKE_BINARY_DIR}/bin"
)
