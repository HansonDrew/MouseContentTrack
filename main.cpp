#include "MouseTracker.h"
#include <iostream>
#include <locale>
#include <io.h>
#include <fcntl.h>

int main() {
    // 设置控制台支持 Unicode
    _setmode(_fileno(stdout), _O_U16TEXT);
    _setmode(_fileno(stdin), _O_U16TEXT);
    std::locale::global(std::locale(""));

    std::wcout << L"========================================\n";
    std::wcout << L"   Windows 鼠标操作追踪器\n";
    std::wcout << L"   Mouse Content Tracker v1.0\n";
    std::wcout << L"========================================\n\n";

    MouseTracker tracker;
    
    if (!tracker.Initialize()) {
        std::wcerr << L"错误: 初始化失败!\n";
        std::wcerr << L"请确保以管理员权限运行此程序。\n";
        return 1;
    }

    std::wcout << L"初始化成功!\n";
    std::wcout << L"正在启动鼠标追踪...\n\n";
    
    tracker.Start();

    std::wcout << L"追踪已开始! 系统将记录过去1小时内的所有鼠标操作。\n";
    std::wcout << L"功能说明:\n";
    std::wcout << L"  - 捕获鼠标单击、双击、右键事件\n";
    std::wcout << L"  - 识别点击位置的元素内容（按钮、链接、文本等）\n";
    std::wcout << L"  - 记录所属应用程序和窗口信息\n";
    std::wcout << L"  - 自动清理1小时前的记录\n";
    std::wcout << L"  - 忽略拖动窗口的操作\n\n";
    std::wcout << L"操作说明:\n";
    std::wcout << L"  按 's' + Enter 保存记录到 JSON 文件\n";
    std::wcout << L"  按 'p' + Enter 打印所有记录\n";
    std::wcout << L"  按 'q' + Enter 退出程序\n\n";
    std::wcout << L"----------------------------------------\n";

    // 消息循环
    MSG msg;
    bool running = true;
    
    // 创建一个线程来处理用户输入
    std::thread inputThread([&]() {
        while (running) {
            wchar_t input;
            std::wcin >> input;
            
            if (input == L'q' || input == L'Q') {
                std::wcout << L"\n正在退出程序...\n";
                running = false;
                PostQuitMessage(0);
            }
            else if (input == L's' || input == L'S') {
                std::wstring filename = L"mouse_records_" + GetCurrentTimeString() + L".json";
                // 替换文件名中的非法字符
                for (auto& c : filename) {
                    if (c == L':' || c == L' ') c = L'_';
                }
                tracker.SaveToFile(filename);
                std::wcout << L"\n记录已保存到: " << filename << L"\n";
            }
            else if (input == L'p' || input == L'P') {
                std::wcout << L"\n========== 所有记录 (JSON格式) ==========\n";
                std::wcout << tracker.GetAllRecordsAsJson() << L"\n";
                std::wcout << L"========================================\n\n";
            }
        }
    });

    // 主消息循环
    while (running && GetMessage(&msg, nullptr, 0, 0)) {
        TranslateMessage(&msg);
        DispatchMessage(&msg);
    }

    tracker.Stop();
    
    if (inputThread.joinable()) {
        inputThread.join();
    }

    std::wcout << L"\n程序已退出。\n";
    return 0;
}
