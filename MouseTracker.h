#pragma once

#include <windows.h>
#include <UIAutomation.h>
#include <string>
#include <vector>
#include <chrono>
#include <mutex>
#include <fstream>
#include <comdef.h>
#include <oleacc.h>
#include <queue>
#include <thread>
#include <condition_variable>
#include <atomic>

#pragma comment(lib, "oleacc.lib")

// 鼠标事件类型
enum class MouseEventType {
    LEFT_CLICK,
    LEFT_DOUBLE_CLICK,
    RIGHT_CLICK,
    TEXT_SELECTION,
    UNKNOWN
};

// 鼠标操作记录结构
struct MouseOperationRecord {
    std::chrono::system_clock::time_point timestamp;
    MouseEventType eventType;
    POINT position;
    std::wstring content;           // 交互的具体内容（链接、按钮名称、文本等）
    std::wstring applicationName;   // 所属应用程序名称
    std::wstring windowTitle;       // 窗口标题
    std::wstring elementType;       // 元素类型（按钮、链接、文本框等）

    std::wstring toJson() const;
};

// 待处理的鼠标事件
struct PendingMouseEvent {
    MouseEventType eventType;
    POINT position;
    HWND foregroundWindow;      // 前台活动窗口
    HWND pointWindow;           // 坐标位置的窗口（用于 UI Automation）
    std::chrono::system_clock::time_point timestamp;
};

class MouseTracker {
public:
    MouseTracker();
    ~MouseTracker();

    bool Initialize();
    void Start();
    void Stop();
    void SaveToFile(const std::wstring& filename);
    std::wstring GetAllRecordsAsJson();

private:
    static LRESULT CALLBACK MouseHookProc(int nCode, WPARAM wParam, LPARAM lParam);
    static MouseTracker* s_instance;

    void ProcessMouseEvent(WPARAM wParam, const MSLLHOOKSTRUCT* mouseInfo);
    void RecordMouseOperation(MouseEventType eventType, POINT position, HWND foregroundWindow, HWND pointWindow);
    void ProcessRecordQueue();  // 处理记录队列的工作线程
    
    // 返回元素内容和类型
    struct ElementInfo {
        std::wstring content;
        std::wstring elementType;
    };
    ElementInfo GetElementContentAtPoint(POINT pt, HWND targetWindow);
    
    std::wstring GetApplicationName(HWND hwnd);
    std::wstring GetWindowTitle(HWND hwnd);
    std::wstring GetElementTypeString(CONTROLTYPEID controlType);
    HWND GetRootOwnerWindow(HWND hwnd);  // 获取顶层窗口
    
    void CleanupOldRecords();  // 清理超过1小时的记录
    
    HHOOK m_mouseHook;
    IUIAutomation* m_pAutomation;
    
    std::vector<MouseOperationRecord> m_records;
    std::mutex m_recordsMutex;
    
    // 异步处理队列
    std::queue<PendingMouseEvent> m_eventQueue;
    std::mutex m_queueMutex;
    std::condition_variable m_queueCondition;
    std::thread m_processingThread;
    std::atomic<bool> m_isRunning;
    
    DWORD m_lastClickTime;
    POINT m_lastClickPos;
    
    std::wofstream m_logFile;
};

// 辅助函数
std::wstring MouseEventTypeToString(MouseEventType type);
std::wstring GetCurrentTimeString();
