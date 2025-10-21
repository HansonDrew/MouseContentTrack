#include "MouseTracker.h"
#include <iostream>
#include <sstream>
#include <iomanip>
#include <psapi.h>
#include <atlbase.h>
#include <UIAutomationClient.h>

#pragma comment(lib, "psapi.lib")

MouseTracker* MouseTracker::s_instance = nullptr;

MouseTracker::MouseTracker() 
    : m_mouseHook(nullptr)
    , m_pAutomation(nullptr)
    , m_isRunning(false)
    , m_lastClickTime(0)
{
    m_lastClickPos.x = 0;
    m_lastClickPos.y = 0;
    s_instance = this;
}

MouseTracker::~MouseTracker() {
    Stop();
    if (m_pAutomation) {
        m_pAutomation->Release();
        m_pAutomation = nullptr;
    }
    if (m_logFile.is_open()) {
        m_logFile.close();
    }
    s_instance = nullptr;
}

bool MouseTracker::Initialize() {
    // 初始化 COM
    HRESULT hr = CoInitializeEx(nullptr, COINIT_MULTITHREADED);
    if (FAILED(hr)) {
        return false;
    }

    // 创建 UI Automation 实例
    hr = CoCreateInstance(__uuidof(CUIAutomation), nullptr, 
                          CLSCTX_INPROC_SERVER, 
                          __uuidof(IUIAutomation), 
                          (void**)&m_pAutomation);
    if (FAILED(hr)) {
        CoUninitialize();
        return false;
    }

    // 打开日志文件
    m_logFile.open(L"mouse_operations_log.txt", std::ios::app);
    if (!m_logFile.is_open()) {
        return false;
    }

    m_logFile << L"\n========== Mouse Tracker Started at " << GetCurrentTimeString() << L" ==========\n" << std::flush;

    return true;
}

void MouseTracker::Start() {
    if (m_isRunning) return;

    m_isRunning = true;

    // 启动处理线程
    m_processingThread = std::thread(&MouseTracker::ProcessRecordQueue, this);

    // 安装鼠标钩子
    m_mouseHook = SetWindowsHookEx(WH_MOUSE_LL, MouseHookProc, GetModuleHandle(nullptr), 0);
    if (m_mouseHook) {
        m_logFile << L"Mouse hook installed successfully.\n" << std::flush;
    }
}

void MouseTracker::Stop() {
    if (!m_isRunning) return;

    m_isRunning = false;

    // 唤醒处理线程并等待其结束
    m_queueCondition.notify_all();
    if (m_processingThread.joinable()) {
        m_processingThread.join();
    }

    if (m_mouseHook) {
        UnhookWindowsHookEx(m_mouseHook);
        m_mouseHook = nullptr;
    }

    if (m_logFile.is_open()) {
        m_logFile << L"========== Mouse Tracker Stopped at " << GetCurrentTimeString() << L" ==========\n" << std::flush;
    }
}

LRESULT CALLBACK MouseTracker::MouseHookProc(int nCode, WPARAM wParam, LPARAM lParam) {
    if (nCode >= 0 && s_instance && s_instance->m_isRunning) {
        const MSLLHOOKSTRUCT* mouseInfo = reinterpret_cast<MSLLHOOKSTRUCT*>(lParam);
        
        // 忽略拖动窗口的情况（通过检测是否在非客户区）
        HWND hwnd = WindowFromPoint(mouseInfo->pt);
        if (hwnd) {
            LRESULT hitTest = SendMessage(hwnd, WM_NCHITTEST, 0, 
                                         MAKELPARAM(mouseInfo->pt.x, mouseInfo->pt.y));
            // 如果在标题栏或边框，忽略
            if (hitTest == HTCAPTION || hitTest == HTBORDER || hitTest == HTLEFT || 
                hitTest == HTRIGHT || hitTest == HTTOP || hitTest == HTBOTTOM) {
                return CallNextHookEx(nullptr, nCode, wParam, lParam);
            }
        }

        s_instance->ProcessMouseEvent(wParam, mouseInfo);
    }
    return CallNextHookEx(nullptr, nCode, wParam, lParam);
}

void MouseTracker::ProcessMouseEvent(WPARAM wParam, const MSLLHOOKSTRUCT* mouseInfo) {
    MouseEventType eventType = MouseEventType::UNKNOWN;
    DWORD currentTime = GetTickCount();

    switch (wParam) {
        case WM_LBUTTONDOWN: {
            // 检测双击
            if (currentTime - m_lastClickTime < GetDoubleClickTime() &&
                abs(mouseInfo->pt.x - m_lastClickPos.x) < 5 &&
                abs(mouseInfo->pt.y - m_lastClickPos.y) < 5) {
                eventType = MouseEventType::LEFT_DOUBLE_CLICK;
                m_lastClickTime = 0; // 重置，避免三连击被识别为双击
            } else {
                eventType = MouseEventType::LEFT_CLICK;
                m_lastClickTime = currentTime;
                m_lastClickPos = mouseInfo->pt;
            }
            break;
        }
        case WM_RBUTTONDOWN:
            eventType = MouseEventType::RIGHT_CLICK;
            break;
        default:
            return;
    }

    if (eventType != MouseEventType::UNKNOWN) {
        // 快速入队，不阻塞钩子
        HWND hwnd = WindowFromPoint(mouseInfo->pt);
        PendingMouseEvent event;
        event.eventType = eventType;
        event.position = mouseInfo->pt;
        event.hwnd = hwnd;
        event.timestamp = std::chrono::system_clock::now();

        {
            std::lock_guard<std::mutex> lock(s_instance->m_queueMutex);
            s_instance->m_eventQueue.push(event);
        }
        s_instance->m_queueCondition.notify_one();
    }
}

void MouseTracker::ProcessRecordQueue() {
    // 在工作线程中初始化 COM（每个线程需要单独初始化）
    CoInitializeEx(nullptr, COINIT_MULTITHREADED);

    while (m_isRunning) {
        PendingMouseEvent event;
        
        {
            std::unique_lock<std::mutex> lock(m_queueMutex);
            // 等待队列有数据或停止信号
            m_queueCondition.wait(lock, [this] { 
                return !m_eventQueue.empty() || !m_isRunning; 
            });

            if (!m_isRunning && m_eventQueue.empty()) {
                break;
            }

            if (!m_eventQueue.empty()) {
                event = m_eventQueue.front();
                m_eventQueue.pop();
            } else {
                continue;
            }
        }

        // 在工作线程中处理耗时操作
        RecordMouseOperation(event.eventType, event.position, event.hwnd);
    }

    CoUninitialize();
}

void MouseTracker::RecordMouseOperation(MouseEventType eventType, POINT position, HWND hwnd) {
    MouseOperationRecord record;
    record.timestamp = std::chrono::system_clock::now();
    record.eventType = eventType;
    record.position = position;

    // 获取鼠标位置的窗口信息（这些操作现在在工作线程中执行）
    if (hwnd && IsWindow(hwnd)) {
        record.applicationName = GetApplicationName(hwnd);
        record.windowTitle = GetWindowTitle(hwnd);
        
        // UI Automation 调用（耗时操作）
        try {
            record.content = GetElementContentAtPoint(position);
        } catch (...) {
            record.content = L"[Error getting content]";
        }
    }

    // 添加到记录列表
    {
        std::lock_guard<std::mutex> lock(m_recordsMutex);
        m_records.push_back(record);
        CleanupOldRecords();
    }

    // 打印到控制台（异步，不会阻塞钩子）
    std::wcout << L"\n[" << GetCurrentTimeString() << L"] "
               << L"Event: " << MouseEventTypeToString(eventType) << L"\n"
               << L"Position: (" << position.x << L", " << position.y << L")\n"
               << L"Application: " << record.applicationName << L"\n"
               << L"Window: " << record.windowTitle << L"\n"
               << L"Content: " << record.content << L"\n"
               << L"Element Type: " << record.elementType << L"\n"
               << std::flush;

    // 写入日志文件（异步）
    if (m_logFile.is_open()) {
        m_logFile << record.toJson() << L"\n" << std::flush;
    }
}

std::wstring MouseTracker::GetElementContentAtPoint(POINT pt) {
    if (!m_pAutomation) return L"";

    IUIAutomationElement* element = nullptr;
    
    // 添加超时保护
    HRESULT hr = m_pAutomation->ElementFromPoint(pt, &element);
    
    if (FAILED(hr) || !element) {
        return L"";
    }

    std::wstring content;
    BSTR name = nullptr;
    CONTROLTYPEID controlType;

    // 获取元素名称
    hr = element->get_CurrentName(&name);
    if (SUCCEEDED(hr) && name) {
        content = name;
        SysFreeString(name);
    }

    // 获取控件类型
    element->get_CurrentControlType(&controlType);
    
    // 根据控件类型获取特定信息
    switch (controlType) {
        case UIA_HyperlinkControlTypeId: {
            // 超链接 - 尝试获取 URL
            BSTR url = nullptr;
            IUIAutomationValuePattern* valuePattern = nullptr;
            element->GetCurrentPatternAs(UIA_ValuePatternId, __uuidof(IUIAutomationValuePattern), 
                                        (void**)&valuePattern);
            if (valuePattern) {
                valuePattern->get_CurrentValue(&url);
                if (url) {
                    content += L" [URL: " + std::wstring(url) + L"]";
                    SysFreeString(url);
                }
                valuePattern->Release();
            }
            break;
        }
        case UIA_ButtonControlTypeId: {
            // 按钮
            if (content.empty()) {
                content = L"[Button]";
            }
            break;
        }
        case UIA_TextControlTypeId:
        case UIA_EditControlTypeId: {
            // 文本框 - 尝试获取文本值
            IUIAutomationTextPattern* textPattern = nullptr;
            element->GetCurrentPatternAs(UIA_TextPatternId, __uuidof(IUIAutomationTextPattern), 
                                        (void**)&textPattern);
            if (textPattern) {
                IUIAutomationTextRange* textRange = nullptr;
                textPattern->get_DocumentRange(&textRange);
                if (textRange) {
                    BSTR text = nullptr;
                    textRange->GetText(-1, &text);
                    if (text) {
                        // 如果是选中的文本
                        IUIAutomationTextRangeArray* selections = nullptr;
                        textPattern->GetSelection(&selections);
                        if (selections) {
                            int count = 0;
                            selections->get_Length(&count);
                            if (count > 0) {
                                IUIAutomationTextRange* selection = nullptr;
                                selections->GetElement(0, &selection);
                                if (selection) {
                                    BSTR selectedText = nullptr;
                                    selection->GetText(-1, &selectedText);
                                    if (selectedText && wcslen(selectedText) > 0) {
                                        content = L"[Selected Text: " + std::wstring(selectedText) + L"]";
                                        SysFreeString(selectedText);
                                    }
                                    selection->Release();
                                }
                            }
                            selections->Release();
                        }
                        SysFreeString(text);
                    }
                    textRange->Release();
                }
                textPattern->Release();
            }
            break;
        }
        case UIA_TabItemControlTypeId: {
            // Tab 页
            if (!content.empty()) {
                content = L"[Tab: " + content + L"]";
            }
            break;
        }
    }

    // 存储元素类型
    MouseOperationRecord* lastRecord = nullptr;
    {
        std::lock_guard<std::mutex> lock(m_recordsMutex);
        if (!m_records.empty()) {
            lastRecord = &m_records.back();
        }
    }
    if (lastRecord) {
        lastRecord->elementType = GetElementTypeAtPoint(pt, element);
    }

    element->Release();
    return content;
}

std::wstring MouseTracker::GetElementTypeAtPoint(POINT pt, IUIAutomationElement* element) {
    if (!element) return L"Unknown";

    CONTROLTYPEID controlType;
    element->get_CurrentControlType(&controlType);

    switch (controlType) {
        case UIA_ButtonControlTypeId: return L"Button";
        case UIA_HyperlinkControlTypeId: return L"Hyperlink";
        case UIA_TextControlTypeId: return L"Text";
        case UIA_EditControlTypeId: return L"TextBox";
        case UIA_TabItemControlTypeId: return L"Tab";
        case UIA_MenuItemControlTypeId: return L"MenuItem";
        case UIA_CheckBoxControlTypeId: return L"CheckBox";
        case UIA_RadioButtonControlTypeId: return L"RadioButton";
        case UIA_ComboBoxControlTypeId: return L"ComboBox";
        case UIA_ListItemControlTypeId: return L"ListItem";
        case UIA_ImageControlTypeId: return L"Image";
        default: return L"Unknown";
    }
}

std::wstring MouseTracker::GetApplicationName(HWND hwnd) {
    DWORD processId = 0;
    GetWindowThreadProcessId(hwnd, &processId);

    HANDLE hProcess = OpenProcess(PROCESS_QUERY_INFORMATION | PROCESS_VM_READ, FALSE, processId);
    if (!hProcess) return L"Unknown";

    wchar_t processName[MAX_PATH] = L"";
    DWORD size = MAX_PATH;
    
    if (QueryFullProcessImageName(hProcess, 0, processName, &size)) {
        std::wstring fullPath(processName);
        size_t lastSlash = fullPath.find_last_of(L"\\/");
        if (lastSlash != std::wstring::npos) {
            CloseHandle(hProcess);
            return fullPath.substr(lastSlash + 1);
        }
    }

    CloseHandle(hProcess);
    return L"Unknown";
}

std::wstring MouseTracker::GetWindowTitle(HWND hwnd) {
    wchar_t title[256] = L"";
    GetWindowText(hwnd, title, 256);
    return std::wstring(title);
}

void MouseTracker::CleanupOldRecords() {
    auto now = std::chrono::system_clock::now();
    auto oneHourAgo = now - std::chrono::hours(1);

    m_records.erase(
        std::remove_if(m_records.begin(), m_records.end(),
            [oneHourAgo](const MouseOperationRecord& record) {
                return record.timestamp < oneHourAgo;
            }),
        m_records.end()
    );
}

void MouseTracker::SaveToFile(const std::wstring& filename) {
    std::wofstream file(filename);
    if (!file.is_open()) return;

    file << L"{\n  \"records\": [\n";
    
    std::lock_guard<std::mutex> lock(m_recordsMutex);
    for (size_t i = 0; i < m_records.size(); ++i) {
        file << L"    " << m_records[i].toJson();
        if (i < m_records.size() - 1) {
            file << L",";
        }
        file << L"\n";
    }
    
    file << L"  ]\n}\n";
    file.close();
}

std::wstring MouseTracker::GetAllRecordsAsJson() {
    std::wstringstream ss;
    ss << L"{\n  \"records\": [\n";
    
    std::lock_guard<std::mutex> lock(m_recordsMutex);
    for (size_t i = 0; i < m_records.size(); ++i) {
        ss << L"    " << m_records[i].toJson();
        if (i < m_records.size() - 1) {
            ss << L",";
        }
        ss << L"\n";
    }
    
    ss << L"  ]\n}";
    return ss.str();
}

std::wstring MouseOperationRecord::toJson() const {
    std::wstringstream ss;
    
    // 转换时间戳为字符串
    auto time_t_val = std::chrono::system_clock::to_time_t(timestamp);
    std::tm tm_val;
    localtime_s(&tm_val, &time_t_val);
    
    wchar_t timeStr[100];
    wcsftime(timeStr, 100, L"%Y-%m-%d %H:%M:%S", &tm_val);

    // JSON 转义函数
    auto escapeJson = [](const std::wstring& str) -> std::wstring {
        std::wstring escaped;
        for (wchar_t c : str) {
            switch (c) {
                case L'\\': escaped += L"\\\\"; break;
                case L'\"': escaped += L"\\\""; break;
                case L'\n': escaped += L"\\n"; break;
                case L'\r': escaped += L"\\r"; break;
                case L'\t': escaped += L"\\t"; break;
                default: escaped += c; break;
            }
        }
        return escaped;
    };

    ss << L"{\n"
       << L"      \"timestamp\": \"" << timeStr << L"\",\n"
       << L"      \"eventType\": \"" << MouseEventTypeToString(eventType) << L"\",\n"
       << L"      \"position\": {\"x\": " << position.x << L", \"y\": " << position.y << L"},\n"
       << L"      \"content\": \"" << escapeJson(content) << L"\",\n"
       << L"      \"applicationName\": \"" << escapeJson(applicationName) << L"\",\n"
       << L"      \"windowTitle\": \"" << escapeJson(windowTitle) << L"\",\n"
       << L"      \"elementType\": \"" << escapeJson(elementType) << L"\"\n"
       << L"    }";

    return ss.str();
}

std::wstring MouseEventTypeToString(MouseEventType type) {
    switch (type) {
        case MouseEventType::LEFT_CLICK: return L"LeftClick";
        case MouseEventType::LEFT_DOUBLE_CLICK: return L"DoubleClick";
        case MouseEventType::RIGHT_CLICK: return L"RightClick";
        case MouseEventType::TEXT_SELECTION: return L"TextSelection";
        default: return L"Unknown";
    }
}

std::wstring GetCurrentTimeString() {
    auto now = std::chrono::system_clock::now();
    auto time_t_val = std::chrono::system_clock::to_time_t(now);
    std::tm tm_val;
    localtime_s(&tm_val, &time_t_val);
    
    wchar_t buffer[100];
    wcsftime(buffer, 100, L"%Y-%m-%d %H:%M:%S", &tm_val);
    return std::wstring(buffer);
}
