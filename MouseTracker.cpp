#include "MouseTracker.h"
#include <iostream>
#include <sstream>
#include <iomanip>
#include <psapi.h>
#include <atlbase.h>
#include <UIAutomationClient.h>
#include <ShellScalingApi.h>

#pragma comment(lib, "psapi.lib")
#pragma comment(lib, "Shcore.lib")

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
    // ✅ 设置 DPI 感知：使用 Per-Monitor V2 模式
    // 这样可以确保在高 DPI 显示器上坐标系统正确
    SetProcessDpiAwarenessContext(DPI_AWARENESS_CONTEXT_PER_MONITOR_AWARE_V2);
    
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
        PendingMouseEvent event;
        event.eventType = eventType;
        event.position = mouseInfo->pt;
        
        // ✅ 关键修复：在多显示器环境下，WindowFromPoint 可能返回子窗口，其坐标系统可能不正确
        // 应该获取顶层窗口，而不是子窗口
        HWND pointWindow = WindowFromPoint(mouseInfo->pt);
        HWND foregroundWindow = GetForegroundWindow();
        
        // 获取 pointWindow 的顶层父窗口
        HWND topLevelWindow = pointWindow;
        if (pointWindow) {
            HWND parent = pointWindow;
            while (parent) {
                HWND nextParent = GetParent(parent);
                if (!nextParent) {
                    topLevelWindow = parent;
                    break;
                }
                parent = nextParent;
            }
        }
        
        // 验证顶层窗口是否与前台窗口一致
        bool useForeground = (topLevelWindow != foregroundWindow);
        
        // 优先使用前台窗口（更可靠），除非 topLevelWindow 确实包含点击坐标
        if (topLevelWindow && IsWindow(topLevelWindow)) {
            RECT rect;
            if (GetWindowRect(topLevelWindow, &rect)) {
                POINT pt = mouseInfo->pt;
                if (pt.x >= rect.left && pt.x < rect.right && 
                    pt.y >= rect.top && pt.y < rect.bottom) {
                    useForeground = false;  // 坐标在范围内，使用 topLevelWindow
                }
            }
        }
        
        event.pointWindow = useForeground ? foregroundWindow : topLevelWindow;
        
        // 调试：输出点击信息
        #ifdef _DEBUG
        wchar_t className[256] = {0};
        wchar_t topClassName[256] = {0};
        wchar_t fgClassName[256] = {0};
        if (pointWindow && IsWindow(pointWindow)) {
            GetClassNameW(pointWindow, className, 256);
        }
        if (topLevelWindow && IsWindow(topLevelWindow)) {
            GetClassNameW(topLevelWindow, topClassName, 256);
        }
        if (foregroundWindow && IsWindow(foregroundWindow)) {
            GetClassNameW(foregroundWindow, fgClassName, 256);
        }
        std::wcout << L"[HOOK] Click at (" << mouseInfo->pt.x << L", " << mouseInfo->pt.y << L")\n"
                   << L"  PointWindow: " << pointWindow << L" Class: " << className << L"\n"
                   << L"  TopLevelWindow: " << topLevelWindow << L" Class: " << topClassName << L"\n"
                   << L"  ForegroundWindow: " << foregroundWindow << L" Class: " << fgClassName << L"\n"
                   << L"  Using: " << (useForeground ? L"ForegroundWindow" : L"TopLevelWindow") << L"\n";
        
        // 输出显示器信息
        RECT topRect = {0};
        if (event.pointWindow && IsWindow(event.pointWindow)) {
            GetWindowRect(event.pointWindow, &topRect);
            std::wcout << L"  TargetWindow Rect: (" << topRect.left << L", " << topRect.top 
                       << L") - (" << topRect.right << L", " << topRect.bottom << L")\n";
            
            bool isInside = (mouseInfo->pt.x >= topRect.left && mouseInfo->pt.x < topRect.right &&
                           mouseInfo->pt.y >= topRect.top && mouseInfo->pt.y < topRect.bottom);
            std::wcout << L"  Click is " << (isInside ? L"INSIDE" : L"OUTSIDE") 
                       << L" target window bounds\n";
        }
        #endif
        
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
        RecordMouseOperation(event.eventType, event.position, event.pointWindow);
    }

    CoUninitialize();
}

void MouseTracker::RecordMouseOperation(MouseEventType eventType, POINT position, HWND pointWindow) {
    MouseOperationRecord record;
    record.timestamp = std::chrono::system_clock::now();
    record.eventType = eventType;
    record.position = position;

    // ✅ 关键改进：先立即获取元素内容（在UI状态改变之前）
    // 不要延迟，否则UI可能已经更新，元素内容会改变
    ElementInfo contentInfo;
    try {
        contentInfo = GetElementContentAtPoint(position, pointWindow);
    } catch (...) {
        contentInfo.content = L"[Error getting content]";
        contentInfo.elementType = L"Unknown";
    }

    // 然后延迟获取前台窗口（用于应用名称和窗口标题）
    // 这个延迟只影响窗口识别，不影响内容获取
    Sleep(50);  // 50ms 延迟足够窗口切换完成

    // 现在获取前台窗口，此时应该已经切换到新窗口了
    HWND foregroundWindow = GetForegroundWindow();
    
    // 调试输出：对比坐标窗口和前台窗口
    #ifdef _DEBUG
    if (pointWindow && IsWindow(pointWindow)) {
        HWND pointRoot = GetRootOwnerWindow(pointWindow);
        std::wstring pointApp = GetApplicationName(pointRoot);
        std::wstring foreApp = GetApplicationName(GetRootOwnerWindow(foregroundWindow));
        std::wcout << L"[DEBUG] PointWindow: " << pointApp 
                   << L", ForegroundWindow: " << foreApp << L"\n";
    }
    #endif

    // 使用前台活动窗口来确定应用程序（更准确）
    if (foregroundWindow && IsWindow(foregroundWindow)) {
        // 获取顶层窗口（避免子窗口导致的错误）
        HWND rootWindow = GetRootOwnerWindow(foregroundWindow);
        
        record.applicationName = GetApplicationName(rootWindow);
        record.windowTitle = GetWindowTitle(rootWindow);
        
        // 使用之前立即获取的内容（在延迟之前获取的）
        record.content = contentInfo.content;
        record.elementType = contentInfo.elementType;
    } else {
        // 降级处理：如果前台窗口无效，使用坐标窗口
        if (pointWindow && IsWindow(pointWindow)) {
            HWND rootWindow = GetRootOwnerWindow(pointWindow);
            record.applicationName = GetApplicationName(rootWindow);
            record.windowTitle = GetWindowTitle(rootWindow);
            
            // 使用之前立即获取的内容
            record.content = contentInfo.content;
            record.elementType = contentInfo.elementType;
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

MouseTracker::ElementInfo MouseTracker::GetElementContentAtPoint(POINT pt, HWND targetWindow) {
    ElementInfo result;
    result.content = L"";
    result.elementType = L"Unknown";
    
    if (!m_pAutomation) return result;

    // ✅ 调试：输出原始点击坐标和窗口信息
    std::wcout << L"[DEBUG] ===== GetElementContentAtPoint called =====\n";
    std::wcout << L"[DEBUG] Click coordinates: (" << pt.x << L", " << pt.y << L")\n";
    std::wcout << L"[DEBUG] Target window handle: " << targetWindow << L"\n";

    // 输出目标窗口的详细信息
    if (targetWindow && IsWindow(targetWindow)) {
        wchar_t windowTitle[256] = {0};
        GetWindowTextW(targetWindow, windowTitle, 256);
        wchar_t className[256] = {0};
        GetClassNameW(targetWindow, className, 256);
        std::wcout << L"[DEBUG] Target window title: \"" << windowTitle << L"\"\n";
        std::wcout << L"[DEBUG] Target window class: \"" << className << L"\"\n";
    }

    // ✅ 关键修复：使用传入的 targetWindow（点击时的窗口），而不是 GetForegroundWindow
    // 因为这个函数在延迟之前调用，此时前台窗口可能还没有切换
    HWND hwnd = targetWindow;
    if (!hwnd || !IsWindow(hwnd)) {
        std::wcout << L"[DEBUG] Target window invalid, using foreground window as fallback\n";
        hwnd = GetForegroundWindow();
    }
    
    std::wcout << L"[DEBUG] Using window handle: " << hwnd << L"\n";
    
    if (!hwnd || !IsWindow(hwnd)) {
        std::wcout << L"[DEBUG] Invalid window handle!\n";
        return result;
    }
    
    // ✅ 检查窗口位置和大小
    RECT windowRect;
    GetWindowRect(hwnd, &windowRect);
    std::wcout << L"[DEBUG] Window rect (screen coordinates): (" 
               << windowRect.left << L", " << windowRect.top << L") - (" 
               << windowRect.right << L", " << windowRect.bottom << L")\n";
    
    // ✅ 检查点击位置是否在窗口内
    bool isInsideWindow = (pt.x >= windowRect.left && pt.x <= windowRect.right &&
                          pt.y >= windowRect.top && pt.y <= windowRect.bottom);
    std::wcout << L"[DEBUG] Click is " << (isInsideWindow ? L"INSIDE" : L"OUTSIDE") 
               << L" window rect (expected: INSIDE)\n";
    
    IUIAutomationElement* rootElement = nullptr;
    HRESULT hr = m_pAutomation->ElementFromHandle(hwnd, &rootElement);
    
    if (FAILED(hr) || !rootElement) {
        std::wcout << L"[DEBUG] Failed to get root element from window handle\n";
        return result;
    }
    
    std::wcout << L"[DEBUG] Got root element, searching for content area...\n";
    
    // ✅ 关键：先找到内容区域（Document 或合适的 Pane），像 BrowserContentExtractor 那样
    IUIAutomationElement* contentArea = FindContentArea(rootElement);
    
    IUIAutomationElement* searchRoot = contentArea ? contentArea : rootElement;
    
    if (contentArea) {
        std::wcout << L"[DEBUG] Found content area, searching within it\n";
    } else {
        std::wcout << L"[DEBUG] No content area found, searching entire window\n";
    }
    
    // 获取 TreeWalker
    IUIAutomationTreeWalker* walker = nullptr;
    hr = m_pAutomation->get_RawViewWalker(&walker);
    
    if (FAILED(hr) || !walker) {
        if (contentArea) contentArea->Release();
        rootElement->Release();
        return result;
    }
    
    // 在内容区域中查找包含点击坐标的元素
    std::wcout << L"[DEBUG] Searching for element at point (" << pt.x << L", " << pt.y << L")\n";
    
    IUIAutomationElement* targetElement = FindElementAtPointInTree(searchRoot, pt, walker, 0);
    
    // ✅ 如果在内容区域中找不到，尝试在整个窗口中查找
    if (!targetElement && contentArea) {
        std::wcout << L"[DEBUG] Not found in content area, searching entire window...\n";
        targetElement = FindElementAtPointInTree(rootElement, pt, walker, 0);
    }
    
    walker->Release();
    
    if (targetElement) {
        std::wcout << L"[DEBUG] Found target element in tree!\n";
        
        // 获取元素信息
        CONTROLTYPEID controlType;
        targetElement->get_CurrentControlType(&controlType);
        result.elementType = GetElementTypeString(controlType);
        
        // 调试：输出元素信息
        BSTR debugName = nullptr;
        BSTR debugClassName = nullptr;
        targetElement->get_CurrentName(&debugName);
        targetElement->get_CurrentClassName(&debugClassName);
        
        std::wcout << L"[DEBUG] Target element info:\n"
                   << L"  Type: " << result.elementType << L"\n"
                   << L"  Name: " << (debugName ? debugName : L"(empty)") << L"\n"
                   << L"  ClassName: " << (debugClassName ? debugClassName : L"(empty)") << L"\n";
        
        if (debugName) SysFreeString(debugName);
        if (debugClassName) SysFreeString(debugClassName);
        
        // 获取内容
        result.content = TryGetElementContent(targetElement, controlType);
        
        if (result.content.empty()) {
            // 如果当前元素没内容，递归查找子元素
            std::wcout << L"[DEBUG] Target element has no content, searching children...\n";
            IUIAutomationTreeWalker* contentWalker = nullptr;
            m_pAutomation->get_RawViewWalker(&contentWalker);
            if (contentWalker) {
                result.content = TraverseForContent(targetElement, contentWalker, 0, 3);
                contentWalker->Release();
            }
        }
        
        targetElement->Release();
    } else {
        std::wcout << L"[DEBUG] Could not find target element in tree, trying ElementFromPoint as fallback...\n";
        
        // ✅ 后备方案：使用 ElementFromPoint 直接获取
        IUIAutomationElement* pointElement = nullptr;
        hr = m_pAutomation->ElementFromPoint(pt, &pointElement);
        
        if (SUCCEEDED(hr) && pointElement) {
            std::wcout << L"[DEBUG] ElementFromPoint succeeded!\n";
            
            CONTROLTYPEID controlType;
            pointElement->get_CurrentControlType(&controlType);
            result.elementType = GetElementTypeString(controlType);
            
            // 调试信息
            BSTR debugName = nullptr;
            pointElement->get_CurrentName(&debugName);
            std::wcout << L"[DEBUG] ElementFromPoint - Type: " << result.elementType 
                       << L", Name: " << (debugName ? debugName : L"(empty)") << L"\n";
            if (debugName) SysFreeString(debugName);
            
            // 获取内容
            result.content = TryGetElementContent(pointElement, controlType);
            
            if (result.content.empty()) {
                IUIAutomationTreeWalker* contentWalker = nullptr;
                m_pAutomation->get_RawViewWalker(&contentWalker);
                if (contentWalker) {
                    result.content = TraverseForContent(pointElement, contentWalker, 0, 3);
                    contentWalker->Release();
                }
            }
            
            pointElement->Release();
        } else {
            std::wcout << L"[DEBUG] ElementFromPoint also failed\n";
        }
    }
    
    if (contentArea) contentArea->Release();
    rootElement->Release();
    
    if (result.content.empty()) {
        result.content = L"[No Content Found]";
    }
    
    return result;
}

// 新增：查找内容区域（类似 BrowserContentExtractor::FindDocumentElement）
IUIAutomationElement* MouseTracker::FindContentArea(IUIAutomationElement* rootElement) {
    if (!m_pAutomation || !rootElement) {
        return nullptr;
    }
    
    // 1. 首先尝试查找 Document 控件（适用于浏览器）
    IUIAutomationCondition* condition = nullptr;
    VARIANT varProp;
    varProp.vt = VT_I4;
    varProp.lVal = UIA_DocumentControlTypeId;
    
    HRESULT hr = m_pAutomation->CreatePropertyCondition(UIA_ControlTypePropertyId, varProp, &condition);
    
    if (SUCCEEDED(hr) && condition) {
        IUIAutomationElement* docElement = nullptr;
        hr = rootElement->FindFirst(TreeScope_Descendants, condition, &docElement);
        condition->Release();
        
        if (SUCCEEDED(hr) && docElement) {
            std::wcout << L"[DEBUG] Found Document element (content area)\n";
            return docElement;
        }
    }
    
    // 2. 如果没找到 Document，查找合适的 Pane（适用于 Teams 等应用）
    varProp.lVal = UIA_PaneControlTypeId;
    hr = m_pAutomation->CreatePropertyCondition(UIA_ControlTypePropertyId, varProp, &condition);
    
    if (SUCCEEDED(hr) && condition) {
        IUIAutomationElementArray* paneArray = nullptr;
        hr = rootElement->FindAll(TreeScope_Descendants, condition, &paneArray);
        condition->Release();
        
        if (SUCCEEDED(hr) && paneArray) {
            int length = 0;
            paneArray->get_Length(&length);
            
            std::wcout << L"[DEBUG] Found " << length << L" Pane elements, checking for content area...\n";
            
            // 遍历所有 Pane，找到最有可能是内容区的那个
            for (int i = 0; i < length && i < 20; i++) {
                IUIAutomationElement* pane = nullptr;
                if (SUCCEEDED(paneArray->GetElement(i, &pane)) && pane) {
                    // 检查 Name 和 AutomationId，排除工具栏、书签栏等
                    BSTR name = nullptr;
                    BSTR automationId = nullptr;
                    pane->get_CurrentName(&name);
                    pane->get_CurrentAutomationId(&automationId);
                    
                    std::wstring nameStr = name ? name : L"";
                    std::wstring idStr = automationId ? automationId : L"";
                    
                    if (name) SysFreeString(name);
                    if (automationId) SysFreeString(automationId);
                    
                    // 排除明显不是内容区的 Pane
                    bool isExcluded = 
                        nameStr.find(L"Toolbar") != std::wstring::npos ||
                        nameStr.find(L"Bookmark") != std::wstring::npos ||
                        nameStr.find(L"Tab Bar") != std::wstring::npos ||
                        nameStr.find(L"Navigation") != std::wstring::npos ||
                        idStr.find(L"Toolbar") != std::wstring::npos ||
                        idStr.find(L"TabBar") != std::wstring::npos;
                    
                    if (!isExcluded) {
                        std::wcout << L"[DEBUG] Found potential content Pane: " << nameStr << L"\n";
                        paneArray->Release();
                        return pane;
                    }
                    
                    pane->Release();
                }
            }
            
            paneArray->Release();
        }
    }
    
    // 3. 如果都没找到，返回 nullptr（使用根元素）
    std::wcout << L"[DEBUG] No specific content area found\n";
    return nullptr;
}

// 新增：在元素树中查找包含指定坐标的元素（返回最小的匹配元素）
IUIAutomationElement* MouseTracker::FindElementAtPointInTree(IUIAutomationElement* element, POINT pt, IUIAutomationTreeWalker* walker, int depth) {
    if (!element || !walker || depth > 15) {
        return nullptr;
    }
    
    // 检查当前元素的边界矩形
    RECT rect;
    HRESULT hr = element->get_CurrentBoundingRectangle(&rect);
    
    if (FAILED(hr)) {
        std::wcout << L"[DEBUG] Depth " << depth << L": Failed to get bounding rectangle\n";
        return nullptr;
    }
    
    // 调试：输出元素边界和点击位置
    if (depth == 0) {
        std::wcout << L"[DEBUG] Root element rect: (" << rect.left << L", " << rect.top 
                   << L") - (" << rect.right << L", " << rect.bottom << L")\n";
        std::wcout << L"[DEBUG] Click point: (" << pt.x << L", " << pt.y << L")\n";
    }
    
    // 如果点不在当前元素内，返回 null
    if (pt.x < rect.left || pt.x > rect.right || pt.y < rect.top || pt.y > rect.bottom) {
        if (depth == 0) {
            std::wcout << L"[DEBUG] Click point is OUTSIDE root element bounds!\n";
        }
        return nullptr;
    }
    
    if (depth == 0) {
        std::wcout << L"[DEBUG] Click point is inside root element, starting traversal...\n";
    }
    
    // ✅ 关键改进：点在当前元素内，先检查当前元素是否有文本内容
    CONTROLTYPEID controlType;
    element->get_CurrentControlType(&controlType);
    std::wstring currentContent = TryGetElementContent(element, controlType);
    
    if (!currentContent.empty()) {
        // 当前元素有内容，但还要检查是否有更精确的子元素也有内容
        std::wcout << L"[DEBUG] Element at depth " << depth << " has content: " << currentContent.substr(0, 30) << L"...\n";
    }
    
    // 继续查找子元素，看是否有更精确（面积更小）且有内容的子元素
    IUIAutomationElement* child = nullptr;
    hr = walker->GetFirstChildElement(element, &child);
    
    IUIAutomationElement* bestMatch = nullptr;
    LONG bestArea = LONG_MAX;
    bool bestHasContent = false;
    
    while (SUCCEEDED(hr) && child) {
        // 递归查找子元素
        IUIAutomationElement* childMatch = FindElementAtPointInTree(child, pt, walker, depth + 1);
        
        if (childMatch) {
            // 检查这个子元素是否有内容
            CONTROLTYPEID childType;
            childMatch->get_CurrentControlType(&childType);
            std::wstring childContent = TryGetElementContent(childMatch, childType);
            bool childHasContent = !childContent.empty();
            
            // 计算面积
            RECT childRect;
            if (SUCCEEDED(childMatch->get_CurrentBoundingRectangle(&childRect))) {
                LONG area = (childRect.right - childRect.left) * (childRect.bottom - childRect.top);
                
                // 优先选择有内容的元素，其次选择面积更小的元素
                bool isBetter = false;
                if (childHasContent && !bestHasContent) {
                    isBetter = true;  // 有内容的优于没内容的
                } else if (childHasContent == bestHasContent && area > 0 && area < bestArea) {
                    isBetter = true;  // 同样有/没有内容，选择面积更小的
                }
                
                if (isBetter) {
                    if (bestMatch) bestMatch->Release();
                    bestMatch = childMatch;
                    bestArea = area;
                    bestHasContent = childHasContent;
                } else {
                    childMatch->Release();
                }
            } else {
                childMatch->Release();
            }
        }
        
        // 获取下一个兄弟
        IUIAutomationElement* next = nullptr;
        hr = walker->GetNextSiblingElement(child, &next);
        child->Release();
        child = next;
    }
    
    // ✅ 决策逻辑：
    // 1. 如果找到有内容的子元素，返回它
    // 2. 如果当前元素有内容但没找到有内容的子元素，返回当前元素
    // 3. 如果都没内容，返回面积最小的子元素或当前元素
    if (bestMatch && bestHasContent) {
        // 找到了有内容的子元素
        std::wcout << L"[DEBUG] Returning child element with content\n";
        return bestMatch;
    } else if (!currentContent.empty()) {
        // 当前元素有内容，没找到更好的子元素
        std::wcout << L"[DEBUG] Returning current element with content\n";
        if (bestMatch) bestMatch->Release();
        element->AddRef();
        return element;
    } else if (bestMatch) {
        // 都没内容，返回最小的子元素
        std::wcout << L"[DEBUG] Returning smallest child element (no content found)\n";
        return bestMatch;
    } else {
        // 没有子元素，返回当前元素
        std::wcout << L"[DEBUG] Returning current element (no children)\n";
        element->AddRef();
        return element;
    }
}

// 新增：递归遍历元素树查找内容（类似 BrowserContentExtractor::TraverseElementTree）
std::wstring MouseTracker::TraverseForContent(IUIAutomationElement* element, IUIAutomationTreeWalker* walker, int depth, int maxDepth) {
    if (!element || !walker || depth > maxDepth) {
        return L"";
    }
    
    // 先尝试当前元素
    CONTROLTYPEID controlType;
    element->get_CurrentControlType(&controlType);
    std::wstring content = TryGetElementContent(element, controlType);
    
    if (!content.empty()) {
        std::wcout << L"[DEBUG] Found content at depth " << depth << L": " << content.substr(0, 50) << L"...\n";
        return content;
    }
    
    // 递归遍历子元素
    IUIAutomationElement* child = nullptr;
    HRESULT hr = walker->GetFirstChildElement(element, &child);
    
    while (SUCCEEDED(hr) && child) {
        std::wstring childContent = TraverseForContent(child, walker, depth + 1, maxDepth);
        
        if (!childContent.empty()) {
            child->Release();
            return childContent;
        }
        
        // 获取下一个兄弟元素
        IUIAutomationElement* next = nullptr;
        hr = walker->GetNextSiblingElement(child, &next);
        child->Release();
        child = next;
    }
    
    return L"";
}

// 新增辅助函数：尝试从元素获取内容（封装所有获取方法）
std::wstring MouseTracker::TryGetElementContent(IUIAutomationElement* element, CONTROLTYPEID controlType) {
    if (!element) return L"";
    
    std::wstring result;
    HRESULT hr;
    
    // 1. 首先尝试获取 Name 属性
    BSTR name = nullptr;
    if (SUCCEEDED(element->get_CurrentName(&name)) && name) {
        std::wstring nameStr = name;
        SysFreeString(name);
        if (!nameStr.empty()) {
            result = nameStr;
            
            // 对于超链接，尝试附加 URL
            if (controlType == UIA_HyperlinkControlTypeId) {
                IUIAutomationValuePattern* valuePattern = nullptr;
                if (SUCCEEDED(element->GetCurrentPatternAs(UIA_ValuePatternId, 
                    __uuidof(IUIAutomationValuePattern), (void**)&valuePattern)) && valuePattern) {
                    BSTR url = nullptr;
                    if (SUCCEEDED(valuePattern->get_CurrentValue(&url)) && url) {
                        std::wstring urlStr = url;
                        SysFreeString(url);
                        if (!urlStr.empty()) {
                            result += L" → " + urlStr;
                        }
                    }
                    valuePattern->Release();
                }
            }
            
            return result;
        }
    }
    
    // 2. 尝试 ValuePattern（适用于编辑框、输入框等）
    IUIAutomationValuePattern* valuePattern = nullptr;
    if (SUCCEEDED(element->GetCurrentPatternAs(UIA_ValuePatternId, 
        __uuidof(IUIAutomationValuePattern), (void**)&valuePattern)) && valuePattern) {
        BSTR value = nullptr;
        if (SUCCEEDED(valuePattern->get_CurrentValue(&value)) && value) {
            std::wstring valueStr = value;
            SysFreeString(value);
            if (!valueStr.empty()) {
                result = valueStr;
                valuePattern->Release();
                return result;
            }
        }
        valuePattern->Release();
    }
    
    // 3. 尝试 TextPattern（适用于文本内容、文档等）
    IUIAutomationTextPattern* textPattern = nullptr;
    if (SUCCEEDED(element->GetCurrentPatternAs(UIA_TextPatternId, 
        __uuidof(IUIAutomationTextPattern), (void**)&textPattern)) && textPattern) {
        IUIAutomationTextRange* textRange = nullptr;
        if (SUCCEEDED(textPattern->get_DocumentRange(&textRange)) && textRange) {
            BSTR text = nullptr;
            if (SUCCEEDED(textRange->GetText(-1, &text)) && text) {
                std::wstring textStr = text;
                SysFreeString(text);
                if (!textStr.empty()) {
                    result = textStr;
                    textRange->Release();
                    textPattern->Release();
                    return result;
                }
            }
            textRange->Release();
        }
        textPattern->Release();
    }
    
    // 4. 尝试 HelpText 作为后备
    BSTR helpText = nullptr;
    if (SUCCEEDED(element->get_CurrentHelpText(&helpText)) && helpText) {
        std::wstring helpStr = helpText;
        SysFreeString(helpText);
        if (!helpStr.empty()) {
            result = helpStr;
            return result;
        }
    }
    
    // 5. 最后尝试 AutomationId
    //BSTR autoId = nullptr;
    //if (SUCCEEDED(element->get_CurrentAutomationId(&autoId)) && autoId) {
    //    std::wstring idStr = autoId;
    //    SysFreeString(autoId);
    //    if (!idStr.empty()) {
    //        result = L"[ID: " + idStr + L"]";
    //        return result;
    //    }
    //}
    
    return result;  // 返回空字符串表示未找到内容
}

std::wstring MouseTracker::GetElementTypeString(CONTROLTYPEID controlType) {
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

HWND MouseTracker::GetRootOwnerWindow(HWND hwnd) {
    if (!hwnd || !IsWindow(hwnd)) {
        return nullptr;
    }
    
    // 获取顶层所有者窗口
    HWND rootWindow = GetAncestor(hwnd, GA_ROOTOWNER);
    if (!rootWindow) {
        rootWindow = hwnd;
    }
    
    return rootWindow;
}

std::wstring MouseTracker::GetApplicationName(HWND hwnd) {
    if (!hwnd || !IsWindow(hwnd)) {
        return L"Unknown";
    }
    
    DWORD processId = 0;
    GetWindowThreadProcessId(hwnd, &processId);
    
    if (processId == 0) {
        return L"Unknown";
    }

    HANDLE hProcess = OpenProcess(PROCESS_QUERY_INFORMATION | PROCESS_VM_READ, FALSE, processId);
    if (!hProcess) {
        return L"Unknown";
    }

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
    if (!hwnd || !IsWindow(hwnd)) {
        return L"";
    }
    
    wchar_t title[512] = L"";
    int length = GetWindowText(hwnd, title, 512);
    
    if (length > 0) {
        return std::wstring(title);
    }
    
    // 如果窗口标题为空，尝试获取类名
    wchar_t className[256] = L"";
    if (GetClassName(hwnd, className, 256) > 0) {
        return std::wstring(L"[") + className + L"]";
    }
    
    return L"";
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
