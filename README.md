# Windows 鼠标操作追踪器 (Mouse Content Tracker)

## 项目简介

这是一个用于记录 Windows 系统上用户鼠标操作行踪的 C++ 应用程序。程序可以捕获并记录过去一个小时内的所有鼠标交互操作，包括点击位置、交互内容和所属应用程序等详细信息。

## 主要功能

### 1. 鼠标事件捕获
- ✅ **单击检测**: 捕获鼠标左键单击事件
- ✅ **双击检测**: 智能识别双击操作
- ✅ **右键检测**: 捕获鼠标右键点击事件
- ✅ **文本选择**: 识别用户选择的文本内容
- ✅ **智能过滤**: 自动忽略拖动窗口的操作

### 2. 内容识别
程序可以识别鼠标点击位置的各种元素内容：
- 🔗 **超链接**: 获取链接地址和链接文本
- 📑 **Tab 页**: 识别 Tab 页名称
- 🔘 **按钮**: 获取按钮名称和标签
- 📝 **文本**: 捕获选中的文本内容
- 🖥️ **应用程序**: 记录所属应用程序名称
- 🪟 **窗口**: 记录窗口标题
- 🎯 **元素类型**: 识别交互元素的类型（按钮、链接、文本框等）

### 3. 数据存储
- 📊 **JSON 格式**: 所有记录以 JSON 格式存储
- 💾 **实时日志**: 自动写入本地日志文件
- 🖨️ **控制台输出**: 实时打印操作记录到控制台
- ⏱️ **自动清理**: 自动删除超过 1 小时的旧记录

## 技术特性

- **Windows API**: 使用低级鼠标钩子 (WH_MOUSE_LL) 捕获全局鼠标事件
- **UI Automation**: 使用 Microsoft UI Automation 获取界面元素信息
- **线程安全**: 使用互斥锁保护共享数据
- **内存管理**: 智能指针和 RAII 确保资源正确释放
- **Unicode 支持**: 完整支持中文和其他 Unicode 字符

## 编译要求

### 系统要求
- Windows 7 或更高版本
- Visual Studio 2017 或更高版本 (推荐 Visual Studio 2019/2022)
- CMake 3.15 或更高版本

### 依赖库
以下库将自动链接：
- `oleacc.lib` - Active Accessibility
- `psapi.lib` - Process Status API
- `ole32.lib` - COM 库
- `oleaut32.lib` - OLE Automation
- `uuid.lib` - GUID 支持

## 编译步骤

### 方法 1: 使用 CMake (推荐)

```bash
# 创建 build 目录
mkdir build
cd build

# 生成 Visual Studio 项目
cmake ..

# 编译项目
cmake --build . --config Release
```

### 方法 2: 使用 Visual Studio

```bash
# 生成 Visual Studio 解决方案
mkdir build
cd build
cmake -G "Visual Studio 16 2019" ..

# 然后在 Visual Studio 中打开生成的 .sln 文件进行编译
```

### 方法 3: 直接使用 Visual Studio

1. 创建新的 C++ 控制台项目
2. 添加源文件：`main.cpp`, `MouseTracker.cpp`, `MouseTracker.h`
3. 项目属性设置：
   - **字符集**: Unicode
   - **C++ 标准**: C++17 或更高
   - **附加依赖项**: `oleacc.lib psapi.lib ole32.lib oleaut32.lib uuid.lib`

## 使用说明

### 启动程序

```bash
# 以管理员权限运行
.\MouseContentTracker.exe
```

⚠️ **重要**: 程序需要管理员权限才能安装全局鼠标钩子！

### 运行时操作

程序启动后，支持以下交互命令：

- **按 's' + Enter**: 保存当前所有记录到 JSON 文件
- **按 'p' + Enter**: 在控制台打印所有记录（JSON 格式）
- **按 'q' + Enter**: 退出程序

### 输出文件

1. **日志文件**: `mouse_operations_log.txt`
   - 实时记录所有操作
   - 追加模式，不会覆盖旧数据

2. **JSON 记录文件**: `mouse_records_[时间戳].json`
   - 手动保存时生成
   - 包含完整的 JSON 格式记录

## JSON 数据格式

每条记录包含以下字段：

```json
{
  "records": [
    {
      "timestamp": "2025-10-21 14:30:45",
      "eventType": "LeftClick",
      "position": {"x": 520, "y": 340},
      "content": "确定",
      "applicationName": "chrome.exe",
      "windowTitle": "Google Chrome",
      "elementType": "Button"
    }
  ]
}
```

### 字段说明

| 字段 | 类型 | 说明 |
|------|------|------|
| `timestamp` | String | 操作时间戳 |
| `eventType` | String | 事件类型 (LeftClick/DoubleClick/RightClick/TextSelection) |
| `position` | Object | 鼠标位置坐标 {x, y} |
| `content` | String | 交互的具体内容（按钮名、链接、文本等） |
| `applicationName` | String | 所属应用程序名称 |
| `windowTitle` | String | 窗口标题 |
| `elementType` | String | 元素类型（Button/Hyperlink/Tab/TextBox等） |

## 示例输出

### 控制台输出示例

```
[2025-10-21 14:30:45] Event: LeftClick
Position: (520, 340)
Application: chrome.exe
Window: Google Chrome
Content: [Tab: 新标签页]
Element Type: Tab

[2025-10-21 14:30:52] Event: LeftClick
Position: (650, 180)
Application: chrome.exe
Window: GitHub
Content: [Button: Sign in]
Element Type: Button

[2025-10-21 14:31:05] Event: RightClick
Position: (420, 280)
Application: notepad.exe
Window: 无标题 - 记事本
Content: [Selected Text: Hello World]
Element Type: Text
```

## 注意事项

1. **管理员权限**: 程序必须以管理员身份运行才能安装全局钩子
2. **性能影响**: 全局钩子会轻微影响系统性能，建议按需使用
3. **隐私安全**: 程序会记录所有鼠标操作，请妥善保管生成的日志文件
4. **兼容性**: 某些应用程序可能使用自定义界面，内容识别可能不完整
5. **内存管理**: 程序自动清理 1 小时前的记录以控制内存使用

## 故障排除

### 问题: 程序启动失败

**解决方案**:
- 确保以管理员权限运行
- 检查是否已安装 Visual C++ Redistributable

### 问题: 无法识别某些元素内容

**原因**: 某些应用程序使用自定义渲染或不支持 UI Automation

**解决方案**: 这是正常现象，程序会尽可能获取可用信息

### 问题: 编译错误

**解决方案**:
- 确保 C++ 标准设置为 C++17 或更高
- 检查是否正确链接所有必需的库文件
- 确保使用 Unicode 字符集

## 许可证

本项目仅供学习和研究使用。

## 免责声明

使用本软件记录他人操作可能涉及隐私问题，请遵守当地法律法规，仅在合法授权的情况下使用。
