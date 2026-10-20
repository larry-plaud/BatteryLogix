# BatteryLogix

> 充放电 LOG 批量转 Excel 工具

BatteryLogix 是一款 Windows 桌面小工具，用于把电池充放电测试产生的 `.log` 文本日志**批量、递归地**解析成 Excel（`.xlsx`）表格。选定一个文件夹后，程序会扫描其下所有子目录中的 LOG 文件，从每条记录中提取电压、电量、电流、温度等关键字段，并在**每个 LOG 文件所在目录**生成一个同名的 Excel 文档。

发布产物为**自包含单文件 EXE**，下载后双击即可运行，无需安装 .NET 运行时。

---

## 目录

- [功能特性](#功能特性)
- [界面预览](#界面预览)
- [快速开始](#快速开始)
- [使用说明](#使用说明)
- [LOG 解析规则](#log-解析规则)
- [输出格式](#输出格式)
- [从源码构建](#从源码构建)
- [自动发布（CI）](#自动发布ci)
- [项目结构](#项目结构)
- [技术说明](#技术说明)
- [常见问题](#常见问题)
- [许可证](#许可证)

---

## 功能特性

- 🗂️ **递归批量转换** —— 选定一个文件夹，自动扫描所有子目录中的 `*.log` 文件。
- 🚀 **零依赖生成 Excel** —— 直接按 OOXML 规范写出 `.xlsx`，不依赖 Office，也不引用任何第三方库。
- 📊 **结构化字段提取** —— 从日志中提取时间、电压、电量（两路）、电流、温度。
- 🧵 **后台转换 + 实时进度** —— 转换在后台线程执行，界面实时显示进度、记录数与失败数，全程不卡顿。
- 🎯 **就地输出** —— Excel 与源 LOG 同名、同目录，便于对照归档。
- 🛡️ **原子写入** —— 先写临时文件再整体替换，转换中断不会留下损坏的表格。
- 📦 **单文件绿色版** —— 自包含发布，双击即用，无需安装运行时。

## 界面预览

主界面为深色卡片式布局，分为四个区域：

- **文件夹与操作**：显示当前文件夹路径，提供「选择文件夹」「开始转换」两个按钮。
- **提取规则**：展示当前从日志中提取的字段。
- **扫描统计**：实时显示 LOG 文件数、已提取的数据记录数、失败数。
- **处理进度**：进度条与当前正在处理的文件名。

窗口标题：`BatteryLogix - 充放电LOG 批量转 Excel`

## 快速开始

### 方式一：直接下载运行（推荐）

1. 前往 [Releases](https://github.com/larry-lvjin/BatteryLogix/releases) 页面。
2. 下载最新的 `BatteryLogixV<版本号>.exe`。
3. 双击运行，无需安装。

> 系统要求：Windows 10/11 x64。

### 方式二：从源码运行

```bash
git clone https://github.com/larry-lvjin/BatteryLogix.git
cd BatteryLogix/src
dotnet run
```

> 需要安装 [.NET 10 SDK](https://dotnet.microsoft.com/)（Windows Desktop 工作负载）。

## 使用说明

1. 点击 **「选择文件夹」**，选中包含 LOG 文件的目录。程序会立即递归扫描并显示发现的 LOG 文件数量。
2. 点击 **「开始转换」**，程序在后台逐个解析并生成 Excel。
3. 转换完成后，每个 `.log` 文件旁会生成一个同名 `.xlsx` 文件（例如 `A/B/test.log` → `A/B/test.xlsx`）。
4. 底部状态栏会汇总结果：成功数、失败数、共提取的数据条数。若有失败，通常是目标 Excel 被占用（例如正被 Excel 打开），关闭后重试即可。

## LOG 解析规则

解析器逐行读取日志，采用「主行 + 温度补充行」的配对模型（见 [`LogParser.cs`](src/Services/LogParser.cs)）：

**主行**（含时间戳与电池主要信息），示例：

```
09-25#14:03:22.123  ... bat info voltage = 4123, percent = 87, g_bat_info = 88, current = -450 ...
```

匹配并提取：

| 字段          | 来源关键字        | 说明        |
| ------------- | ----------------- | ----------- |
| 时间          | `MM-DD#HH:MM:SS`  | 记录时间戳  |
| 电压          | `voltage=`        | 单位由日志本身决定（如 mV） |
| 电量1         | `percent=`        | 电量百分比  |
| 电量2         | `g_bat_info=`     | 另一路电量  |
| 电流          | `current=`        | 可为负值（放电） |

**温度行**（紧随主行之后，可选），示例：

```
... temp_info = 32 ...
```

匹配后补充到上一条主行记录的**温度**字段。若主行之后没有出现温度行，则该记录温度留空。

> 匹配大小写不敏感，字段之间允许任意空白。数值支持负号。

## 输出格式

生成的 `.xlsx` 工作表名为 **「电池数据」**，首行为冻结的表头，列宽已预设：

| 时间 | 电压 | 电量1 | 电量2 | 电流 | 温度 |
| ---- | ---- | ----- | ----- | ---- | ---- |

- 表头为蓝底白字（`#2563EB`）、居中，字体为 Microsoft YaHei。
- 数值列以数字类型写入，便于在 Excel 中排序、绘图、计算。
- 首行冻结，滚动时表头始终可见。

## 从源码构建

### 环境要求

- [.NET 10 SDK](https://dotnet.microsoft.com/)（含 Windows Desktop / WPF 支持）
- Windows x64

### 调试构建 / 运行

```bash
cd src
dotnet build
dotnet run
```

### 发布单文件 EXE

```bash
cd src
dotnet publish -c Release -r win-x64 --self-contained true \
  -p:PublishSingleFile=true \
  -p:IncludeNativeLibrariesForSelfExtract=true \
  -p:EnableCompressionInSingleFile=true
```

产物位于 `src/bin/Release/net10.0-windows/win-x64/publish/BatteryLogix.exe`。

关键发布配置（见 [`BatteryLogix.csproj`](src/BatteryLogix.csproj)）：

- `SelfContained` / `PublishSingleFile`：自包含单文件。
- `PublishReadyToRun`：AOT 预编译，加快启动。
- `EnableCompressionInSingleFile`：压缩以减小体积。
- `SatelliteResourceLanguages=zh-Hans`：仅保留简体中文资源。

## 自动发布（CI）

仓库内置 GitHub Actions 工作流 [`.github/workflows/release.yml`](.github/workflows/release.yml)，**推送 tag 时自动构建并发布**：

```bash
git tag v1.0
git push origin v1.0
```

工作流会自动：

1. 定位唯一的 `.csproj`（多个时需在 `env.PROJECT` 指定）。
2. 以 `win-x64` 自包含单文件方式 `dotnet publish`。
3. 自动识别主 EXE，重命名为 `BatteryLogixV<版本号>.exe`（版本号取自 tag）。
4. 创建 GitHub Release，附带产物与自动生成的更新日志。

也可在 Actions 页面通过 `workflow_dispatch` 手动触发（仅上传构建产物，不创建 Release）。

## 项目结构

```
BatteryLogix/
├─ .github/workflows/
│  └─ release.yml            # 打 tag 自动构建并发布单文件 EXE
└─ src/
   ├─ App.xaml(.cs)          # 应用入口，启动时创建主窗口
   ├─ MainWindow.xaml(.cs)   # 主界面（深色卡片式 UI）
   ├─ BatteryLogix.csproj    # 项目与发布配置
   ├─ Models/
   │  └─ BatteryRecord.cs    # 单条电池记录（record）
   ├─ Services/
   │  ├─ LogParser.cs        # LOG 解析（源生成正则）
   │  ├─ ExcelWriter.cs      # 零依赖生成 .xlsx（OOXML）
   │  └─ LifetimeService.cs  # 运行有效期校验
   └─ ViewModels/
      ├─ MainViewModel.cs    # 主逻辑：扫描、转换、进度
      ├─ ObservableObject.cs # INotifyPropertyChanged 基类
      └─ RelayCommand.cs     # ICommand 实现
```

## 技术说明

- **框架**：WPF，`net10.0-windows`，启用 `Nullable` 与 `ImplicitUsings`。
- **架构**：MVVM。视图（XAML）通过数据绑定与 `MainViewModel` 通信，`RelayCommand` 承载命令，`ObservableObject` 提供属性变更通知。
- **解析**：使用 C# **源生成正则**（`[GeneratedRegex]`），编译期生成匹配代码，性能优于运行时正则；逐行流式读取，内存占用低。
- **Excel 生成**：**不依赖任何第三方库**（如 EPPlus、NPOI）。直接用 `ZipArchive` + `XmlWriter` 按 OOXML SpreadsheetML 规范手写 `.xlsx` 各部件（`workbook.xml`、`styles.xml`、`sheet1.xml` 等），先写 `.tmp` 再原子替换。
- **并发**：转换通过 `Task.Run` 放到后台线程，配合 `IProgress<T>` 回主线程更新 UI。
- **有效期**：`LifetimeService` 根据可执行文件的最后写入时间判断是否在有效期内（约 360 天），超期则启动时退出。

## 常见问题

**Q：部分文件转换失败怎么办？**
A：最常见原因是目标 `.xlsx` 正被 Excel 或其他程序占用/锁定。关闭占用后重新转换即可。也请确认文件夹有读写权限。

**Q：为什么有的记录没有温度？**
A：温度来自主行之后紧跟的 `temp_info` 行。若某条记录后没有对应的温度行，温度字段将留空。

**Q：运行后窗口一闪而过 / 打不开？**
A：程序内置了有效期校验；若使用的是较旧的构建产物，可能已超过有效期。请从 [Releases](https://github.com/larry-lvjin/BatteryLogix/releases) 下载最新版本。

**Q：需要安装 .NET 吗？**
A：下载 Release 中的单文件 EXE 无需安装；仅在从源码构建时需要 .NET 10 SDK。

## 许可证

本仓库暂未声明开源许可证。如需在你的场景下使用、分发或二次开发，请先与作者确认。
