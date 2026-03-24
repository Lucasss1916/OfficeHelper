# OfficeHelper

一个基于 WPF 和 .NET 8 的 Excel 分数分配桌面工具。

它可以读取 Excel 中的学生总分和各学科满分信息，按设定规则自动生成每科分数，支持多工作表查看、应用内预览编辑、按百分比限制单科最高分，并可导出新的 Excel 结果文件。

## 功能特性

- 支持读取单个 Excel 文件中的多个 Sheet，并在应用内切换查看
- 自动识别表头中的姓名、学科满分、总分列
- 按总分自动分配各学科分数
- 支持设置单科最低分
- 支持设置单科最高分百分比上限
- 支持固定随机种子，保证同一批数据生成结果稳定
- 生成预览后可直接在应用内编辑某一行数据
- 编辑总分或单科分数后，会按当前规则自动联动修正该行其他数值
- 支持将结果导出为新的 Excel 文件
- 支持打包为单文件 exe
- 支持 GitHub Actions 自动编译

## 运行环境

- Windows
- .NET 8 SDK

## Excel 格式要求

Excel 第一行应包含以下信息：

- 姓名列，例如：`姓名`
- 学科列，列名中要包含该学科满分，例如：`语文(120分)`、`数学（150分）`
- 总分列，例如：`总分`

示例表头：

```text
姓名 | 语文(120分) | 数学(150分) | 英语(120分) | 总分
```

## 使用方式

1. 启动程序。
2. 选择 Excel 文件。
3. 在下拉框中选择需要处理的 Sheet。
4. 设置单科最低分。
5. 如有需要，设置“最高分百分比上限”。
6. 点击“生成预览数据”。
7. 在预览表中检查结果，必要时可以直接修改某一行的数据。
8. 点击“导出结果到 Excel”保存结果。

## 预览编辑规则

- 修改 `原总分`：程序会重新计算该行各学科分数。
- 修改某一学科分数：程序会尽量保留你修改的该科分数，并按规则重新分配该行其他学科。
- `计算和` 为只读列，用于显示当前行实际总和。
- 所有分配都会同时遵守以下规则：
  - 不低于单科最低分
  - 不高于该学科满分
  - 不高于该学科的最高分百分比上限

## 本地开发

克隆项目后，在项目目录运行：

```powershell
dotnet build .\RandomScoreAllocatorWPF.csproj
```

运行程序：

```powershell
dotnet run --project .\RandomScoreAllocatorWPF.csproj
```

## 本地打包单文件 EXE

项目已经内置单文件发布配置，执行下面命令即可：

```powershell
dotnet publish .\RandomScoreAllocatorWPF.csproj /p:PublishProfile=Properties\PublishProfiles\SingleFile-win-x64.pubxml
```

打包完成后，单文件 exe 位于：

```text
bin\Release\net8.0-windows\publish\win-x64-single-file\RandomScoreAllocatorWPF.exe
```

## GitHub Actions 自动编译

仓库已包含 GitHub Actions 工作流：

- 推送到 `main` 或 `master` 时自动编译
- 创建 Pull Request 时自动编译
- 推送 `v*` 标签时自动创建 GitHub Release，并上传单文件 exe

示例：

```powershell
git tag v1.0.0
git push origin v1.0.0
```

推送标签后，可以在 GitHub 的 Releases 页面下载自动生成的 exe。

## 项目结构

```text
OfficeHelper
├─ App.xaml
├─ MainWindow.xaml
├─ MainWindow.xaml.cs
├─ ScoreAllocator.cs
├─ RandomScoreAllocatorWPF.csproj
├─ icon.ico
├─ Properties
│  └─ PublishProfiles
└─ .github
   └─ workflows
```

## 说明

- 当前项目是轻量级单体 WPF 应用，主界面逻辑集中在 `MainWindow.xaml.cs`
- 核心分数分配算法位于 `ScoreAllocator.cs`
- 单文件发布为 `win-x64` 自包含模式，因此 exe 文件体积会相对较大，这是正常现象
