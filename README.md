# Word 格式助手

一个 Word VSTO 加载项，用于一键应用中文报告格式（公文报告 / 审计报告 / 自定义格式）。

## 功能

- **预设方案**：公文报告（GB/T 9704-2012）、审计报告、自定义格式
- **格式设置**：可视化设置标题、正文、表格格式
- **段落级别**：一键设置大纲级别（一级/二级/三级/四级标题、正文）
- **标点转换**：保持原样 / 全部转全角 / 全部转半角 / 智能转换

## 项目结构

- `src/WordFormatHelper.AddIn/` — VSTO 加载项主项目
- `src/WordFormatHelper.Installer/` — 安装引导 exe
- `build/build.ps1` — 一键构建脚本

## 构建

```powershell
powershell -ExecutionPolicy Bypass -File build\build.ps1
```

构建产物：`dist/WordFormatHelperSetup.exe`，以管理员身份运行即可安装。
