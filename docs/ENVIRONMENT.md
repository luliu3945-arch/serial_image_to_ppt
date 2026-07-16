# 环境、工具与依赖

本文面向一台没有配置过本项目的新机器，说明每个组件为什么存在、什么时候必须安装，以及如何验证。

## 1. 支持范围

### 推荐生产环境

- Windows 10/11 64 位
- ChatGPT Windows 桌面应用中的 Codex，或 Codex CLI/IDE
- Python 3.10+
- PowerShell 5.1+
- Microsoft PowerPoint 桌面版
- 8 GB 以上内存和足够存放逐页预览、裁图与 PPTX 的磁盘空间

Windows + PowerPoint 是完整验收与有序合并的参考环境。PowerPoint COM 能发现字体替换、换行变化、连接线错误和图片裁切等 artifact 预览未必暴露的问题。

### macOS/Linux

两个 skill、Python 队列、基线拆解和普通 QA 可以运行；但仓库中的 `merge_delivery.ps1` 和 PowerPoint COM 真实预览只支持 Windows。跨平台环境应使用 Codex presentations/artifact-tool 预览，并把最终 PowerPoint 验收与合并转移到 Windows + PowerPoint 机器。

跨平台安装命令：

```bash
bash ./setup.sh
```

## 2. Codex 侧依赖

### 必需能力

- 能读取本地图片和文件。
- 能运行 shell/Python 脚本。
- 能使用 presentations/artifact-tool 创建和渲染 PPTX。
- 能查看原图与渲染预览并做视觉对照。

artifact-tool 属于 Codex 的演示文稿运行时。本仓库不通过 npm 声明它，也不应安装来源不明的同名包。

### Skill 安装位置

根据当前 OpenAI Codex 文档，用户级 skills 放在：

```text
$HOME/.agents/skills
```

Windows 对应：

```text
C:\Users\<用户名>\.agents\skills
```

仓库级 skills 可放在仓库根目录的 `.agents/skills`。本项目的 `setup.ps1` 使用用户级位置，使两个 skill 对所有本地项目可用。

旧环境可能仍使用 `%USERPROFILE%\.codex\skills`。代码保留了该位置的兼容搜索，但新安装应使用 `.agents\skills`。

## 3. Python

### 版本

要求 Python 3.10 或更高版本。脚本使用了现代类型注解、`pathlib`、`dataclasses` 和标准库 ZIP/JSON/CSV 能力。

### 包与职责

| 包 | 职责 |
|---|---|
| `opencv-python` | 图像分割、轮廓与视觉区域分析 |
| `numpy` | 像素矩阵和差异计算 |
| `pillow` | 裁图、透明度、叠图、预览与差异图 |
| `python-pptx` | 基线 PPTX、对象计数、结构 QA 与审计 |
| `pytesseract` | 调用外部 Tesseract OCR |
| `PyYAML` | 读取 YAML 配置 |

`setup.ps1` 把依赖安装到：

```text
%USERPROFILE%\.serial-image-to-ppt\venv
```

并在已安装的串行 skill 中写入 `runtime.json`，让 Codex 知道应优先使用哪个 Python。

## 4. 系统软件

### Microsoft PowerPoint

用途：

- 通过 COM 打开单页 PPTX 并导出真实渲染预览。
- 通过 `Slides.InsertFromFile` 按顺序合并已验收的单页文件。
- 验证对象在实际 PowerPoint 中仍可编辑且没有字体/排版漂移。

如果缺失：不影响基础图像拆解，但不能完成本仓库定义的最强 Windows 验收与 COM 合并。

检查：

```powershell
& "$HOME\.serial-image-to-ppt\venv\Scripts\python.exe" scripts\doctor.py --require-powerpoint
```

### Tesseract OCR

用途：自动提取图片文字，尤其是基线拆解阶段。

注意：

- `pip install pytesseract` 不会安装 Tesseract 主程序。
- 英文需要 `eng`。
- 简体中文需要 `chi_sim`。
- Tesseract 安装目录必须在 PATH，或由运行环境显式配置。

检查：

```powershell
tesseract --version
tesseract --list-langs
& "$HOME\.serial-image-to-ppt\venv\Scripts\python.exe" scripts\doctor.py --require-ocr
```

缺失 OCR 时，skill 仍可把图片作为视觉真值，由 Codex重建清晰可读的主要文字；自动 OCR 能力会下降。

### LibreOffice

用途：直接输入 PPT/PPTX/PDF 时，把文档转为 PDF。对于已经准备好的 PNG/JPG 页面不需要。

脚本会查找 `libreoffice` 命令。Windows 安装后如果命令不在 PATH，应把 LibreOffice 的 `program` 目录加入 PATH，或在终端中提供可调用的命令包装。

### Poppler / pdftoppm

用途：把 PDF 渲染为逐页 PNG。对于已经准备好的 PNG/JPG 页面不需要。

检查：

```powershell
pdftoppm -v
& "$HOME\.serial-image-to-ppt\venv\Scripts\python.exe" scripts\doctor.py --require-document-input
```

### Git

用于克隆、更新、提交和恢复项目。下载 ZIP 使用时不是运行必需项。

### Node.js

Codex 和演示文稿工具常用的运行时。项目 Python 脚本不直接调用 Node，但完整 Codex 环境建议安装 LTS 版。

## 5. 字体

目标机器缺少源字体时，PowerPoint 会替换字体，可能造成：

- 标题宽度变化；
- 中文换行变化；
- 文本溢出或缩小；
- 符号替换。

建议安装源稿实际字体。中文通用回退字体包括微软雅黑、黑体和思源黑体。字体是否允许随项目分发取决于字体许可证，本仓库不包含字体文件。

## 6. 网络与账号

- 第一次下载仓库和安装 Python 包需要互联网。
- 使用 Codex 需要可用的 ChatGPT/Codex 账号及产品权限。
- 本仓库脚本不直接调用 OpenAI API，不要求单独设置 `OPENAI_API_KEY`。
- 图片、PPTX 和证据文件默认保存在用户指定的本地目录。

## 7. 权限

Codex 必须能读取输入图片并写入输出目录。输出放在桌面、其他磁盘或受保护目录时，可能需要用户批准文件写入。

Microsoft PowerPoint COM 首次运行时可能弹出安全或文件访问提示。不要通过关闭全部系统安全策略来绕过问题；只授予本项目所需的目录和命令权限。

## 8. 一键安装行为

运行：

```powershell
powershell -ExecutionPolicy Bypass -File .\setup.ps1
```

默认写入：

```text
Skills:  %USERPROFILE%\.agents\skills
Runtime: %USERPROFILE%\.serial-image-to-ppt\venv
```

自定义：

```powershell
powershell -ExecutionPolicy Bypass -File .\setup.ps1 `
  -SkillRoot "D:\CodexData\skills" `
  -RuntimeRoot "D:\CodexData\serial-image-to-ppt-runtime" `
  -PythonCommand "py"
```

如果自定义 SkillRoot 不属于 Codex 会扫描的位置，还需要通过 Codex 配置或符号链接暴露该目录；新手建议保留默认值。

## 9. 验证层级

### 最小图片输入环境

```powershell
python scripts\doctor.py --strict
```

要求：两个 skill、Python 3.10+ 和全部 Python 包。

### 中英文 OCR 环境

```powershell
python scripts\doctor.py --strict --require-ocr
```

额外要求：Tesseract 的 `eng` 和 `chi_sim`。

### PPT/PDF 输入环境

```powershell
python scripts\doctor.py --strict --require-document-input
```

额外要求：LibreOffice 与 `pdftoppm`。

### 完整 Windows 生产环境

```powershell
python scripts\doctor.py --strict --require-powerpoint --require-ocr
```

如果输入还包含 PPT/PDF，再加 `--require-document-input`。

## 10. 常见错误定位

| 错误 | 原因 | 处理 |
|---|---|---|
| `Missing companion skill` | 两个 skill 未放在同一 skills 根目录 | 重跑 `setup.ps1` |
| `No module named pptx/cv2/...` | 使用了未安装依赖的 Python | 使用 `runtime.json` 中的 Python或重跑安装 |
| `tesseract is not installed` | 只装了 pytesseract | 安装 Tesseract 主程序并加入 PATH |
| `chi_sim.traineddata not found` | 缺少简体中文语言包 | 安装 `chi_sim` 并检查 `tesseract --list-langs` |
| `libreoffice did not produce a PDF` | LibreOffice 缺失或不在 PATH | 安装并重开终端；纯图片输入可跳过 |
| `pdftoppm did not produce page PNG files` | Poppler 缺失 | 安装 Poppler 并将 `pdftoppm` 加入 PATH |
| `PowerPoint.Application` 创建失败 | 未安装桌面版 PowerPoint 或 COM 注册异常 | 修复 Microsoft 365/Office 安装 |
| skill 不显示 | 安装路径错误或 Codex 尚未刷新 | 检查 `.agents/skills/<name>/SKILL.md` 并重启 Codex |
