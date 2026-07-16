# Serial Image to Editable PPT

把按页编号的 PPT 视觉稿（PNG/JPG）严格串行地重建为“复杂视觉保真、主要文字和简单结构可编辑”的 PowerPoint，并为每页保留预览、差异图、元素清单和质量报告。

这个仓库同时包含两个协作 skill：

- `serial-image-to-editable-ppt`：队列、断点恢复、逐页门禁、QA、审计、打包与有序合并。
- `codeximage-to-editable-ppt-v1`：单页视觉拆解与精修可编辑重建。

它们必须一起安装。

## 给新手的最快路径（Windows）

### 1. 安装 Codex 和基础工具

安装 Windows 版 ChatGPT 桌面应用并使用其中的 Codex。OpenAI 官方 Windows 文档提供 [Microsoft Store 下载和 winget 安装方式](https://learn.chatgpt.com/docs/windows/windows-app)。

再安装 Python 3.10 或更高版本。推荐同时安装 Git 和 Node.js LTS：

```powershell
winget install --id Git.Git
winget install --id OpenJS.NodeJS.LTS
winget install --id Python.Python.3.14
```

安装后关闭并重新打开 PowerShell，确认：

```powershell
python --version
git --version
node --version
```

### 2. 下载本仓库

会使用 Git：

```powershell
git clone https://github.com/luliu3945-arch/serial_image_to_ppt.git
cd serial_image_to_ppt
```

不会使用 Git：在 GitHub 页面点击 **Code → Download ZIP**，解压后在该目录打开 PowerShell。

### 3. 运行一键安装

```powershell
powershell -ExecutionPolicy Bypass -File .\setup.ps1
```

该脚本会：

1. 检查 Python 版本和两个 skill 的完整性。
2. 在 `%USERPROFILE%\.serial-image-to-ppt\venv` 创建隔离的 Python 环境。
3. 安装本项目的 Python 包。
4. 把两个 skill 安装到官方用户级目录 `%USERPROFILE%\.agents\skills`。
5. 写入运行时 Python 地址并执行环境自检。

如果 skill 没立即出现，彻底关闭并重新打开 Codex。Codex 官方文档说明，skills 可用于桌面应用、CLI 和 IDE；用户级 skill 当前放在 `$HOME/.agents/skills`。详见 [Build skills](https://learn.chatgpt.com/docs/build-skills)。

macOS/Linux 可在仓库目录执行：

```bash
bash ./setup.sh
```

这会安装跨平台部分；PowerPoint COM 真实预览和仓库自带的合并脚本仍只支持 Windows。

### 4. 检查生产环境

最完整的 Windows 工作流建议安装 Microsoft PowerPoint 桌面版，然后执行：

```powershell
& "$HOME\.serial-image-to-ppt\venv\Scripts\python.exe" .\scripts\doctor.py --require-powerpoint
```

如果要自动识别中英文文字，再安装 Tesseract OCR 及 `eng`、`chi_sim` 语言包，然后执行：

```powershell
& "$HOME\.serial-image-to-ppt\venv\Scripts\python.exe" .\scripts\doctor.py --require-ocr
```

出现 `Environment: PASS` 后即可开始。

### 5. 准备图片并给 Codex 提示词

把图片放入同一目录：

```text
page_1.png
page_2.png
page_3.png
...
page_N.png
```

在 Codex 中发送：

```text
请使用 $serial-image-to-editable-ppt。

图片目录：D:\PPT截图\物流培训课件
输出目录：D:\PPT成品\物流培训课件_可编辑版

识别 page_1.png 到最大页码，按照 1 到 N 严格串行处理。
每次只制作一页；当前页完成视觉、结构、元素清单和边界检查后，才能处理下一页。
采用“复杂视觉保真 + 主要文字可编辑”的混合重建方式，不要把整页截图铺进 PPT。
全部完成后运行总审计、生成交付包，需要时合并为一个 PPTX，并清理临时 hook。
```

skill 在创建队列前必须收到用户提供的图片绝对路径，不会通过全盘搜索猜测输入目录。

## 环境与依赖总览

| 组件 | 级别 | 用途 |
|---|---|---|
| ChatGPT/Codex 桌面应用、Codex CLI 或 IDE | 必需 | 识别并执行两个 skill，完成视觉判断和重建编排 |
| Codex presentations/artifact-tool 运行时 | 必需 | 创建可编辑 PPTX、导出预览和检查对象；由 Codex 提供，不要安装同名的第三方 npm 包 |
| Python 3.10+ | 必需 | 队列、拆解、QA、审计和打包脚本 |
| Python 包（见 `requirements.txt`） | 必需 | OpenCV、NumPy、Pillow、python-pptx、pytesseract、PyYAML |
| Microsoft PowerPoint 桌面版（Windows） | 完整生产流程强烈建议 | 真实 PowerPoint 验收预览；`merge_delivery.ps1` 的有序合并依赖 COM |
| PowerShell 5.1+ | Windows 合并需要 | 运行安装脚本和 PowerPoint COM 合并脚本 |
| Tesseract + `eng`/`chi_sim` | 启用自动 OCR 时需要 | 自动识别英文/简体中文；缺失时仍可由 Codex人工重建可读文字 |
| LibreOffice | 输入为 PPT/PPTX/PDF 时需要 | 将文档转为 PDF 供基线拆解；纯 PNG/JPG 输入不需要 |
| Poppler `pdftoppm` | 输入为 PPT/PPTX/PDF 时需要 | 将 PDF 渲染成逐页 PNG；纯 PNG/JPG 输入不需要 |
| Node.js LTS | 推荐 | Codex 与演示文稿工具常用运行时；桌面版可能自带运行时 |
| Git | 推荐 | 克隆、更新仓库和版本管理；下载 ZIP 时可不装 |
| 中文字体 | 推荐 | 减少 PowerPoint 字体替换和换行漂移，例如微软雅黑、思源黑体 |

项目本身不直接调用 OpenAI API，因此不要求单独配置 `OPENAI_API_KEY`；但使用 Codex 需要可用的 ChatGPT/Codex 账号和相应产品权限。

完整说明见 [环境、工具与依赖](docs/ENVIRONMENT.md)。

## Python 依赖

一键安装会读取根目录 `requirements.txt`：

```text
opencv-python>=4.8,<5
numpy>=1.24,<3
pillow>=10,<13
python-pptx>=1.0,<2
pytesseract>=0.3.10,<1
PyYAML>=6,<7
```

如果不用一键安装，也可以手动创建环境：

```powershell
python -m venv "$HOME\.serial-image-to-ppt\venv"
& "$HOME\.serial-image-to-ppt\venv\Scripts\python.exe" -m pip install -r requirements.txt
```

注意：`pytesseract` 只是 Python 调用层，不包含 Tesseract OCR 主程序和语言包。

## 环境自检工具

基础检查：

```powershell
& "$HOME\.serial-image-to-ppt\venv\Scripts\python.exe" .\scripts\doctor.py --strict
```

按工作场景加严：

```powershell
# 要求真实 PowerPoint 预览与 COM 合并
& "$HOME\.serial-image-to-ppt\venv\Scripts\python.exe" .\scripts\doctor.py --strict --require-powerpoint

# 要求自动中英文 OCR
& "$HOME\.serial-image-to-ppt\venv\Scripts\python.exe" .\scripts\doctor.py --strict --require-ocr

# 要求能够处理 PPT/PPTX/PDF 输入
& "$HOME\.serial-image-to-ppt\venv\Scripts\python.exe" .\scripts\doctor.py --strict --require-document-input

# 机器可读报告
& "$HOME\.serial-image-to-ppt\venv\Scripts\python.exe" .\scripts\doctor.py --json
```

## 仓库结构

```text
.
├─ README.md
├─ AGENTS.md
├─ setup.ps1
├─ setup.sh
├─ requirements.txt
├─ docs/
│  └─ ENVIRONMENT.md
├─ scripts/
│  └─ doctor.py
├─ serial-image-to-editable-ppt/
│  ├─ SKILL.md
│  ├─ agents/
│  ├─ references/
│  └─ scripts/
└─ codeximage-to-editable-ppt-v1/
   ├─ SKILL.md
   ├─ agents/
   ├─ references/
   └─ scripts/
```

`AGENTS.md` 是给打开本仓库的 Codex 使用的项目级操作约束；`SKILL.md` 是可复用工作流本身。

## 工作流程

```text
用户提供图片绝对地址
        ↓
检查图片与页码范围
        ↓
创建临时串行 hook
        ↓
当前页：基线拆解 → 精修可编辑重建
        ↓
artifact 预览 → PowerPoint 真实预览（可用时）
        ↓
视觉、结构、清单、边界 QA
        ↓
通过后才领取下一页
        ↓
全量审计、交付 ZIP、按需合并
        ↓
清理临时 hook
```

默认顺序为第 `1` 页到第 `N` 页，或用户指定范围内的升序。

## 中断后恢复

回到原 Codex 任务发送：

```text
continue
```

或：

```text
继续使用 $serial-image-to-editable-ppt，从现有队列状态恢复。先检查 status，不要重做已经验收通过的页面。
```

任务未完成时不要删除输出目录中的：

```text
.serial_image_to_ppt_state.json
```

## 输出

每页输出：

```text
page_1_refined_editable.pptx
page_1_refined_editable_output/
```

质量证据目录包含预览、原图、重建图、差异图、元素清单、裁剪素材和 QA 报告。全部通过后会生成总审计 JSON/CSV 和交付 ZIP。

## 合并为一个 PPTX

必须先通过全部单页审计。Windows + Microsoft PowerPoint 环境下：

```powershell
powershell -ExecutionPolicy Bypass -File `
  serial-image-to-editable-ppt/scripts/merge_delivery.ps1 `
  -OutputDir "D:\PPT成品\物流培训课件_可编辑版" `
  -Start 1 -End 48 `
  -OutputFile "D:\PPT成品\物流培训课件_48页合并版.pptx" `
  -PreviewDir "D:\PPT成品\物流培训课件_合并预览"
```

脚本使用 PowerPoint 原生 `InsertFromFile`，保留可编辑对象并核对页数。

## 更新

Git 安装方式：

```powershell
git pull
powershell -ExecutionPolicy Bypass -File .\setup.ps1
```

更新后重新运行 `setup.ps1`，确保 skill 文件和 Python 依赖同步。

## 常见问题

### Codex 找不到 skill

确认下面两个文件存在：

```text
%USERPROFILE%\.agents\skills\serial-image-to-editable-ppt\SKILL.md
%USERPROFILE%\.agents\skills\codeximage-to-editable-ppt-v1\SKILL.md
```

然后重启 Codex，并在提示词里显式写 `$serial-image-to-editable-ppt`。

### PowerShell 不允许运行脚本

只对本次命令绕过：

```powershell
powershell -ExecutionPolicy Bypass -File .\setup.ps1
```

不要在不了解影响时全局关闭执行策略。

### 安装了 pytesseract 仍无法 OCR

还需要安装 Tesseract 主程序，并确保 `tesseract` 在 PATH 中；中文还需要 `chi_sim` 语言数据。用 `doctor.py --require-ocr` 检查。

### 没有 PowerPoint 能否使用

可以执行图片拆解和 artifact 预览，但不能运行基于 PowerPoint COM 的真实验收预览和 `merge_delivery.ps1`。完整生产交付建议在装有 PowerPoint 桌面版的 Windows 机器完成最终验收与合并。

### 是否必须安装 LibreOffice 和 Poppler

输入已经是 PNG/JPG 时不需要。只有直接处理 PPT、PPTX 或 PDF 输入时才需要。

### 为什么不能并行做多页

串行门禁用于阻止字体、裁边、缺图和版式错误连续污染后续页面。当前页未通过时，下一页不得开始。

## 许可证与第三方软件

仓库目前未包含许可证文件。公开分发前，仓库维护者应选择并添加合适的开源许可证。Microsoft PowerPoint、Tesseract、LibreOffice、Poppler、Python 包和 Codex 各自适用其自身许可与使用条款。
