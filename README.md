# Serial Image to Editable PPT

面向 Codex 的“image图批量重建为可编辑 PowerPoint”工作流。

它会严格按照页码顺序，一次只处理一张截图；当前页完成 artifact 预览、真实 PowerPoint 预览、结构、元素清单和页面边界检查后，才继续下一页。任务中断后可以从已保存的队列状态恢复，全部完成后会生成审计结果和交付包，可按顺序合并为一个 PPTX，并清理临时 hook。

## 适合什么场景

- 把一组 图片 重新制作成可编辑 PPTX
- 将标题、正文、表格、箭头、边框和色块恢复为可编辑对象
- 把照片、设备图、软件截图和复杂图标裁成独立图片
- 按页串行处理，避免同一种错误连续污染几十页
- 中断后继续，不重做已经验收通过的页面
- 将生成的页面按顺序追加到已有 PPTX
- 使用 PowerPoint 原生插入将全部已验收单页合并为一个有序、可编辑的 PPTX

## 工作方式

```text
用户提供图片地址
        ↓
检查图片和页码范围
        ↓
创建临时串行 hook
        ↓
读取当前页截图
        ↓
基础拆解 → 精修可编辑重建
        ↓
artifact 预览 → 真实 PowerPoint 预览
        ↓
视觉、结构、清单、边界检查
        ↓
通过后领取下一页
        ↓
全量审计与打包
        ↓
按需合并为一个 PPTX
        ↓
清理临时 hook
```

默认处理顺序是第 `1` 页到第 `N` 页，或用户指定范围内的升序。

## 仓库内容

```text
.
├─ AGENTS.md
├─ README.md
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

本仓库包含两个 skill：

- `serial-image-to-editable-ppt`：负责串行队列、恢复、QA、审计、打包、合并和清理。
- `codeximage-to-editable-ppt-v1`：负责单页截图的精修可编辑重建。

两个文件夹需要一起安装。

如果要让全新的智能体直接接手仓库，请先让它读取根目录 [AGENTS.md](AGENTS.md)。该文件包含强制串行门禁、真实 PowerPoint 验收、失败恢复、最终审计、合并和 hook 清理规则。

## 安装

### 方法一：下载 ZIP

1. 在 GitHub 仓库页面点击 **Code → Download ZIP**。
2. 解压下载的文件。
3. 把下面两个文件夹复制到 Codex skills 目录：

```text
serial-image-to-editable-ppt
codeximage-to-editable-ppt-v1
```

Windows 默认位置：

```text
C:\Users\你的Windows用户名\.codex\skills\
```

macOS 或 Linux 默认位置：

```text
~/.codex/skills/
```

4. 完全关闭并重新打开 Codex。

安装后的目录应类似：

```text
.codex/skills/
├─ serial-image-to-editable-ppt/
│  └─ SKILL.md
└─ codeximage-to-editable-ppt-v1/
   └─ SKILL.md
```

不要在外面额外套一层同名目录，否则 Codex 可能找不到 `SKILL.md`。

### 方法二：使用 Git

```bash
git clone <你的GitHub仓库地址>
```

克隆后，仍需把仓库中的两个 skill 文件夹复制到 `.codex/skills/`。

## 准备截图

建议把所有图片放在同一个文件夹中，并按页码命名：

```text
page_1.png
page_2.png
page_3.png
...
page_36.png
```

要求：

- 页码清晰且不重复。
- 建议所有图片使用相同尺寸。
- 图片越清晰，重建效果通常越好。
- 推荐使用 `page_<页码>.png` 命名。

## 快速开始

先复制图片文件夹的绝对路径，例如：

```text
D:\PPT截图\物流培训课件
```

然后在 Codex 中发送：

```text
请使用 serial-image-to-editable-ppt skill。

图片目录：D:\PPT截图\物流培训课件
输出目录：D:\PPT成品\物流培训课件_可编辑版

请识别 page_1.png 到最大页码，按照 1 到 N 的顺序严格串行处理。
每次只制作一页；当前页完成视觉、结构、清单和边界检查后，才能处理下一页。
每页必须先做 artifact 预览，再做真实 PowerPoint 预览并运行 QA。
全部完成后运行总审计、生成交付包，按顺序合并为一个 PPTX，并清理临时 hook。
```

skill 必须先收到用户提供的真实图片地址，才会创建 hook。它不会通过全盘搜索猜测输入目录。

## 只处理部分页面

例如只处理第 5–12 页：

```text
请使用 serial-image-to-editable-ppt skill。

图片目录：D:\PPT截图\物流培训课件
输出目录：D:\PPT成品\物流培训课件_可编辑版
只处理第 5 页到第 12 页，按照页码升序严格串行执行。
```

## 合并为一个 PPTX

全部单页通过总审计后，可以使用仓库脚本按页码升序合并：

```powershell
powershell -ExecutionPolicy Bypass -File `
  serial-image-to-editable-ppt/scripts/merge_delivery.ps1 `
  -OutputDir "D:\PPT成品\物流培训课件_可编辑版" `
  -Start 1 -End 48 `
  -OutputFile "D:\PPT成品\物流培训课件_48页合并版.pptx" `
  -PreviewDir "D:\PPT成品\物流培训课件_合并预览"
```

脚本使用 PowerPoint 原生插入方式保留可编辑对象，并验证合并页数。建议检查导出的全部预览，确认第 `N` 页仍对应 `page_N_refined_editable.pptx`。

## 追加到已有 PPTX

单页全部制作完成后，可以继续发送：

```text
请把本项目生成的第 13–36 页按照页码顺序，追加到下面这个 PPTX 的末尾：

D:\PPT成品\原课件_追加前备份.pptx

追加前检查目标文件原有页数；追加后重新打开并渲染全部页面，确认原页面不变、总页数正确、没有缺图和裁边。
```

建议提前复制一份原 PPTX 作为备份。

## 中断后恢复

在原 Codex 任务中发送：

```text
continue
```

或者：

```text
继续使用 serial-image-to-editable-ppt，从现有队列状态恢复。先检查 status，不要重做已经验收通过的页面。
```

队列状态保存在输出目录：

```text
.serial_image_to_ppt_state.json
```

任务未完成时不要手动删除该文件。全部页面通过审计后，skill 会清理临时状态。

## 输出文件

每页会生成一个可编辑 PPTX：

```text
page_1_refined_editable.pptx
page_2_refined_editable.pptx
...
page_N_refined_editable.pptx
```

每页还会生成质量证据目录：

```text
page_1_refined_editable_output/
```

其中可能包含：

- 完整页面预览
- 原图与重建图的对比图
- 差异图
- 元素清单
- QA 检查结果
- 独立裁剪图片

全部完成后会生成总审计 JSON、CSV 和交付 ZIP。

如果用户要求一个完整课件，还会生成按页码合并的可编辑 PPTX。单页文件和 QA 证据仍会保留，便于定位和返修。

## 为什么不是所有内容都变成文字和形状

这个工作流追求“可编辑性”和“视觉还原度”的平衡：

- 标题、正文、标签、简单表格和基础图形会尽量重建为原生 PowerPoint 对象。
- 照片、设备图、复杂软件界面和复杂图标通常保留为独立 PNG。
- 不允许简单地把整张截图铺满页面并宣称已经可编辑。

## 常见问题

### Codex 要求先提供图片地址

这是正常的输入安全规则。请提供图片文件夹或某张编号图片的绝对路径。

### 无法识别页码

检查文件名是否包含清晰数字。推荐使用 `page_1.png`、`page_2.png`，避免“最终版”“副本”等模糊名称。

### 为什么不能同时生成很多页

本 skill 的目标是稳定质量。当前页验收通过后才处理下一页，可以及时拦截字体、错位、缺图和裁边问题。

### 可以倒序处理吗

通用默认规则是从第 1 页到第 N 页升序处理。特殊项目确实需要其他顺序时，应在请求中明确说明。

### 处理速度为什么比较慢

每页都会经历拆解、重建、渲染和检查。复杂图表、软件界面和设备图片较多时，需要更长时间。

### 总审计为什么提示缺少 baseline manifest

基础拆解的清单位于嵌套批次目录。每页 baseline 完成后必须运行 `stage_baseline_evidence.py`，把 JSON 和 CSV 清单归档到 refined bundle。审计失败时不要清理 hook，应先按照审计 JSON 的 `missing_artifacts` 补齐文件并重跑全量审计。

## 最简使用口诀

```text
安装两个 skill → 重启 Codex → 准备 page_1 到 page_N
→ 提供绝对路径 → 粘贴提示词 → 等待逐页验收
→ 获取单页 PPTX、合并 PPTX、质量证据和交付 ZIP
```

