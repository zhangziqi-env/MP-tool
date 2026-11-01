# MP Label-V1

[English](#english) · [中文](#中文)

---

## English

MP Label-V1 is a Windows desktop application for annotating microplastic particles in microscopy images. It combines marking, classification, depth management, ID management and export in one local tool, so a batch of images can be organised by depth or sampling position.

The program runs fully offline; images and annotation data stay on the local machine. English and Chinese builds are provided.

### Files

| File | Description |
| --- | --- |
| `MP Label-V1.exe` | English single-file application |
| `MP Label-V1_中文版.exe` | Chinese single-file application |
| `MP Label-V1.py` | English Python source |
| `MP Label-V1_中文版.py` | Chinese Python source |
| `User-Guide-EN.md` | English user guide |
| `使用说明.md` | Chinese user guide |
| `requirements.txt` | Dependencies for the Python source |
| `CITATION.cff` | Citation metadata |
| `LICENSE` | GNU GPL v3 license text |

### Quick start

#### Using the exe (recommended)

1. Put `MP Label-V1.exe` or `MP Label-V1_中文版.exe` in a **writable folder** (Desktop or Downloads; not `C:\Program Files`).
2. Double-click to run.
3. If Windows shows “Unknown publisher”, choose “More info → Run anyway”.

The exe already contains Tcl/Tk, Pillow, openpyxl and the python-docx template; Python is not required. It extracts `_MEI...` next to the exe on startup and removes it on a normal close. If the process is killed, delete the leftover folder manually.

#### Using the Python source

```bat
python -m pip install -r requirements.txt
python "MP Label-V1.py"
:: Chinese version
python "MP Label-V1_中文版.py"
```

Requires Windows 10/11, Python 3.10+, with Tcl/Tk (the python.org installer includes it by default).

### Features

- Single-particle annotation: type, judgment (MP / NoMP / Test), Test subtype, colour and remark.
- Colour record: transparent, white, black, blue, red, yellow, green, brown, grey, other.
- Rectangle and polygon lasso batch annotation; saved regions can be exported to Word by judgment.
- Depth grouping and ID management: Renumber, Continue Num, Delete ID, Edit Remark, Edit Colour.
- Undo/redo, JSON project save/open, autosave recovery.
- Excel export: 7-column detail plus depth × type statistics.
- Word export: full-resolution image report grouped by depth; lasso regions by MP / NoMP / Test.
- Fully local; no data is uploaded.

### Design goals

- Keep a complete photographic record of every sample for long-term storage and review.
- Annotate and export offline after image capture, reducing microscope time.
- Use the remark field mainly for the microplastic category/species, so exported tables are easy to filter.
- Triage unknown particles quickly; when the same type occurs many times, choose a representative for confirmatory testing instead of testing every occurrence.

### Documentation

- [English user guide](User-Guide-EN.md)
- [Chinese user guide（中文使用说明）](使用说明.md)

### Notes

- The exe must be in a writable folder. If local extraction is blocked, use the Python source.
- Word reports are large by design: images keep their original pixel resolution and are not thumbnailed.
- Projects store absolute image paths; images need to be re-added after moving to another computer or folder.
- The program is not code-signed, so SmartScreen may warn.

### License

GNU General Public License v3.0（GPL-3.0-only）. See [LICENSE](LICENSE).

Copyright (C) 2026 MP Label-V1 contributors.

If you redistribute a modified version or an exe, the GPL requires you to make the corresponding source code available.

### Citation

If you use MP Label-V1 in your work, citation metadata is provided in [CITATION.cff](CITATION.cff).

---

## 中文

MP Label-V1 是一个面向**微塑料显微图像**的 Windows 桌面标注工具。它把拍照后的标注、分类、深度管理、编号和导出集中在一个本地程序里完成，适合按深度或站位处理一批显微照片。

程序完全离线运行，图片和标注数据只保存在本机。提供英文版和中文版。

### 文件清单

| 文件 | 说明 |
| --- | --- |
| `MP Label-V1.exe` | 英文版单文件程序 |
| `MP Label-V1_中文版.exe` | 中文版单文件程序 |
| `MP Label-V1.py` | 英文版 Python 源码 |
| `MP Label-V1_中文版.py` | 中文版 Python 源码 |
| `User-Guide-EN.md` | 英文使用说明 |
| `使用说明.md` | 中文使用说明 |
| `requirements.txt` | 运行 Python 源码所需依赖 |
| `CITATION.cff` | 引用信息 |
| `LICENSE` | GNU GPL v3 许可证全文 |

### 快速开始

#### 使用 exe（推荐）

1. 把 `MP Label-V1.exe` 或 `MP Label-V1_中文版.exe` 放到**可写的文件夹**（桌面、下载目录即可；不要放 `C:\Program Files`）；
2. 双击运行；
3. Windows 提示“未知发布者”时选择“更多信息 → 仍要运行”。

exe 已包含 Tcl/Tk、Pillow、openpyxl 和 python-docx 模板，不需要安装 Python。启动时会在 exe 同目录临时解压 `_MEI...`，正常关闭会自动清理；强制结束可能留下该目录，手动删除即可。

#### 使用 Python 源码

```bat
python -m pip install -r requirements.txt
python "MP Label-V1.py"
:: 中文版
python "MP Label-V1_中文版.py"
```

需要 Windows 10/11、Python 3.10+，且 Python 包含 Tcl/Tk。

### 功能

- 单张标注：类型、判断（MP / NoMP / Test）、Test 子类型、颜色和备注。
- 颜色记录：透明、白色、黑色、蓝色、红色、黄色、绿色、棕色、灰色、其他。
- 矩形/多边形套索批量标注；保存的套索区域可按判断导出到 Word。
- 按深度分组编号：Renumber、Continue Num、Delete ID、Edit Remark、Edit Colour。
- 撤销/重做、JSON 工程保存/打开、自动保存恢复。
- Excel 导出：7 列明细 + 深度 × 类型统计。
- Word 导出：按深度分组的全分辨率整图报告；按 MP / NoMP / Test 导出套索区域。
- 纯本地运行，数据不上传。

### 设计初衷

- 给所有样品留下完整的照片档案，便于长期保存和复查。
- 拍照后离线标注与导出，减少显微镜占用时间。
- 备注主要记录微塑料类别/种类，便于导出后筛选。
- 未知颗粒先快速分类；同一类出现多次时，挑代表做确认检测，不必逐个检测。

### 文档

- [中文使用说明](使用说明.md)
- [English user guide（英文使用说明）](User-Guide-EN.md)

### 注意

- exe 必须放在可写目录；如果公司安全策略禁止本地解压，请改用 Python 源码。
- Word 报告体积较大是预期行为：图片保持原始像素分辨率，没有缩略图。
- 工程文件保存的是图片绝对路径；移动图片或换电脑后需要重新添加图片。
- 程序未签名，Windows SmartScreen 可能提示。

### 许可证

GNU General Public License v3.0（GPL-3.0-only），见 [LICENSE](LICENSE)。

Copyright (C) 2026 MP Label-V1 contributors。

如果重新分发修改版或 exe，GPL 要求同时提供对应源码或源码获取方式。

### 引用

如果 MP Label-V1 对你的工作有帮助，引用信息见 [CITATION.cff](CITATION.cff)。
