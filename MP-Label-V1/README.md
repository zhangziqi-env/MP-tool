# MP Label-V1 1.0.0

MP Label-V1 是一个面向**微塑料显微图像**的 Windows 桌面标注工具。它把拍照后的标注、分类、深度管理、编号和导出集中在一个本地程序里完成，适合按深度/站位处理一批显微照片。

程序完全离线运行，图片和标注数据只保存在本机。提供英文版和中文版。

## 文件清单

| 文件 | 说明 |
| --- | --- |
| `MP Label-V1.exe` | 英文版单文件程序，免安装 |
| `MP Label-V1_中文版.exe` | 中文版单文件程序，免安装 |
| `MP Label-V1.py` | 英文版 Python 源码 |
| `MP Label-V1_中文版.py` | 中文版 Python 源码 |
| `使用说明.md` | 中英文使用说明（英文在前，中文在后） |
| `requirements.txt` | 运行 Python 源码所需依赖 |
| `LICENSE` | MIT 许可证 |

## 快速开始

### 使用 exe（推荐）

1. 把 `MP Label-V1.exe` 或 `MP Label-V1_中文版.exe` 放到**可写的文件夹**（桌面、下载目录即可；不要放 `C:\Program Files`）；
2. 双击运行；
3. Windows 提示“未知发布者”时选择“更多信息 → 仍要运行”。

exe 已包含 Tcl/Tk、Pillow、openpyxl 和 python-docx 模板，不需要安装 Python。启动时会在 exe 同目录临时解压 `_MEI...`，正常关闭会自动清理；强制结束可能留下该目录，手动删除即可。

### 使用 Python 源码

```bat
python -m pip install -r requirements.txt
python "MP Label-V1.py"
:: 中文版
python "MP Label-V1_中文版.py"
```

需要 Windows 10/11、Python 3.10+，且 Python 包含 Tcl/Tk。

## 功能

- 单张标注：类型、判断（MP / NoMP / Test）、Test 子类型、颜色、备注。
- 颜色记录：透明、白色、黑色、蓝色、红色、黄色、绿色、棕色、灰色、其他。
- 矩形/多边形套索批量标注，可保存套索区域供 Word 按判断导出。
- 按深度分组编号，支持 Renumber、Continue Num、Delete ID、Edit Remark、Edit Colour。
- 撤销/重做、JSON 工程保存/打开、自动保存恢复。
- Excel 导出：明细 7 列 + 深度 × 类型统计。
- Word 导出：按深度分组的全分辨率整图报告；按 MP / NoMP / Test 导出套索区域。
- 纯本地运行，数据不上传。

## 设计初衷

- 给所有样品留下完整的照片档案，便于长期保存和复查。
- 拍照后离线标注与导出，减少显微镜占用时间。
- 备注主要记录微塑料类别/种类，便于导出后筛选。
- 未知颗粒先快速分类；多次出现的同类颗粒导出后挑代表做确认检测，不必逐个检测。

## 使用说明

完整操作步骤、快捷键、未知颗粒记录方式和导出说明见 [使用说明.md](使用说明.md)。

## 注意

- exe 必须放在可写目录；如果公司安全策略禁止本地解压，请改用 Python 源码。
- Word 报告体积较大是预期行为：图片保持原始像素分辨率，没有缩略图。
- 工程文件保存的是图片绝对路径；移动图片或换电脑后需要重新添加图片。
- 程序未签名，Windows SmartScreen 可能提示。

## 许可证

本项目使用 [MIT License](LICENSE)。
