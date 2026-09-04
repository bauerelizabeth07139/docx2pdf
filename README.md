# docx2pdf

在 Windows 上把 Word 文档（`.doc` / `.docx`）转成 PDF 的小工具。

自动探测并使用你机器上已有的办公软件，无需额外安装大型软件：

- **WPS Office**（COM `KWPS.Application`，需要已登录）
- **Microsoft Word**（COM `Word.Application`）
- **WPS 命令行转换**（`kwpsconvert.exe`，优先级最低，也需要登录）

Convert Word documents (.doc / .docx) to PDF on Windows by reusing an existing
WPS Office or Microsoft Word installation. No extra office software needed.

## 安装 / Install

```bash
pip install -r requirements.txt        # 实际上仅 Windows 需要 pywin32
# 或
pip install .                          # 安装为库 + `docx2pdf` 命令
```

需要本机已安装 **WPS Office** 或 **Microsoft Word** 之一。

## 用法 / Usage

命令有三种等效的启动方式：

| 方式 | 说明 |
| --- | --- |
| `docx2pdf.exe` | PyInstaller 打包的单文件 exe（见下方「打包」） |
| `docx2pdf` | `pip install .` 后安装的命令行入口 |
| `python -m docx2pdf` | 从源码仓库直接运行 |

```bash
# 单文件
docx2pdf convert 输入.docx -o 输出.pdf

# 指定后端
docx2pdf convert 输入.docx -o 输出.pdf --backend word

# 目录批量（-o 指向已存在目录时自动按批量处理）
docx2pdf convert ./docs/ -o ./out/

# 递归子目录
docx2pdf convert ./docs/ -o ./out/ --recursive

# 覆盖已存在文件
docx2pdf convert ./docs/ -o ./out/ --overwrite

# 列出可用后端
docx2pdf convert --list
```

### 子命令总览

| 命令 | 说明 |
| --- | --- |
| `convert` | 文档 → PDF（单文件或目录批量/递归） |
| `convert --list` | 检测并列出本机可用后端 |

### 作为库使用

```python
from docx2pdf.core import convert_file, batch_convert

convert_file("in.docx", "out.pdf", backend="auto")       # 单文件
batch_convert(["a.docx", "b.docx"], outdir="out/")        # 批量
```

## 原理 / How it works

1. 检测本机可用后端：优先 WPS COM → Word COM → WPS 命令行。
2. COM 路径：`Dispatch("KWPS.Application" 或 "Word.Application")`
   → `Documents.Open` → `ExportAsFixedFormat(..., 17)`，17 = PDF 格式。
3. 批量时逐个转换并将 PDF 写到目标目录。

> COM 为每个文件新建/退出一次应用实例，虽慢但稳定；如需更快可自行复用实例。

## 打包为可执行文件 / Build EXE

```bat
pip install pyinstaller
scripts\build_exe.bat
```

产物：`dist/docx2pdf.exe`（PyInstaller onefile，依赖 pywin32 随包打包）。
`launcher.py` 是打包入口，也可以直接 `python launcher.py convert --list` 运行。

## 测试 / Tests

无需 Office 软件即可运行（COM 后端已 mock）：

```bash
python -m unittest discover -s tests -v
```

## 目录 / Layout

```
docx2pdf/
  __init__.py   # 版本号
  __main__.py   # python -m docx2pdf 入口
  cli.py        # 命令行入口
  core.py       # 转换核心（后端探测 / COM / 批量）
launcher.py     # PyInstaller 打包入口
scripts/
  build_exe.bat
  make_sample.py
docs/
  usage.md      # 中文使用说明
tests/
  test_core.py
  test_cli.py
.github/workflows/ci.yml
```

## License

MIT