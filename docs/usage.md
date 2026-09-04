# docx2pdf 使用说明

## 快速开始

命令可用 `docx2pdf.exe`（打包版）、`docx2pdf`（`pip install .` 后）或
`python -m docx2pdf`（源码目录内）任一种方式启动，下文以 `docx2pdf` 为例：

```bash
# 单文件
docx2pdf.exe convert 输入.docx -o 输出.pdf

# 指定后端（wps / word / wps_cli）
docx2pdf.exe convert 输入.docx -o 输出.pdf --backend word

# 目录批量（所有 docx 转 pdf，输出到 out/）
docx2pdf.exe convert ./资料/ -o ./out/

# 递归子目录
docx2pdf.exe convert ./资料/ -o ./out/ --recursive

# 覆盖已存在文件
docx2pdf.exe convert ./资料/ -o ./out/ --overwrite
```

## 查看可用后端

```bash
docx2pdf.exe convert --list
```

## 后端说明

| 后端 | COM 标识 | 来源 | 备注 |
| --- | --- | --- | --- |
| `wps` | `KWPS.Application` | 本机安装的 WPS Office | 需 WPS 已登录 |
| `word` | `Word.Application` | 本机安装的 Microsoft Word | 推荐，稳定 |
| `wps_cli` | — | `kwpsconvert.exe` | 命令行转换，同样需要 WPS 登录 |

`auto` 会依次探测 `wps` → `word` → `wps_cli`，选用第一个可用的。

## 环境要求

- Windows 10/11，装有 WPS Office 或 Microsoft Word
- 直接运行 exe 无需 Python；从源码运行时需要：
  ```bash
  pip install -r requirements.txt   # pywin32（COM 后端必需）
  ```

## 常见问题

1. **`No PDF backend available`**：本机未装 WPS/Word，或 COM 权限受限。
2. **批量很慢**：每个文件都会新起一次办公软件实例（COM），这是为了稳定；
   如需极速，请自行在 `core.py` 里复用同一个 `Application` 实例。
3. **wps_cli 需要登录**：WPS 转换命令行在未登录账户时会报 `Not logged in`，
   此时请改用 `--backend wps` 或 `--backend word`。