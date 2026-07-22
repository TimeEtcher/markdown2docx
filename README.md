# md2word — Markdown 转 Word 工具

可配置页面与字体样式的 Markdown → Word (.docx) 转换器，提供 GUI 和命令行两种用法，可打包为独立可执行文件。

## 功能

- 页面配置：纸张大小（A4/A5/Letter/自定义）、上下左右页边距
- 正文配置：中文字体、英文字体、字号、行距、首行缩进
- 标题配置：H1–H6 每级独立设置中/英文字体、字号、加粗、斜体、颜色、对齐、段前段后间距
- 支持标题、段落、加粗/斜体/行内代码、有序/无序列表（含嵌套）、代码块、引用块、本地图片、链接
- 表格：md 表格转 Word 真表格，可选网格（grid）/ 三线表（three_line）样式，表头居中加粗可配
- 公式支持：`$...$` 行内公式与 `$$...$$` 独立公式自动转为 Word 原生可编辑公式（OMML），无需手动重敲

## 使用

```bash
pip install -r requirements.txt

# GUI
python run.py

# 命令行
python run.py input.md -c config.example.yaml -o output.docx
```

配置文件为 YAML，所有字段可省略（省略则用默认值），完整示例见 [config.example.yaml](config.example.yaml)。

```yaml
body:
  font_en: "Times New Roman"
  font_zh: "宋体"
  font_size: 12
  line_spacing: 1.5
headings:
  h1: { font_zh: "黑体", font_size: 22, bold: true, align: center }
```

## 打包可执行文件

Linux：

```bash
bash scripts/build_linux.sh   # 产物: dist/md2word
```

Windows：在 Windows 机器上双击或在 cmd 中执行 `scripts\build_windows.bat`，产物为 `dist\md2word.exe`。

> PyInstaller 不支持跨平台交叉编译，Windows 版必须在 Windows 上构建。

## 项目结构

```
md2word/            Python 包
  config.py         默认配置 + YAML 合并
  styles.py         页面/字体（中英文分离）/行距应用
  converter.py      Markdown AST → docx
  cli.py            命令行入口
  gui.py            Tkinter GUI
run.py              启动入口（GUI / CLI / PyInstaller 共用）
scripts/            打包脚本
```
