# md2word — Markdown 转 Word 工具

可配置页面与字体样式的 Markdown → Word (.docx) 转换器，提供 GUI 和命令行两种用法，可打包为独立可执行文件。

> **想直接体验？** 无需安装 Python，可直接下载仓库中的 [`dist/md2word.exe`](dist/md2word.exe) 双击运行（Windows 版）。

## 功能

- **页面配置**：纸张大小（A4/A5/Letter/自定义）、上下左右页边距
- **正文配置**：中文字体、英文字体、字号、行距、首行缩进两字符
- **标题配置**：H1–H6 每级独立设置中文字体、字号、加粗、对齐、段前/段后空行、行距
- **字号选择**：GUI 中使用 Word 中文字号（初号、小初、一号、小一、二号、小二、三号、小三、四号、小四、五号、小五、六号、小六、七号、八号），自动映射为磅值
- **样式保存/加载**：GUI 内一键保存当前配置为 YAML，下次启动自动加载默认配置，也可手动加载任意配置文件
- **内容支持**：标题、段落、加粗/斜体/行内代码、有序/无序列表（含嵌套）、代码块、引用块、本地图片、链接
- **表格**：Markdown 表格转 Word 真表格，可选网格（grid）/ 三线表（three_line）样式，表头居中加粗可配
- **公式支持**：`$...$` 行内公式与 `$$...$$` 独立公式自动转为 Word 原生可编辑公式（OMML）
- **参考文献交叉引用**：
  - 自动识别 Markdown 中标题为 `参考文献` 或 `References` 的章节
  - 正文中的 `[1]`、`[1-2]`、`[1,4]` 等引用渲染为带中括号的上标角标
  - 引用数字转为 Word `REF` 域交叉引用，点击（Ctrl+Click）可跳转到参考文献列表对应条目
  - 参考文献条目以 `[N] ...` 格式识别，支持 `[M-N]` 范围批量收集编号

## 使用

### GUI（默认）

```bash
pip install -r requirements.txt
python run.py
```

或直接运行：

```bash
python -m md2word
```

GUI 界面分为：

- **文件**：选择 Markdown 输入文件与输出 docx 路径
- **样式配置**：页面 / 正文 / 标题 / 代码·表格 四个标签页，字体和字号均为下拉选择
- **保存配置 / 加载配置**：把当前界面设置导出为 YAML，或从 YAML 导入；默认自动加载 `~/.md2word.yaml`

### 命令行

```bash
python run.py input.md -c config.example.yaml -o output.docx
```

配置文件为 YAML，所有字段可省略（省略则用默认值），完整示例见 [config.example.yaml](config.example.yaml)。

```yaml
body:
  font_en: "Times New Roman"
  font_zh: "宋体"
  font_size: 12          # 五号=10.5, 小四=12, 四号=14...
  line_spacing: 1.5
  first_line_indent: true

headings:
  h1:
    font_zh: "黑体"
    font_size: 22
    bold: true
    align: center
    space_before_lines: 2.0   # 段前空行（以正文字号为基准）
    space_after_lines: 1.5
    line_spacing: 1.5
```

## 打包可执行文件

Linux：

```bash
bash scripts/build_linux.sh   # 产物: dist/md2word
```

Windows：在 Windows 机器上执行：

```bash
scripts\build_windows.bat     # 产物: dist\md2word.exe
```

> PyInstaller 不支持跨平台交叉编译，Windows 版必须在 Windows 上构建。
> 在 Anaconda 环境下打包时，请确保 `%PATH%` 包含 `Anaconda/Library/bin`，否则生成的 exe 可能缺少 DLL。

## 项目结构

```
md2word/            Python 包
  config.py         默认配置 + YAML 深合并
  styles.py         页面/字体（中英文分离）/行距应用
  converter.py      Markdown AST → docx（含参考文献交叉引用）
  cli.py            命令行入口
  gui.py            Tkinter GUI（配置编辑 + 保存/加载）
run.py              启动入口（GUI / CLI / PyInstaller 共用）
scripts/            打包脚本
```
