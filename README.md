# barcode_tool

从 Excel 指定列生成一维条形码图片的桌面工具。

## 功能

1. 选择表格文件（`.xlsx` / `.xls` / `.csv`）
2. 读取指定行（默认第 2 行）作为列名，选择要生成条形码的列
3. 把该列每个单元格的值生成一维条形码图片（PNG），导出到指定目录
4. 可选“显示文字”：在条码下方绘制人眼可读文本（自动使用系统字体，缺失则仅输出条码）

### 支持的条码类型
`code128`、`code39`、`ean13`、`ean8`、`upca`

> ⚠️ 一维条形码只能编码 **ASCII / 数字**。中文列（如“物品名称”）会被自动跳过，
> 请改用数字编码列（如“物品编码”）。`upca` 在底层由 `EAN13` 处理（接受 11/12 位数字）。

### 列名行 / 数据起始行
- 默认：列名在第 **2** 行，数据从第 **3** 行开始。
- 修改“列名行”时，“数据起始行”会自动同步为列名行的下一行。

## 使用

- **Windows 用户**：到 [Releases](https://github.com/Wflysea/table_to_barcode_rust/releases) 下载 `barcode_tool.exe`，
  双击即可运行，无需安装（GUI 界面）。
- **命令行自检**（无需界面，便于验证）：
  ```bash
  barcode_tool.exe --selftest <表格路径> [输出目录] [列名行]
  ```
  例：`barcode_tool.exe --selftest sample.csv out 2`

## 从源码构建

需要 [Rust 工具链](https://rustup.rs/)（stable）：

```bash
cargo build --release
# 产物：target/release/barcode_tool.exe
```

> 本项目通过 GitHub Actions 在 `windows-latest` 上自动编译，并在推送 `v*` 标签时
> 自动生成 GitHub Release 并上传 `barcode_tool_vX.Y.Z.exe`。

## 目录结构

```
src/main.rs              程序入口 + GUI + 核心逻辑
Cargo.toml              依赖与构建配置
sample.csv              命令行自检用的示例数据
.github/workflows/       自动构建与发布
```

## 许可证

MIT
