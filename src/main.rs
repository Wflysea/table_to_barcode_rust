//! Excel 指定列生成条形码工具（Rust 重写版）
//!
//! 功能与原始 Python 版本一致：
//! 1. 选择 Excel 文件
//! 2. 读取指定行（默认第 2 行）作为列名，选择条形码列
//! 3. 把该列每个单元格的值生成一维条形码图片，导出到指定目录
//!
//! 支持的条码类型：code128, code39, ean13, ean8, upca
//! （一维条码只能编码 ASCII / 数字，中文列会被跳过）

// 发布版（release）编译为 Windows GUI 子系统，避免双击运行时弹出黑色控制台窗口。
// 调试版（cargo run）保留控制台，方便查看 println! 输出。
#![cfg_attr(not(debug_assertions), windows_subsystem = "windows")]

use std::collections::HashMap;
use std::path::{Path, PathBuf};
use std::sync::mpsc::channel;
use std::sync::Arc;
use std::thread;

use anyhow::{anyhow, Context, Result};
use barcoders::generators::image::Image as BarcodeImage;
use calamine::{open_workbook_auto, Data, Reader};
use eframe::egui;
use rfd::FileDialog;

/// 条码图片宽度默认值（像素）
const DEFAULT_BAR_WIDTH: u32 = 350;
/// 条码图片高度默认值（像素）
const DEFAULT_BAR_HEIGHT: u32 = 150;
/// 条码左右留白（边距）默认值（像素）：条码实际宽度 = 图片宽度 - 边距
const DEFAULT_BAR_MARGIN: u32 = 100;
/// 文字高度（像素）
const FONT_SIZE: f32 = 26.0;
/// 默认打印分辨率（DPI）：用于把像素换算成毫米，并写入 PNG 物理尺寸（pHYs）元数据
const DEFAULT_DPI: u32 = 120;
/// 默认最小模块（最窄条）宽度（毫米）：低于此值会自动放大像素，保证条码可扫描
const DEFAULT_MIN_MODULE_MM: f32 = 0.4;

/// 支持的条码类型
const BARCODE_TYPES: &[&str] = &["code128", "code39", "ean13", "ean8", "upca"];

fn is_digit_only(bt: &str) -> bool {
    matches!(bt, "ean13" | "ean8" | "upca")
}

/// 把单元格值转换为干净的字符串（整数不出现小数点）
fn cell_text(d: &Data) -> String {
    format!("{}", d).trim().to_string()
}

/// 生成安全的文件名片段
fn safe_name(s: &str) -> String {
    s.trim()
        .replace(['\\', '/', ':', '*', '?', '"', '<', '>', '|'], "_")
        .chars()
        .take(80)
        .collect()
}

/// 定位一个可用的 TTF/TTC 字体（用于条形码下方的人眼可读文字）
fn find_font() -> Option<Vec<u8>> {
    let candidates = [
        "C:\\Windows\\Fonts\\arial.ttf",
        "C:\\Windows\\Fonts\\segoeui.ttf",
        "C:\\Windows\\Fonts\\calibri.ttf",
        "C:\\Windows\\Fonts\\msyh.ttc",
        "/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf",
        "/System/Library/Fonts/Supplemental/Arial.ttf",
    ];
    for c in candidates {
        if let Ok(bytes) = std::fs::read(c) {
            return Some(bytes);
        }
    }
    None
}

/// 定位系统中文字体，注册到 egui，保证中文 UI 正常显示
fn register_cjk_font(ctx: &egui::Context) {
    let candidates = [
        "C:\\Windows\\Fonts\\msyh.ttc",
        "C:\\Windows\\Fonts\\msyh.ttf",
        "C:\\Windows\\Fonts\\simsun.ttc",
        "C:\\Windows\\Fonts\\simhei.ttf",
        "/usr/share/fonts/truetype/noto/NotoSansCJK-Regular.ttc",
        "/usr/share/fonts/opentype/noto/NotoSansCJK-Regular.ttc",
        "/System/Library/Fonts/Supplemental/STHeiti Light.ttc",
    ];
    for p in candidates {
        if let Ok(bytes) = std::fs::read(p) {
            let mut fonts = egui::FontDefinitions::default();
            fonts.font_data.insert(
                "cjk".to_owned(),
                Arc::new(egui::FontData::from_owned(bytes)),
            );
            if let Some(proportional) = fonts.families.get_mut(&egui::FontFamily::Proportional) {
                proportional.push("cjk".to_owned());
            }
            ctx.set_fonts(fonts);
            break;
        }
    }
}

/// 读取整个表格为「行→列→文本」的二维结构（统一处理 Excel 与 CSV）。
/// - .csv 走自写的 read_table_csv（自动嗅探分隔符 + GBK/GB18030 中文解码）
/// - 其它（xlsx/xls）走 calamine
fn load_table(path: &str) -> Result<Vec<Vec<String>>> {
    if path.to_ascii_lowercase().ends_with(".csv") {
        return read_table_csv(path);
    }
    let mut wb = open_workbook_auto(path)
        .with_context(|| format!("无法打开文件: {}", path))?;
    let range = wb
        .worksheet_range_at(0)
        .ok_or_else(|| anyhow!("工作簿中没有工作表"))??;
    let rows: Vec<&[Data]> = range.rows().collect();
    Ok(rows
        .iter()
        .map(|r| r.iter().map(cell_text).collect())
        .collect())
}

/// 把文件字节解码为 UTF-8 字符串：优先按 UTF-8，失败则回退到 GBK/GB18030（中文 CSV 常见）
fn decode_bytes(bytes: &[u8]) -> String {
    let b = if bytes.starts_with(b"\xef\xbb\xbf") {
        &bytes[3..]
    } else {
        bytes
    };
    match std::str::from_utf8(b) {
        Ok(s) => s.to_string(),
        Err(_) => encoding_rs::GB18030.decode_without_bom_handling(b).0.to_string(),
    }
}

/// 嗅探 CSV 分隔符：扫描首个非空行，选择出现次数最多的候选（, \t ; |），默认逗号
fn detect_delimiter(text: &str) -> u8 {
    let first_line = text.lines().find(|l| !l.trim().is_empty()).unwrap_or("");
    let candidates: &[u8] = &[b',', b'\t', b';', b'|'];
    let mut best = b',';
    let mut best_count = 0usize;
    for &d in candidates {
        let c = first_line.as_bytes().iter().filter(|&&x| x == d).count();
        if c > best_count {
            best_count = c;
            best = d;
        }
    }
    best
}

/// 读取 CSV 文件为二维文本表（支持带引号字段、GBK/GB18030 编码、任意常见分隔符）
fn read_table_csv(path: &str) -> Result<Vec<Vec<String>>> {
    let bytes = std::fs::read(path).with_context(|| format!("无法读取文件: {}", path))?;
    let text = decode_bytes(&bytes);
    let delim = detect_delimiter(&text);
    let mut rdr = csv::ReaderBuilder::new()
        .delimiter(delim)
        .flexible(true)
        .has_headers(false)
        .from_reader(text.as_bytes());
    let mut rows = Vec::new();
    for rec in rdr.records() {
        let rec = rec.map_err(|e| anyhow!("CSV 解析失败: {}", e))?;
        rows.push(rec.iter().map(|s| s.to_string()).collect());
    }
    Ok(rows)
}

/// 读取工作表的指定行作为表头（header_row 为 1 基行号）
fn read_headers(path: &str, header_row: usize) -> Result<Vec<String>> {
    let rows = load_table(path)?;
    if header_row == 0 || header_row > rows.len() {
        return Ok(Vec::new());
    }
    let row = &rows[header_row - 1];
    Ok(row
        .iter()
        .enumerate()
        .map(|(i, c)| {
            let t = c.trim().to_string();
            if t.is_empty() {
                format!("列{}", i + 1)
            } else {
                t
            }
        })
        .collect())
}

/// 读取指定列（从 start_row 行起）的非空单元格，返回 (Excel行号, 文本)
fn read_column(path: &str, col_index: usize, start_row: usize) -> Result<Vec<(usize, String)>> {
    let rows = load_table(path)?;
    let mut out = Vec::new();
    for (i, row) in rows.iter().enumerate() {
        let excel_row = i + 1;
        if excel_row < start_row {
            continue;
        }
        if col_index >= row.len() {
            continue;
        }
        let t = row[col_index].trim().to_string();
        if t.is_empty() {
            continue;
        }
        out.push((excel_row, t));
    }
    Ok(out)
}

/// 统计指定列：返回 (非空数量, 含非ASCII字符数量)
fn analyze_column(path: &str, col_index: usize, start_row: usize) -> Result<(usize, usize)> {
    let cells = read_column(path, col_index, start_row)?;
    let non_empty = cells.len();
    let non_ascii = cells
        .iter()
        .filter(|(_, v)| v.chars().any(|c| !c.is_ascii()))
        .count();
    Ok((non_empty, non_ascii))
}

/// 把数据编码成指定类型的条形码二进制结构
fn encode_barcode(btype: &str, data: &str) -> Result<Vec<u8>> {
    let encoded = match btype {
        "code128" => {
            // Code128 必须指定起始字符集；用 Ɓ (U+0181) 选择字符集 B，可编码常见 ASCII
            let d = format!("\u{0181}{}", data);
            barcoders::sym::code128::Code128::new(&d)?.encode()
        }
        "code39" => barcoders::sym::code39::Code39::new(data)?.encode(),
        "ean13" => barcoders::sym::ean13::EAN13::new(data)?.encode(),
        "ean8" => barcoders::sym::ean8::EAN8::new(data)?.encode(),
        // UPC-A 是 EAN-13 的子集，barcoders 用 EAN13 处理 11/12 位数字
        "upca" => barcoders::sym::ean13::EAN13::new(data)?.encode(),
        _ => return Err(anyhow!("不支持的条码类型: {}", btype)),
    };
    Ok(encoded)
}

/// 在条形码图片下方绘制人眼可读文字，返回合成后的图像
fn compose_with_text(
    bars_png: &[u8],
    text: &str,
    font_data: &Option<Vec<u8>>,
) -> Result<image::RgbaImage> {
    let font_data = font_data.as_ref().ok_or_else(|| anyhow!("无可用的字体"))?;
    let font = ab_glyph::FontRef::try_from_slice(font_data)
        .map_err(|e| anyhow!("字体加载失败: {}", e))?;
    let bars = image::load_from_memory(bars_png)
        .context("解析条码图片失败")?
        .to_rgba8();
    let (w, h) = (bars.width(), bars.height());
    let scale = ab_glyph::PxScale {
        x: FONT_SIZE,
        y: FONT_SIZE,
    };
    let text_h = FONT_SIZE as u32 + 12;
    let mut out = image::RgbaImage::new(w, h + text_h);
    for p in out.pixels_mut() {
        *p = image::Rgba([255u8, 255, 255, 255]);
    }
    image::imageops::replace(&mut out, &bars, 0, 0);
    let tw = imageproc::drawing::text_size(scale, &font, text).0 as i32;
    let x = ((w as i32 - tw) / 2).max(0);
    imageproc::drawing::draw_text_mut(
        &mut out,
        image::Rgba([0u8, 0, 0, 255]),
        x,
        h as i32 + 4,
        scale,
        &font,
        text,
    );
    Ok(out)
}

/// 把 RGBA 图像编码为带 DPI 物理尺寸元数据（PNG pHYs 块）的 PNG 字节。
/// dpi 仅用于写入元数据（打印时按此分辨率换算物理尺寸），不缩放图像本身。
/// 像素/米 = DPI × 1000 / 25.4 = DPI × 10000 / 254。
///
/// 注意：png::Encoder 会“取得”写入目标的所有权，finish() 后无法再取回 Cursor，
/// 因此这里把 `&mut Vec<u8>` 作为写入目标（借用而非拥有），finish 后字节留在 Vec 中。
fn encode_png_with_dpi(img: &image::RgbaImage, dpi: u32) -> Result<Vec<u8>> {
    let ppm = ((dpi as f64) * 10000.0 / 254.0).round() as u32;
    let mut buf: Vec<u8> = Vec::new();
    {
        let mut encoder = png::Encoder::new(&mut buf, img.width(), img.height());
        encoder.set_color(png::ColorType::Rgba);
        encoder.set_depth(png::BitDepth::Eight);
        encoder.set_pixel_dims(Some(png::PixelDimensions {
            xppu: ppm,
            yppu: ppm,
            unit: png::Unit::Meter,
        }));
        let mut writer = encoder
            .write_header()
            .map_err(|e| anyhow!("PNG 编码失败: {}", e))?;
        writer
            .write_image_data(img.as_raw())
            .map_err(|e| anyhow!("PNG 写入失败: {}", e))?;
        writer
            .finish()
            .map_err(|e| anyhow!("PNG 收尾失败: {}", e))?;
    }
    Ok(buf)
}

/// 把 RGBA 图像保存为带 DPI 物理尺寸元数据的 PNG 文件。
fn save_png_with_dpi(img: &image::RgbaImage, path: &Path, dpi: u32) -> Result<()> {
    let bytes = encode_png_with_dpi(img, dpi)?;
    std::fs::write(path, bytes).with_context(|| format!("写入文件失败: {}", path.display()))?;
    Ok(())
}

/// 生成条形码主函数。
/// 返回 (成功数量, 跳过列表[(行号, 值, 原因)], 实际最小模块宽度(mm))
fn generate_barcodes(
    path: &str,
    col_index: usize,
    out_dir: &str,
    btype: &str,
    with_text: bool,
    start_row: usize,
    bar_width: u32,
    bar_height: u32,
    bar_margin: u32,
    dpi: u32,
    min_module_mm: f32,
    font_data: &Option<Vec<u8>>,
) -> Result<(usize, Vec<(usize, String, String)>, f32)> {
    std::fs::create_dir_all(out_dir).context("创建导出目录失败")?;
    let cells = read_column(path, col_index, start_row)?;
    let mut count = 0;
    let mut skipped: Vec<(usize, String, String)> = Vec::new();
    let mut used: HashMap<String, u32> = HashMap::new();
    // 记录整批生成中实际的最小模块（最窄条）像素宽度，用于换算成毫米回显给用户
    let mut min_module_px_actual: u32 = u32::MAX;

    for (row, value) in cells {
        // 数字型条码长度校验
        let data = if is_digit_only(btype) {
            let digits: String = value.chars().filter(|c| c.is_ascii_digit()).collect();
            let valid = match btype {
                "ean8" => digits.len() == 7 || digits.len() == 8,
                "upca" => digits.len() == 11 || digits.len() == 12,
                _ => digits.len() == 12 || digits.len() == 13,
            };
            if !valid {
                skipped.push((row, value.clone(), "数字长度不符合该条码类型要求".into()));
                continue;
            }
            digits
        } else {
            value.clone()
        };

        let encoded = match encode_barcode(btype, &data) {
            Ok(e) => e,
            Err(e) => {
                skipped.push((row, value.clone(), format!("生成失败: {}", e)));
                continue;
            }
        };

        let png = BarcodeImage::png(bar_height);
        let raw = match png.generate(&encoded[..]) {
            Ok(b) => b,
            Err(e) => {
                skipped.push((row, value.clone(), format!("保存失败: {}", e)));
                continue;
            }
        };
        // 将生成的条码缩放到「图片宽度 - 边距」的实际宽度，再水平居中贴到
        // 一张宽为 bar_width 的白色画布上，从而得到左右各 (边距/2) 的留白。
        // 同时保证最窄条（模块）的物理宽度 >= min_module_mm（按 dpi 换算）。
        let bytes = {
            let bars = image::load_from_memory(&raw)
                .context("解析条码图片失败")?
                .to_rgba8();
            // barcoders 生成的位图宽度为「模块数」（每模块 1px），据此可得模块像素宽度
            let num_modules = encoded.len() as u32;
            let inner_w_user = bar_width.saturating_sub(bar_margin).max(1);
            // 满足最小物理宽度所需的最小模块像素宽度（向上取整，至少 1px）
            let min_module_px = ((min_module_mm * dpi as f32 / 25.4).ceil().max(1.0)) as u32;
            // 用户设定意图下的模块像素宽度（取上限），再与最小要求取较大值
            let module_px = if num_modules > 0 {
                let desired = ((inner_w_user as f32 / num_modules as f32).ceil()).max(1.0) as u32;
                desired.max(min_module_px)
            } else {
                min_module_px
            };
            min_module_px_actual = min_module_px_actual.min(module_px);
            // 让实际宽度为模块数的整数倍，保证每个模块等宽（扫描更可靠）
            let inner_w = num_modules.saturating_mul(module_px).max(1);
            let bar_width_actual = inner_w + bar_margin;
            let bar_img = image::imageops::resize(
                &bars,
                inner_w,
                bar_height,
                image::imageops::FilterType::Nearest,
            );
            let mut canvas = image::RgbaImage::new(bar_width_actual, bar_height);
            for p in canvas.pixels_mut() {
                *p = image::Rgba([255u8, 255, 255, 255]);
            }
            let offset_x = (bar_width_actual.saturating_sub(inner_w)) / 2;
            image::imageops::replace(&mut canvas, &bar_img, offset_x as i64, 0);
            // 写入带 DPI 物理尺寸元数据的 PNG（保证打印尺寸正确）
            encode_png_with_dpi(&canvas, dpi).context("编码条码图片失败")?
        };

        let base = safe_name(&value);
        let c = used.entry(base.clone()).or_insert(0);
        let name = if *c == 0 {
            base.clone()
        } else {
            format!("{}_{}", base, *c)
        };
        *c += 1;

        let fpath = Path::new(out_dir).join(format!("{}.png", name));
        if with_text {
            match compose_with_text(&bytes, &value, font_data) {
                Ok(img) => {
                    save_png_with_dpi(&img, &fpath, dpi)?;
                }
                Err(_) => {
                    std::fs::write(&fpath, bytes)?;
                }
            }
        } else {
            std::fs::write(&fpath, bytes)?;
        }
        count += 1;
    }
    // 整批实际最小模块宽度（毫米）= 最小模块像素 / dpi × 25.4
    let min_module_mm_actual = if min_module_px_actual == u32::MAX {
        0.0
    } else {
        min_module_px_actual as f32 * 25.4 / dpi as f32
    };
    Ok((count, skipped, min_module_mm_actual))
}

/// 无界面自检：用于验证生成流程是否可用
fn run_selftest(excel: &str, out_arg: &str, header_row: usize) -> Result<String> {
    if excel.is_empty() {
        return Err(anyhow!("用法: --selftest <excel路径> [输出目录] [列名行]"));
    }
    let headers = read_headers(excel, header_row)?;
    let mut log = vec![format!("[INFO] 列名(第{}行): {:?}", header_row, headers)];
    if headers.is_empty() {
        return Err(anyhow!("[ERROR] 未读取到列名，请检查列名行设置。"));
    }
    // 优先选“物品编码”，否则选第一个全为 ASCII 的非空列
    let target = if headers.iter().any(|h| h == "物品编码") {
        "物品编码".to_string()
    } else {
        let mut found = None;
        for h in &headers {
            let idx = headers.iter().position(|x| x == h).unwrap();
            let (ne, na) = analyze_column(excel, idx, header_row + 1)?;
            if ne > 0 && na == 0 {
                found = Some(h.clone());
                break;
            }
        }
        found.unwrap_or_else(|| headers[0].clone())
    };
    let idx = headers.iter().position(|h| h == &target).unwrap();
    let (ne, na) = analyze_column(excel, idx, header_row + 1)?;
    log.push(format!(
        "[INFO] 选用列【{}】(索引{})，非空 {} 个，其中中文/非ASCII {} 个",
        target, idx, ne, na
    ));

    let out_dir = if out_arg.is_empty() {
        Path::new(excel)
            .parent()
            .map(|p| p.to_path_buf())
            .unwrap_or_else(|| PathBuf::from("."))
            .join("selftest_条形码输出")
            .to_string_lossy()
            .to_string()
    } else {
        out_arg.to_string()
    };

    let (count, skipped, min_mm) = generate_barcodes(
        excel,
        idx,
        &out_dir,
        "code128",
        true,
        header_row + 1,
        DEFAULT_BAR_WIDTH,
        DEFAULT_BAR_HEIGHT,
        DEFAULT_BAR_MARGIN,
        DEFAULT_DPI,
        DEFAULT_MIN_MODULE_MM,
        &find_font(),
    )?;
    log.push(format!(
        "[RESULT] 成功 {} 张，跳过 {} 张；最小模块宽度 ≈ {:.3} mm（要求 ≥ {:.2} mm @ {} DPI）",
        count, skipped.len(), min_mm, DEFAULT_MIN_MODULE_MM, DEFAULT_DPI
    ));
    for s in skipped.iter().take(10) {
        log.push(format!("   跳过 行{}: {:?} -> {}", s.0, s.1, s.2));
    }
    log.push(format!("[DONE] 输出目录: {}", out_dir));
    Ok(log.join("\n"))
}

// ----------------------- GUI -----------------------

struct AppState {
    excel_path: String,
    out_dir: String,
    headers: Vec<String>,
    col_idx: usize,
    type_idx: usize,
    with_text: bool,
    header_row: usize,
    start_row: usize,
    bar_width: u32,
    bar_height: u32,
    bar_margin: u32,
    dpi: u32,
    min_module_mm: f32,
    status: String,
    busy: bool,
    font_data: Option<Vec<u8>>,
    rx: Option<std::sync::mpsc::Receiver<GenOutcome>>,
}

struct GenOutcome {
    status: String,
    open: bool,
}

impl AppState {
    fn new(cc: &eframe::CreationContext<'_>) -> Self {
        register_cjk_font(&cc.egui_ctx);
        // 使用浅色高对比主题，避免默认深色（近黑）背景导致文字难以辨认
        let mut visuals = egui::Visuals::light();
        visuals.window_fill = egui::Color32::from_rgb(250, 250, 252);
        visuals.panel_fill = egui::Color32::from_rgb(240, 242, 246);
        cc.egui_ctx.set_visuals(visuals);
        Self {
            excel_path: String::new(),
            out_dir: String::new(),
            headers: Vec::new(),
            col_idx: 0,
            type_idx: 0,
            with_text: true,
            header_row: 2,
            start_row: 3,
            bar_width: DEFAULT_BAR_WIDTH,
            bar_height: DEFAULT_BAR_HEIGHT,
            bar_margin: DEFAULT_BAR_MARGIN,
            dpi: DEFAULT_DPI,
            min_module_mm: DEFAULT_MIN_MODULE_MM,
            status: "请选择表格文件（Excel 或 CSV）并开始。".into(),
            busy: false,
            font_data: find_font(),
            rx: None,
        }
    }

    fn load_cols(&mut self) {
        if self.excel_path.is_empty() {
            self.status = "请先选择表格文件（Excel 或 CSV）。".into();
            return;
        }
        self.start_row = self.header_row + 1;
        match read_headers(&self.excel_path, self.header_row) {
            Ok(h) => {
                if h.is_empty() {
                    self.status = "未读取到列名，请检查列名行设置。".into();
                } else {
                    self.headers = h;
                    self.col_idx = 0;
                    self.status = format!("已读取 {} 个列，请选择条形码列。", self.headers.len());
                }
            }
            Err(e) => {
                self.status = format!("读取表头失败：{}", e);
            }
        }
    }

    fn generate(&mut self) {
        if self.excel_path.is_empty() {
            self.status = "请选择表格文件（Excel 或 CSV）。".into();
            return;
        }
        if self.headers.is_empty() {
            self.status = "请先读取并选择条形码列。".into();
            return;
        }
        if self.out_dir.is_empty() {
            self.status = "请选择条形码导出目录。".into();
            return;
        }
        let btype = BARCODE_TYPES[self.type_idx];
        let col_idx = self.col_idx;
        let path = self.excel_path.clone();
        let out = self.out_dir.clone();
        let header_row = self.header_row;
        let start_row = self.start_row;
        let with_text = self.with_text;
        let bar_width = self.bar_width;
        let bar_height = self.bar_height;
        let bar_margin = self.bar_margin;
        let dpi = self.dpi;
        let min_module_mm = self.min_module_mm;
        let font_data = self.font_data.clone();

        match analyze_column(&path, col_idx, start_row) {
            Ok((non_empty, non_ascii)) => {
                if non_empty == 0 {
                    self.status = format!(
                        "所选列从第 {} 行起没有可读数据。请检查数据起始行（列名在第 {} 行时，数据起始行应为 {}）。",
                        start_row, header_row, header_row + 1
                    );
                    return;
                }
                if non_ascii == non_empty {
                    self.status = "注意：所选列全部含中文/非ASCII字符，一维条码无法编码中文，将全部跳过。建议改用数字编码列（如“物品编码”）。".into();
                } else if non_ascii > 0 {
                    self.status = format!(
                        "注意：所选列有 {}/{} 个值含中文/非ASCII字符，这些会被跳过（一维条码无法编码中文）。",
                        non_ascii, non_empty
                    );
                }
            }
            Err(_) => {}
        }

        self.busy = true;
        self.status = "处理中，请稍候…".into();
        let (tx, rx) = channel::<GenOutcome>();
        self.rx = Some(rx);
        thread::spawn(move || {
            let res = generate_barcodes(
                &path, col_idx, &out, btype, with_text, start_row, bar_width, bar_height, bar_margin, dpi, min_module_mm, &font_data,
            );
            let outcome = match res {
                Ok((count, skipped, min_mm)) => {
                    let mut s = format!("✅ 生成完成：成功 {} 张", count);
                    if !skipped.is_empty() {
                        s.push_str(&format!("，跳过 {} 张", skipped.len()));
                        let sample: Vec<String> = skipped
                            .iter()
                            .take(5)
                            .map(|(r, v, reason)| format!("  行{}: {} -> {}", r, v, reason))
                            .collect();
                        s.push_str("\n⚠️ 部分被跳过（一维条码无法编码中文，请改用数字编码列）。\n");
                        s.push_str(&sample.join("\n"));
                    }
                    s.push_str(&format!(
                        "\n最小模块宽度 ≈ {:.3} mm（要求 ≥ {:.2} mm @ {} DPI）",
                        min_mm, min_module_mm, dpi
                    ));
                    s.push_str(&format!("\n导出目录：{}", out));
                    GenOutcome {
                        status: s,
                        open: true,
                    }
                }
                Err(e) => GenOutcome {
                    status: format!("❌ 出错：{}", e),
                    open: false,
                },
            };
            let _ = tx.send(outcome);
        });
    }
}

impl eframe::App for AppState {
    fn ui(&mut self, ui: &mut egui::Ui, _frame: &mut eframe::Frame) {
        ui.label(
            "选择表格文件（Excel 或 CSV）→ 读取列名并选择条形码列 → 选择导出目录 → 生成。\n（列名行/数据起始行可调：默认列名在第 2 行、数据从第 3 行开始）",
        );

        ui.horizontal(|ui| {
            ui.label("表格文件");
            ui.text_edit_singleline(&mut self.excel_path);
            if ui.button("浏览…").clicked() {
                if let Some(p) = FileDialog::new()
                    .add_filter("表格文件", &["xlsx", "xls", "csv"])
                    .pick_file()
                {
                    self.excel_path = p.to_string_lossy().to_string();
                    self.load_cols();
                }
            }
        });

        ui.horizontal(|ui| {
            ui.label("条形码列");
            if self.headers.is_empty() {
                ui.label("（请先读取列名）");
            } else {
                egui::ComboBox::from_id_salt("barcode_col")
                    .show_index(ui, &mut self.col_idx, self.headers.len(), |i| {
                        self.headers.get(i).cloned().unwrap_or_default()
                    });
            }
            if ui.button("读取列名").clicked() {
                self.load_cols();
            }
        });

        ui.horizontal(|ui| {
            ui.label("条码类型");
            egui::ComboBox::from_id_salt("barcode_type").show_index(ui, &mut self.type_idx, BARCODE_TYPES.len(), |i| {
                BARCODE_TYPES[i].to_string()
            });
        });

        ui.horizontal(|ui| {
            ui.label("条码宽度(px)");
            ui.add(egui::DragValue::new(&mut self.bar_width).range(50..=2000));
            ui.label("条码高度(px)");
            ui.add(egui::DragValue::new(&mut self.bar_height).range(20..=1000));
            ui.label("边距(px)");
            ui.add(egui::DragValue::new(&mut self.bar_margin).range(0..=1000));
        });

        ui.horizontal(|ui| {
            ui.label("打印分辨率(DPI)");
            ui.add(egui::DragValue::new(&mut self.dpi).range(72..=1200));
            ui.label("最小模块宽(mm)");
            ui.add(egui::DragValue::new(&mut self.min_module_mm).range(0.1..=5.0).speed(0.05));
            ui.label("（最窄条物理宽度下限，不足自动放大）");
        });

        ui.checkbox(&mut self.with_text, "显示文字");

        ui.horizontal(|ui| {
            ui.label("列名行");
            let old = self.header_row;
            ui.add(egui::DragValue::new(&mut self.header_row).range(1..=100));
            ui.label("数据起始行");
            ui.add(egui::DragValue::new(&mut self.start_row).range(1..=100));
            if self.header_row != old {
                self.start_row = self.header_row + 1;
            }
        });

        ui.horizontal(|ui| {
            ui.label("导出目录");
            ui.text_edit_singleline(&mut self.out_dir);
            if ui.button("选择目录…").clicked() {
                if let Some(p) = FileDialog::new().pick_folder() {
                    self.out_dir = p.to_string_lossy().to_string();
                }
            }
        });

        ui.horizontal(|ui| {
            if ui.button("▶ 生成条形码").clicked() {
                self.generate();
            }
            if ui.button("打开导出目录").clicked() {
                if !self.out_dir.is_empty() {
                    let _ = open::that(&self.out_dir);
                }
            }
        });

        ui.label(&self.status);

        // 接收后台生成结果
        if let Some(rx) = &self.rx {
            if let Ok(outcome) = rx.try_recv() {
                self.busy = false;
                self.rx = None;
                self.status = outcome.status;
                if outcome.open {
                    let _ = open::that(&self.out_dir);
                }
            }
        }
        if self.busy {
            ui.ctx().request_repaint();
        }
    }
}

fn main() -> Result<(), eframe::Error> {
    let args: Vec<String> = std::env::args().collect();
    if let Some(pos) = args.iter().position(|a| a == "--selftest") {
        let excel = args.get(pos + 1).cloned().unwrap_or_default();
        let out = args.get(pos + 2).cloned().unwrap_or_default();
        let hrow = args
            .get(pos + 3)
            .and_then(|s| s.parse::<usize>().ok())
            .unwrap_or(2);
        match run_selftest(&excel, &out, hrow) {
            Ok(msg) => {
                println!("{}", msg);
                // windows 子系统下无控制台，同时写入文件便于排查
                let _ = std::fs::write("selftest_log.txt", &msg);
            }
            Err(e) => {
                eprintln!("selftest error: {}", e);
                let _ = std::fs::write("selftest_log.txt", format!("selftest error: {}", e));
                std::process::exit(1);
            }
        }
        return Ok(());
    }

    let options = eframe::NativeOptions {
        viewport: egui::ViewportBuilder::default().with_inner_size([640.0, 560.0]),
        ..Default::default()
    };
    eframe::run_native(
        "表格转条形码工具 (Rust)",
        options,
        Box::new(|cc| Ok(Box::new(AppState::new(cc)))),
    )
}
