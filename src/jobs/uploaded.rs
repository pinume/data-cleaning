use std::collections::HashMap;
use std::collections::HashSet;
use std::path::{Path, PathBuf};

use crate::io::paths::list_xlsx_files;
use crate::io::xlsx_reader::{RawCell, SheetGrid, open_sheets};
use crate::model::{Column, ColumnType, DecimalScale, ProcessError, Row, Table, Value};

use super::{
    Category, Job, amount_value, cell_amount, cell_date_or_text, cell_datetime_or_text,
    cell_display, cell_text, check_duplicate_fingerprint, data_error, text_value,
};

/// 前25列两组数据组结构完全相同，按固定列位置读取。
const FRONT_LEN: usize = 25;
const FIELD_COUNT: usize = 58;
const TAIL_LEN: usize = FIELD_COUNT - FRONT_LEN;

const FRONT_HEADERS: [&str; FRONT_LEN] = [
    "实时清分UUID",
    "商户号",
    "商户名称",
    "订单号",
    "交易日期",
    "交易金额",
    "检索参考号",
    "模版类型",
    "状态",
    "描述",
    "提交时间",
    "更新时间",
    "终端号",
    "分店id",
    "分店名",
    "所在地区",
    "详细地址",
    "地区编码",
    "tel",
    "发票号码",
    "发票金额",
    "购买方名称",
    "图片1",
    "S/N码",
    "是否属于 AI 产品",
];

const FRONT_COLUMN_TYPES: [ColumnType; FRONT_LEN] = [
    ColumnType::Text,                            // 实时清分UUID
    ColumnType::Text,                            // 商户号
    ColumnType::Text,                            // 商户名称
    ColumnType::Text,                            // 订单号
    ColumnType::Date,                            // 交易日期
    ColumnType::Decimal(DecimalScale::Original), // 交易金额
    ColumnType::Text,                            // 检索参考号
    ColumnType::Text,                            // 模版类型
    ColumnType::Text,                            // 状态
    ColumnType::Text,                            // 描述
    ColumnType::DateTime,                        // 提交时间
    ColumnType::DateTime,                        // 更新时间
    ColumnType::Text,                            // 终端号
    ColumnType::Text,                            // 分店id
    ColumnType::Text,                            // 分店名
    ColumnType::Text,                            // 所在地区
    ColumnType::Text,                            // 详细地址
    ColumnType::Text,                            // 地区编码
    ColumnType::Text,                            // tel
    ColumnType::Text,                            // 发票号码
    ColumnType::Text, // 发票金额（源值含“3000,00”等非标准写法，按文本保留）
    ColumnType::Text, // 购买方名称
    ColumnType::Text, // 图片1（第一次出现）
    ColumnType::Text, // S/N码
    ColumnType::Text, // 是否属于 AI 产品
];

/// 第26列起两组数据组结构不同，且已知存在同一数据组内部的表头变体（如“电脑”工作表），
/// 因此按名称在“第26列起”的范围内查找，不按固定列号；范围限定可避免与第23列同名的
/// `图片1`混淆。
struct TailField {
    name: &'static str,
    ty: ColumnType,
    /// 已确认的候选源字段名；第一项为该数据组的标准名称，用于测试生成标准表头。
    synonyms: &'static [&'static str],
    /// 是否为必需字段；`false`表示已知有工作表变体缺少该字段（如“电脑”工作表无
    /// `airConditionerKitInfo`），缺失时留空而非终止处理。
    required: bool,
}

const APPLIANCE_TAIL: [TailField; TAIL_LEN] = [
    TailField {
        name: "图片1",
        ty: ColumnType::Text,
        synonyms: &["图片1"],
        required: true,
    },
    TailField {
        name: "图片2",
        ty: ColumnType::Text,
        synonyms: &["图片2"],
        required: true,
    },
    TailField {
        name: "图片3",
        ty: ColumnType::Text,
        synonyms: &["图片3"],
        required: true,
    },
    TailField {
        name: "图片4",
        ty: ColumnType::Text,
        synonyms: &["图片4"],
        required: true,
    },
    TailField {
        name: "img5",
        ty: ColumnType::Text,
        // “电脑”工作表中该字段名为“图片5”，与“家电”工作表的“img5”同义。
        synonyms: &["img5", "图片5"],
        required: true,
    },
    TailField {
        name: "img6",
        ty: ColumnType::Text,
        synonyms: &["img6"],
        required: true,
    },
    TailField {
        name: "img7",
        ty: ColumnType::Text,
        synonyms: &["img7"],
        required: true,
    },
    TailField {
        name: "img8",
        ty: ColumnType::Text,
        synonyms: &["img8"],
        required: true,
    },
    TailField {
        name: "img9",
        ty: ColumnType::Text,
        synonyms: &["img9"],
        required: true,
    },
    TailField {
        name: "img10",
        ty: ColumnType::Text,
        synonyms: &["img10"],
        required: true,
    },
    TailField {
        name: "img11",
        ty: ColumnType::Text,
        synonyms: &["img11"],
        required: true,
    },
    TailField {
        name: "img12",
        ty: ColumnType::Text,
        synonyms: &["img12"],
        required: true,
    },
    TailField {
        name: "img13",
        ty: ColumnType::Text,
        synonyms: &["img13"],
        required: true,
    },
    TailField {
        name: "img14",
        ty: ColumnType::Text,
        synonyms: &["img14"],
        required: true,
    },
    TailField {
        name: "img15",
        ty: ColumnType::Text,
        synonyms: &["img15"],
        required: true,
    },
    TailField {
        name: "签收时间",
        ty: ColumnType::DateTime,
        synonyms: &["签收时间"],
        required: true,
    },
    TailField {
        name: "remark",
        ty: ColumnType::Text,
        synonyms: &["remark"],
        required: true,
    },
    TailField {
        name: "EEG",
        ty: ColumnType::Text,
        synonyms: &["EEG"],
        required: true,
    },
    TailField {
        name: "物流单号",
        ty: ColumnType::Text,
        synonyms: &["物流单号"],
        required: true,
    },
    TailField {
        name: "erpOrderNum",
        ty: ColumnType::Text,
        synonyms: &["erpOrderNum"],
        required: true,
    },
    TailField {
        name: "ocrModify",
        ty: ColumnType::Text,
        synonyms: &["ocrModify"],
        required: true,
    },
    TailField {
        name: "modifyStatus",
        ty: ColumnType::Text,
        synonyms: &["modifyStatus"],
        required: true,
    },
    TailField {
        name: "introduceInvoiceFlag",
        ty: ColumnType::Text,
        synonyms: &["introduceInvoiceFlag"],
        required: true,
    },
    TailField {
        name: "是否交旧",
        ty: ColumnType::Text,
        synonyms: &["是否交旧"],
        required: true,
    },
    TailField {
        name: "是否自提",
        ty: ColumnType::Text,
        synonyms: &["是否自提"],
        required: true,
    },
    TailField {
        name: "receiverName",
        ty: ColumnType::Text,
        synonyms: &["receiverName"],
        required: true,
    },
    TailField {
        name: "productCode",
        ty: ColumnType::Text,
        synonyms: &["productCode"],
        required: true,
    },
    TailField {
        name: "subsideAmt",
        ty: ColumnType::Text,
        synonyms: &["subsideAmt"],
        required: true,
    },
    TailField {
        name: "productName",
        ty: ColumnType::Text,
        synonyms: &["productName"],
        required: true,
    },
    TailField {
        name: "交旧品类",
        ty: ColumnType::Text,
        synonyms: &["交旧品类"],
        required: true,
    },
    TailField {
        name: "收货地址是否农村地区",
        ty: ColumnType::Text,
        synonyms: &["收货地址是否农村地区"],
        required: true,
    },
    TailField {
        name: "airConditionerKitInfo",
        ty: ColumnType::Text,
        // “电脑”工作表没有此字段：缺失时留空，不终止处理。
        synonyms: &["airConditionerKitInfo"],
        required: false,
    },
    TailField {
        name: "开票日期",
        ty: ColumnType::Date,
        synonyms: &["开票日期"],
        required: true,
    },
];

const DIGITAL_TAIL: [TailField; TAIL_LEN] = [
    TailField {
        name: "IMEI1",
        ty: ColumnType::Text,
        synonyms: &["IMEI1"],
        required: true,
    },
    TailField {
        name: "IMEI2",
        ty: ColumnType::Text,
        synonyms: &["IMEI2"],
        required: true,
    },
    TailField {
        name: "图片1",
        ty: ColumnType::Text,
        synonyms: &["图片1"],
        required: true,
    },
    TailField {
        name: "图片2",
        ty: ColumnType::Text,
        synonyms: &["图片2"],
        required: true,
    },
    TailField {
        name: "图片3",
        ty: ColumnType::Text,
        synonyms: &["图片3"],
        required: true,
    },
    TailField {
        name: "图片4",
        ty: ColumnType::Text,
        synonyms: &["图片4"],
        required: true,
    },
    TailField {
        name: "图片5",
        ty: ColumnType::Text,
        synonyms: &["图片5"],
        required: true,
    },
    TailField {
        name: "图片6",
        ty: ColumnType::Text,
        synonyms: &["图片6"],
        required: true,
    },
    TailField {
        name: "img7",
        ty: ColumnType::Text,
        synonyms: &["img7"],
        required: true,
    },
    TailField {
        name: "img8",
        ty: ColumnType::Text,
        synonyms: &["img8"],
        required: true,
    },
    TailField {
        name: "img9",
        ty: ColumnType::Text,
        synonyms: &["img9"],
        required: true,
    },
    TailField {
        name: "img10",
        ty: ColumnType::Text,
        synonyms: &["img10"],
        required: true,
    },
    TailField {
        name: "img11",
        ty: ColumnType::Text,
        synonyms: &["img11"],
        required: true,
    },
    TailField {
        name: "img12",
        ty: ColumnType::Text,
        synonyms: &["img12"],
        required: true,
    },
    TailField {
        name: "img13",
        ty: ColumnType::Text,
        synonyms: &["img13"],
        required: true,
    },
    TailField {
        name: "img14",
        ty: ColumnType::Text,
        synonyms: &["img14"],
        required: true,
    },
    TailField {
        name: "img15",
        ty: ColumnType::Text,
        synonyms: &["img15"],
        required: true,
    },
    TailField {
        name: "签收时间",
        ty: ColumnType::DateTime,
        synonyms: &["签收时间"],
        required: true,
    },
    TailField {
        name: "remark",
        ty: ColumnType::Text,
        synonyms: &["remark"],
        required: true,
    },
    TailField {
        name: "物流单号",
        ty: ColumnType::Text,
        synonyms: &["物流单号"],
        required: true,
    },
    TailField {
        name: "erpOrderNum",
        ty: ColumnType::Text,
        synonyms: &["erpOrderNum"],
        required: true,
    },
    TailField {
        name: "ocrModify",
        ty: ColumnType::Text,
        synonyms: &["ocrModify"],
        required: true,
    },
    TailField {
        name: "modifyStatus",
        ty: ColumnType::Text,
        synonyms: &["modifyStatus"],
        required: true,
    },
    TailField {
        name: "introduceInvoiceFlag",
        ty: ColumnType::Text,
        synonyms: &["introduceInvoiceFlag"],
        required: true,
    },
    TailField {
        name: "是否交旧",
        ty: ColumnType::Text,
        synonyms: &["是否交旧"],
        required: true,
    },
    TailField {
        name: "是否自提",
        ty: ColumnType::Text,
        synonyms: &["是否自提"],
        required: true,
    },
    TailField {
        name: "receiverName",
        ty: ColumnType::Text,
        synonyms: &["receiverName"],
        required: true,
    },
    TailField {
        name: "productCode",
        ty: ColumnType::Text,
        synonyms: &["productCode"],
        required: true,
    },
    TailField {
        name: "subsideAmt",
        ty: ColumnType::Text,
        synonyms: &["subsideAmt"],
        required: true,
    },
    TailField {
        name: "productName",
        ty: ColumnType::Text,
        synonyms: &["productName"],
        required: true,
    },
    TailField {
        name: "oldExchangeType",
        ty: ColumnType::Text,
        synonyms: &["oldExchangeType"],
        required: true,
    },
    TailField {
        name: "收货地址是否农村地区",
        ty: ColumnType::Text,
        synonyms: &["收货地址是否农村地区"],
        required: true,
    },
    TailField {
        name: "开票日期",
        ty: ColumnType::Date,
        synonyms: &["开票日期"],
        required: true,
    },
];

// 固定列位置（1 基）。
const COL_UUID: u32 = 1;
const COL_MERCHANT_NO: u32 = 2;

struct UploadedConfig {
    category: Category,
    title: &'static str,
    output_stem: &'static str,
    merchant_no: &'static str,
    tail: &'static [TailField; TAIL_LEN],
}

const DIGITAL_CONFIG: UploadedConfig = UploadedConfig {
    category: Category::UploadedDigital,
    title: "数码已上传数据",
    output_stem: "已上传数码",
    merchant_no: "89813014812B06R",
    tail: &DIGITAL_TAIL,
};

const APPLIANCE_CONFIG: UploadedConfig = UploadedConfig {
    category: Category::UploadedAppliance,
    title: "家电、电脑已上传数据",
    output_stem: "已上传家电电脑",
    merchant_no: "89813015722APT1",
    tail: &APPLIANCE_TAIL,
};

pub struct UploadedJob(&'static UploadedConfig);

pub const UPLOADED_DIGITAL: UploadedJob = UploadedJob(&DIGITAL_CONFIG);
pub const UPLOADED_APPLIANCE: UploadedJob = UploadedJob(&APPLIANCE_CONFIG);

impl Job for UploadedJob {
    fn category(&self) -> Category {
        self.0.category
    }

    fn title(&self) -> &'static str {
        self.0.title
    }

    fn output_stem(&self) -> &'static str {
        self.0.output_stem
    }

    fn run(&self, input_dir: &Path) -> Result<Table, ProcessError> {
        run_uploaded(self.0, input_dir)
    }
}

/// 文件名须完整符合`MER_<商户号>_yyyymmddhhmmss_yjhx.xlsx`。
fn matches_filename(name: &str, merchant_no: &str) -> bool {
    let Some(stem) = name
        .strip_prefix("MER_")
        .and_then(|s| s.strip_suffix(".xlsx"))
    else {
        return false;
    };
    let Some(rest) = stem
        .strip_prefix(merchant_no)
        .and_then(|s| s.strip_prefix('_'))
    else {
        return false;
    };
    match rest.split_once('_') {
        Some((timestamp, suffix)) => {
            timestamp.len() == 14
                && timestamp.bytes().all(|b| b.is_ascii_digit())
                && suffix == "yjhx"
        }
        None => false,
    }
}

/// 在“第26列起”的范围内按名称查找某统一字段的实际列号：候选同义词中恰好一个出现时
/// 返回该列号；均未出现时返回`None`；多个同义词同时出现视为结构异常。
fn resolve_tail_column(tail_domain: &[String], synonyms: &[&str]) -> Result<Option<u32>, String> {
    let mut found = Vec::new();
    for &synonym in synonyms {
        if let Some(position) = tail_domain.iter().position(|name| name == synonym) {
            found.push(FRONT_LEN as u32 + position as u32 + 1);
        }
    }
    match found.as_slice() {
        [] => Ok(None),
        [column] => Ok(Some(*column)),
        _ => Err(format!("字段候选名称 {synonyms:?} 在表头中出现多个匹配")),
    }
}

/// 前25列须按固定列位置与名称完全一致；第26列起按名称在其范围内查找（第5.1节：
/// 数据组内部存在“电脑”等表头变体，不能仅按固定列号映射）。
fn resolve_columns(sheet: &SheetGrid, config: &UploadedConfig) -> Result<Vec<Option<u32>>, String> {
    let header = sheet.row_texts(2);
    if header.len() < FRONT_LEN
        || header[..FRONT_LEN]
            .iter()
            .map(String::as_str)
            .collect::<Vec<_>>()
            != FRONT_HEADERS
    {
        return Err(format!("第2行前{FRONT_LEN}列表头与规定字段及顺序不一致"));
    }

    let tail_domain = &header[FRONT_LEN..];
    let mut tail_columns = Vec::with_capacity(TAIL_LEN);
    for field in config.tail {
        let resolved = resolve_tail_column(tail_domain, field.synonyms)?;
        if resolved.is_none() && field.required {
            return Err(format!("缺少必需字段：{}", field.name));
        }
        tail_columns.push(resolved);
    }
    Ok(tail_columns)
}

/// 一个工作表的内容指纹（标题+表头+全部明细），用于检测两个不同文件是否为完全相同的重复导出。
fn sheet_fingerprint(sheet: &SheetGrid, last_row: u32) -> String {
    let mut rows = Vec::with_capacity((last_row + 1) as usize);
    rows.push(sheet.row_texts(1).join("\u{1}"));
    rows.push(sheet.row_texts(2).join("\u{1}"));
    for row in 3..=last_row {
        rows.push(sheet.row_texts(row).join("\u{1}"));
    }
    rows.join("\u{2}")
}

fn read_typed_cell(
    cell: &RawCell,
    ty: ColumnType,
    field: &str,
    file: &str,
    sheet_name: &str,
    row: u32,
) -> Result<Value, ProcessError> {
    Ok(match ty {
        ColumnType::Date => cell_date_or_text(cell),
        ColumnType::DateTime => cell_datetime_or_text(cell),
        ColumnType::Decimal(_) => cell_amount(cell).map(amount_value).map_err(|detail| {
            data_error(file, sheet_name, row, field, cell_display(cell), detail)
        })?,
        _ => text_value(cell_text(cell).map_err(|detail| {
            data_error(file, sheet_name, row, field, cell_display(cell), detail)
        })?),
    })
}

fn read_row(
    sheet: &SheetGrid,
    row: u32,
    config: &UploadedConfig,
    tail_columns: &[Option<u32>],
    file: &str,
    sheet_name: &str,
) -> Result<Row, ProcessError> {
    let merchant_cell = sheet.cell(row, COL_MERCHANT_NO);
    let merchant_no = cell_text(&merchant_cell).map_err(|detail| {
        data_error(
            file,
            sheet_name,
            row,
            "商户号",
            cell_display(&merchant_cell),
            detail,
        )
    })?;
    if merchant_no != config.merchant_no {
        return Err(data_error(
            file,
            sheet_name,
            row,
            "商户号",
            merchant_no,
            format!("与文件名商户号“{}”不一致", config.merchant_no),
        ));
    }

    let mut values = Vec::with_capacity(FIELD_COUNT);
    for (index, &ty) in FRONT_COLUMN_TYPES.iter().enumerate() {
        let col = (index + 1) as u32;
        let cell = sheet.cell(row, col);
        values.push(read_typed_cell(
            &cell,
            ty,
            FRONT_HEADERS[index],
            file,
            sheet_name,
            row,
        )?);
    }
    for (index, field) in config.tail.iter().enumerate() {
        let value = match tail_columns[index] {
            None => Value::Empty,
            Some(col) => {
                let cell = sheet.cell(row, col);
                read_typed_cell(&cell, field.ty, field.name, file, sheet_name, row)?
            }
        };
        values.push(value);
    }

    Ok(Row { values, fill: None })
}

fn output_columns(config: &UploadedConfig) -> Vec<Column> {
    FRONT_HEADERS
        .iter()
        .zip(FRONT_COLUMN_TYPES)
        .map(|(&name, ty)| Column { name, ty })
        .chain(config.tail.iter().map(|field| Column {
            name: field.name,
            ty: field.ty,
        }))
        .collect()
}

fn run_uploaded(config: &UploadedConfig, input_dir: &Path) -> Result<Table, ProcessError> {
    let mut files: Vec<PathBuf> = list_xlsx_files(input_dir)?
        .into_iter()
        .filter(|path| {
            path.file_name()
                .and_then(|n| n.to_str())
                .is_some_and(|n| matches_filename(n, config.merchant_no))
        })
        .collect();
    files.sort();
    if files.is_empty() {
        return Err(ProcessError::NoInput {
            pattern: format!("MER_{}_yyyymmddhhmmss_yjhx.xlsx", config.merchant_no),
        });
    }

    let mut rows = Vec::new();
    let mut fingerprints: HashMap<String, String> = HashMap::new();
    let mut seen_uuid: HashSet<String> = HashSet::new();

    for path in &files {
        let file_name = path.file_name().unwrap().to_string_lossy().into_owned();
        let sheets = open_sheets(path)?;

        for sheet in &sheets {
            let sheet_name = sheet.name().to_string();

            let tail_columns =
                resolve_columns(sheet, config).map_err(|detail| ProcessError::Structure {
                    file: file_name.clone(),
                    sheet: sheet_name.clone(),
                    detail,
                })?;

            let last_row = sheet.last_value_row().unwrap_or(2);

            let fingerprint = sheet_fingerprint(sheet, last_row);
            check_duplicate_fingerprint(&mut fingerprints, fingerprint, &file_name)?;

            for row in 3..=last_row {
                let record_row =
                    read_row(sheet, row, config, &tail_columns, &file_name, &sheet_name)?;
                if let Value::Text(uuid) = &record_row.values[(COL_UUID - 1) as usize]
                    && !seen_uuid.insert(uuid.clone())
                {
                    return Err(ProcessError::Duplicate {
                        detail: format!("实时清分UUID重复：{uuid}"),
                    });
                }
                rows.push(record_row);
            }
        }
    }

    Ok(Table {
        columns: output_columns(config),
        rows,
    })
}

#[cfg(test)]
mod tests {
    use rust_xlsxwriter::Workbook;

    use super::*;
    use crate::test_support::unique_temp_path;

    fn write_workbook(path: &Path, title: &str, headers: &[&str], rows: &[Vec<&str>]) {
        let mut workbook = Workbook::new();
        let sheet = workbook.add_worksheet();
        sheet.write_string(0, 0, title).unwrap();
        for (col, name) in headers.iter().enumerate() {
            sheet.write_string(1, col as u16, *name).unwrap();
        }
        for (row_index, row) in rows.iter().enumerate() {
            for (col, value) in row.iter().enumerate() {
                sheet
                    .write_string((2 + row_index) as u32, col as u16, *value)
                    .unwrap();
            }
        }
        workbook.save(path).unwrap();
    }

    /// 按数据组标准表头（前25列固定 + 该组的尾部字段名）生成表头。
    fn standard_headers(tail: &[TailField]) -> Vec<&'static str> {
        FRONT_HEADERS
            .iter()
            .copied()
            .chain(tail.iter().map(|f| f.synonyms[0]))
            .collect()
    }

    /// 按`APPLIANCE_TAIL`顺序生成一行示例数据；两个`图片1`用不同值验证按位置区分。
    fn appliance_row<'a>(uuid: &'a str, merchant_no: &'a str, order_no: &'a str) -> Vec<&'a str> {
        vec![
            uuid,
            merchant_no,
            "商户甲",
            order_no,
            "20260107",
            "100.00",
            "REF001",
            "模版A",
            "已上传",
            "描述文本",
            "2026-09-14 10:18:09",
            "2026-09-14 10:20:00",
            "T001",
            "STORE001",
            "分店甲",
            "地区甲",
            "详细地址甲",
            "110000",
            "13800000000",
            "INV001",
            "3000,00",
            "购买方甲",
            "PIC_FIRST",
            "SN0001",
            "是",
            "PIC_SECOND",
            "PIC2",
            "PIC3",
            "PIC4",
            "IMG5",
            "IMG6",
            "IMG7",
            "IMG8",
            "IMG9",
            "IMG10",
            "IMG11",
            "IMG12",
            "IMG13",
            "IMG14",
            "IMG15",
            "2026-09-14 12:00:00",
            "备注甲",
            "EEG甲",
            "LOG001",
            "ERP001",
            "ocr甲",
            "状态甲",
            "标记甲",
            "是",
            "否",
            "收件人甲",
            "PC001",
            "",
            "商品甲",
            "品类甲",
            "否",
            "空调信息甲",
            "20260109",
        ]
    }

    /// 按`DIGITAL_TAIL`顺序生成一行示例数据。
    fn digital_row<'a>(uuid: &'a str, merchant_no: &'a str, order_no: &'a str) -> Vec<&'a str> {
        vec![
            uuid,
            merchant_no,
            "商户乙",
            order_no,
            "20260108",
            "200.00",
            "REF002",
            "模版B",
            "已上传",
            "描述文本2",
            "2026-09-15 11:00:00",
            "2026-09-15 11:05:00",
            "T002",
            "STORE002",
            "分店乙",
            "地区乙",
            "详细地址乙",
            "220000",
            "13900000000",
            "INV002",
            "6000,00",
            "购买方乙",
            "PIC_FIRST2",
            "SN0002",
            "否",
            "IMEI0001",
            "IMEI0002",
            "PIC_SECOND2",
            "PIC2",
            "PIC3",
            "PIC4",
            "PIC5",
            "PIC6",
            "IMG7",
            "IMG8",
            "IMG9",
            "IMG10",
            "IMG11",
            "IMG12",
            "IMG13",
            "IMG14",
            "IMG15",
            "2026-09-15 13:00:00",
            "备注乙",
            "LOG002",
            "ERP002",
            "ocr乙",
            "状态乙",
            "标记乙",
            "否",
            "是",
            "收件人乙",
            "PC002",
            "",
            "商品乙",
            "旧换类型乙",
            "是",
            "20260110",
        ]
    }

    #[test]
    fn matches_filename_checks_merchant_timestamp_and_suffix() {
        assert!(matches_filename(
            "MER_89813014812B06R_20260914101809_yjhx.xlsx",
            "89813014812B06R"
        ));
        assert!(!matches_filename(
            "MER_89813015722APT1_20260914101809_yjhx.xlsx",
            "89813014812B06R"
        ));
        assert!(!matches_filename(
            "MER_89813014812B06R_2026_yjhx.xlsx",
            "89813014812B06R"
        ));
        assert!(!matches_filename(
            "MER_89813014812B06R_20260914101809_other.xlsx",
            "89813014812B06R"
        ));
    }

    #[test]
    fn digital_job_reads_58_columns_and_parses_typed_fields() {
        let dir = unique_temp_path("uploaded-digital-happy-path");
        std::fs::create_dir_all(&dir).unwrap();
        write_workbook(
            &dir.join("MER_89813014812B06R_20260914101809_yjhx.xlsx"),
            "以旧换新数据[手机/其他3C]",
            &standard_headers(&DIGITAL_TAIL),
            &[digital_row("UUID-1", "89813014812B06R", "ORDER-1")],
        );

        let table = UPLOADED_DIGITAL.run(&dir).unwrap();

        assert_eq!(table.columns.len(), 58);
        assert_eq!(table.rows.len(), 1);
        let values = &table.rows[0].values;
        assert_eq!(values[0], Value::Text("UUID-1".to_string()));
        assert!(matches!(values[4], Value::Date(_))); // 交易日期
        assert!(matches!(values[10], Value::DateTime(_))); // 提交时间
        assert!(matches!(values[42], Value::DateTime(_))); // 签收时间
        assert!(matches!(values[57], Value::Date(_))); // 开票日期
        assert_eq!(values[20], Value::Text("6000,00".to_string())); // 发票金额：原样保留，不解析为数值
        assert_eq!(values[25], Value::Text("IMEI0001".to_string()));
        assert_eq!(values[26], Value::Text("IMEI0002".to_string()));
        // 两个“图片1”按位置区分，取值不同。
        assert_eq!(values[22], Value::Text("PIC_FIRST2".to_string()));
        assert_eq!(values[27], Value::Text("PIC_SECOND2".to_string()));

        std::fs::remove_dir_all(&dir).unwrap();
    }

    #[test]
    fn appliance_job_has_58_columns_and_ignores_digital_files() {
        let dir = unique_temp_path("uploaded-appliance-happy-path");
        std::fs::create_dir_all(&dir).unwrap();
        write_workbook(
            &dir.join("MER_89813015722APT1_20260914101809_yjhx.xlsx"),
            "家电",
            &standard_headers(&APPLIANCE_TAIL),
            &[appliance_row("UUID-2", "89813015722APT1", "ORDER-2")],
        );
        // 数码商户的文件必须被家电、电脑数据组忽略。
        write_workbook(
            &dir.join("MER_89813014812B06R_20260914101809_yjhx.xlsx"),
            "数码",
            &standard_headers(&DIGITAL_TAIL),
            &[digital_row("UUID-3", "89813014812B06R", "ORDER-3")],
        );

        let table = UPLOADED_APPLIANCE.run(&dir).unwrap();

        assert_eq!(table.columns.len(), 58);
        assert_eq!(table.rows.len(), 1);
        let values = &table.rows[0].values;
        assert_eq!(values[0], Value::Text("UUID-2".to_string()));
        assert_eq!(values[2], Value::Text("商户甲".to_string())); // 商户名称
        assert!(matches!(values[40], Value::DateTime(_))); // 签收时间（家电结构无 IMEI，位置在40）
        assert_eq!(values[20], Value::Text("3000,00".to_string())); // 发票金额
        assert_eq!(values[22], Value::Text("PIC_FIRST".to_string()));
        assert_eq!(values[25], Value::Text("PIC_SECOND".to_string()));

        std::fs::remove_dir_all(&dir).unwrap();
    }

    /// 已知的“电脑”工作表变体：第26列多出`电脑类型`（不输出）、第5张图片改名为`图片5`
    /// （而非`img5`），且没有`airConditionerKitInfo`。必须仍能按名称正确映射到统一输出
    /// 结构，缺失字段留空而不是终止处理。
    #[test]
    fn maps_computer_variant_sheet_by_name_and_leaves_missing_field_blank() {
        let dir = unique_temp_path("uploaded-computer-variant");
        std::fs::create_dir_all(&dir).unwrap();

        let mut headers = standard_headers(&APPLIANCE_TAIL);
        headers.insert(FRONT_LEN, "电脑类型"); // 不属于统一表头的额外字段
        let img5_pos = headers.iter().position(|h| *h == "img5").unwrap();
        headers[img5_pos] = "图片5"; // “电脑”表使用的同义名称
        let ac_pos = headers
            .iter()
            .position(|h| *h == "airConditionerKitInfo")
            .unwrap();
        headers.remove(ac_pos); // “电脑”表没有此字段

        // 行数据与表头做相同的插入/删除，值随其原字段一起移动，位置始终对应正确的表头。
        let mut row = appliance_row("UUID-PC", "89813015722APT1", "ORDER-PC");
        row.insert(FRONT_LEN, "电脑类型值");
        row.remove(ac_pos);

        write_workbook(
            &dir.join("MER_89813015722APT1_20260914102100_yjhx.xlsx"),
            "电脑",
            &headers,
            &[row],
        );

        let table = UPLOADED_APPLIANCE.run(&dir).unwrap();
        assert_eq!(table.rows.len(), 1);
        let values = &table.rows[0].values;
        assert_eq!(values[0], Value::Text("UUID-PC".to_string()));
        assert_eq!(values[29], Value::Text("IMG5".to_string())); // img5 通过同义名称“图片5”映射
        assert_eq!(values[56], Value::Empty); // airConditionerKitInfo 缺失，留空
        assert!(matches!(values[57], Value::Date(_))); // 开票日期仍正确映射

        std::fs::remove_dir_all(&dir).unwrap();
    }

    #[test]
    fn rejects_merchant_mismatch_between_filename_and_row() {
        let dir = unique_temp_path("uploaded-merchant-mismatch");
        std::fs::create_dir_all(&dir).unwrap();
        write_workbook(
            &dir.join("MER_89813015722APT1_20260914101809_yjhx.xlsx"),
            "家电",
            &standard_headers(&APPLIANCE_TAIL),
            &[appliance_row("UUID-5", "89813014812B06R", "ORDER-5")],
        );

        let error = UPLOADED_APPLIANCE.run(&dir).unwrap_err();
        assert!(matches!(error, ProcessError::Data { .. }));

        std::fs::remove_dir_all(&dir).unwrap();
    }

    #[test]
    fn rejects_front_header_that_does_not_match_fixed_structure() {
        let dir = unique_temp_path("uploaded-header-mismatch");
        std::fs::create_dir_all(&dir).unwrap();
        let mut headers = standard_headers(&APPLIANCE_TAIL);
        headers[13] = "门店编号"; // 与规定的“分店id”不一致
        write_workbook(
            &dir.join("MER_89813015722APT1_20260914101809_yjhx.xlsx"),
            "家电",
            &headers,
            &[appliance_row("UUID-6", "89813015722APT1", "ORDER-6")],
        );

        let error = UPLOADED_APPLIANCE.run(&dir).unwrap_err();
        assert!(matches!(error, ProcessError::Structure { .. }));

        std::fs::remove_dir_all(&dir).unwrap();
    }

    #[test]
    fn rejects_duplicate_uuid_across_files() {
        let dir = unique_temp_path("uploaded-duplicate-uuid");
        std::fs::create_dir_all(&dir).unwrap();
        write_workbook(
            &dir.join("MER_89813015722APT1_20260914101809_yjhx.xlsx"),
            "家电",
            &standard_headers(&APPLIANCE_TAIL),
            &[appliance_row("UUID-SAME", "89813015722APT1", "ORDER-8")],
        );
        write_workbook(
            &dir.join("MER_89813015722APT1_20260914101900_yjhx.xlsx"),
            "家电",
            &standard_headers(&APPLIANCE_TAIL),
            &[appliance_row("UUID-SAME", "89813015722APT1", "ORDER-9")],
        );

        let error = UPLOADED_APPLIANCE.run(&dir).unwrap_err();
        assert!(matches!(error, ProcessError::Duplicate { .. }));

        std::fs::remove_dir_all(&dir).unwrap();
    }

    #[test]
    fn rejects_whole_sheet_duplicate_export_across_files() {
        let dir = unique_temp_path("uploaded-duplicate-sheet");
        std::fs::create_dir_all(&dir).unwrap();
        let rows = [appliance_row("UUID-DUP", "89813015722APT1", "ORDER-10")];
        write_workbook(
            &dir.join("MER_89813015722APT1_20260914101809_yjhx.xlsx"),
            "家电",
            &standard_headers(&APPLIANCE_TAIL),
            &rows,
        );
        write_workbook(
            &dir.join("MER_89813015722APT1_20260914101900_yjhx.xlsx"),
            "家电",
            &standard_headers(&APPLIANCE_TAIL),
            &rows,
        );

        let error = UPLOADED_APPLIANCE.run(&dir).unwrap_err();
        assert!(matches!(error, ProcessError::Duplicate { .. }));

        std::fs::remove_dir_all(&dir).unwrap();
    }

    #[test]
    fn unparseable_date_falls_back_to_original_text() {
        let dir = unique_temp_path("uploaded-bad-date");
        std::fs::create_dir_all(&dir).unwrap();
        let mut row = appliance_row("UUID-11", "89813015722APT1", "ORDER-11");
        row[4] = "不是日期";

        write_workbook(
            &dir.join("MER_89813015722APT1_20260914101809_yjhx.xlsx"),
            "家电",
            &standard_headers(&APPLIANCE_TAIL),
            &[row],
        );

        let table = UPLOADED_APPLIANCE.run(&dir).unwrap();
        assert_eq!(table.rows[0].values[4], Value::Text("不是日期".to_string()));

        std::fs::remove_dir_all(&dir).unwrap();
    }

    #[test]
    fn reports_no_input_when_nothing_matches() {
        let dir = unique_temp_path("uploaded-no-input");
        std::fs::create_dir_all(&dir).unwrap();

        let error = UPLOADED_APPLIANCE.run(&dir).unwrap_err();
        assert!(matches!(error, ProcessError::NoInput { .. }));

        std::fs::remove_dir_all(&dir).unwrap();
    }
}
