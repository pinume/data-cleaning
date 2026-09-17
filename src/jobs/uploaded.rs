use std::collections::{HashMap, HashSet};
use std::path::{Path, PathBuf};

use crate::io::paths::list_xlsx_files;
use crate::io::xlsx_reader::{SheetGrid, open_sheets};
use crate::model::{Column, ColumnType, DecimalScale, ProcessError, Row, Table, Value};

use super::{
    Category, Job, amount_value, cell_amount, cell_date_or_text, cell_datetime_or_text,
    cell_display, cell_text, text_value,
};

/// 14 个公共输出字段；数码数据组在此基础上追加 IMEI1/IMEI2 共 16 列。
/// 不含“发票金额”：源数据格式不统一（如“3000,00”），已确认不纳入输出。
const COMMON_FIELDS: [&str; 14] = [
    "实时清分UUID",
    "商户号",
    "订单号",
    "交易日期",
    "交易金额",
    "检索参考号",
    "状态",
    "描述",
    "提交时间",
    "更新时间",
    "终端号",
    "发票号码",
    "购买方名称",
    "S/N码",
];
const IMEI_FIELDS: [&str; 2] = ["IMEI1", "IMEI2"];

struct UploadedConfig {
    category: Category,
    title: &'static str,
    output_stem: &'static str,
    merchant_no: &'static str,
    include_imei: bool,
}

const DIGITAL_CONFIG: UploadedConfig = UploadedConfig {
    category: Category::UploadedDigital,
    title: "数码已上传数据",
    output_stem: "已上传数码",
    merchant_no: "89813014812B06R",
    include_imei: true,
};

const APPLIANCE_CONFIG: UploadedConfig = UploadedConfig {
    category: Category::UploadedAppliance,
    title: "家电、电脑已上传数据",
    output_stem: "已上传家电电脑",
    merchant_no: "89813015722APT1",
    include_imei: false,
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

fn data_error(
    file: &str,
    sheet: &str,
    row: u32,
    field: &str,
    value: String,
    detail: String,
) -> ProcessError {
    ProcessError::Data {
        file: file.to_string(),
        sheet: sheet.to_string(),
        row,
        field: field.to_string(),
        value,
        detail,
    }
}

/// 按第 2 行表头名称把每个必需字段唯一映射到列号；缺失或重名均视为结构异常。
fn resolve_columns(
    sheet: &SheetGrid,
    fields: &[&'static str],
) -> Result<HashMap<&'static str, u32>, String> {
    let header = sheet.row_texts(2);
    let mut positions: HashMap<&str, Vec<u32>> = HashMap::new();
    for (index, name) in header.iter().enumerate() {
        positions
            .entry(name.as_str())
            .or_default()
            .push(index as u32 + 1);
    }

    let mut resolved = HashMap::new();
    for &field in fields {
        match positions.get(field).map(Vec::as_slice) {
            None | Some([]) => return Err(format!("缺少必需字段：{field}")),
            Some([column]) => {
                resolved.insert(field, *column);
            }
            Some(columns) => {
                return Err(format!(
                    "字段“{field}”无法唯一识别（第2行出现{}次）",
                    columns.len()
                ));
            }
        }
    }
    Ok(resolved)
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

fn read_row(
    sheet: &SheetGrid,
    row: u32,
    columns: &HashMap<&'static str, u32>,
    config: &UploadedConfig,
    file: &str,
    sheet_name: &str,
) -> Result<Row, ProcessError> {
    let text_at = |field: &'static str| -> Result<String, ProcessError> {
        let cell = sheet.cell(row, columns[field]);
        cell_text(&cell)
            .map_err(|detail| data_error(file, sheet_name, row, field, cell_display(&cell), detail))
    };
    let amount_at = |field: &'static str| -> Result<Value, ProcessError> {
        let cell = sheet.cell(row, columns[field]);
        cell_amount(&cell)
            .map(amount_value)
            .map_err(|detail| data_error(file, sheet_name, row, field, cell_display(&cell), detail))
    };

    let merchant_no = text_at("商户号")?;
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

    let mut values = vec![
        text_value(text_at("实时清分UUID")?),
        text_value(merchant_no),
        text_value(text_at("订单号")?),
        cell_date_or_text(&sheet.cell(row, columns["交易日期"])),
        amount_at("交易金额")?,
        text_value(text_at("检索参考号")?),
        text_value(text_at("状态")?),
        text_value(text_at("描述")?),
        cell_datetime_or_text(&sheet.cell(row, columns["提交时间"])),
        cell_datetime_or_text(&sheet.cell(row, columns["更新时间"])),
        text_value(text_at("终端号")?),
        text_value(text_at("发票号码")?),
        text_value(text_at("购买方名称")?),
        text_value(text_at("S/N码")?),
    ];

    if config.include_imei {
        values.push(text_value(text_at("IMEI1")?));
        values.push(text_value(text_at("IMEI2")?));
    }

    Ok(Row { values, fill: None })
}

fn output_columns(include_imei: bool) -> Vec<Column> {
    let mut columns = vec![
        Column {
            name: "实时清分UUID",
            ty: ColumnType::Text,
        },
        Column {
            name: "商户号",
            ty: ColumnType::Text,
        },
        Column {
            name: "订单号",
            ty: ColumnType::Text,
        },
        Column {
            name: "交易日期",
            ty: ColumnType::Date,
        },
        Column {
            name: "交易金额",
            ty: ColumnType::Decimal(DecimalScale::Original),
        },
        Column {
            name: "检索参考号",
            ty: ColumnType::Text,
        },
        Column {
            name: "状态",
            ty: ColumnType::Text,
        },
        Column {
            name: "描述",
            ty: ColumnType::Text,
        },
        Column {
            name: "提交时间",
            ty: ColumnType::DateTime,
        },
        Column {
            name: "更新时间",
            ty: ColumnType::DateTime,
        },
        Column {
            name: "终端号",
            ty: ColumnType::Text,
        },
        Column {
            name: "发票号码",
            ty: ColumnType::Text,
        },
        Column {
            name: "购买方名称",
            ty: ColumnType::Text,
        },
        Column {
            name: "S/N码",
            ty: ColumnType::Text,
        },
    ];
    if include_imei {
        columns.push(Column {
            name: "IMEI1",
            ty: ColumnType::Text,
        });
        columns.push(Column {
            name: "IMEI2",
            ty: ColumnType::Text,
        });
    }
    columns
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

    let fields: Vec<&'static str> = if config.include_imei {
        COMMON_FIELDS
            .iter()
            .chain(IMEI_FIELDS.iter())
            .copied()
            .collect()
    } else {
        COMMON_FIELDS.to_vec()
    };

    let mut rows = Vec::new();
    let mut fingerprints: HashMap<String, String> = HashMap::new();
    let mut seen_uuid: HashSet<String> = HashSet::new();

    for path in &files {
        let file_name = path.file_name().unwrap().to_string_lossy().into_owned();
        let sheets = open_sheets(path)?;

        for sheet in &sheets {
            let sheet_name = sheet.name().to_string();

            let columns =
                resolve_columns(sheet, &fields).map_err(|detail| ProcessError::Structure {
                    file: file_name.clone(),
                    sheet: sheet_name.clone(),
                    detail,
                })?;

            let last_row = sheet.last_value_row().unwrap_or(2);

            let fingerprint = sheet_fingerprint(sheet, last_row);
            if let Some(previous_file) = fingerprints.get(&fingerprint) {
                if previous_file != &file_name {
                    return Err(ProcessError::Duplicate {
                        detail: format!(
                            "{file_name} 与 {previous_file} 的工作表内容完全相同，疑似重复导出"
                        ),
                    });
                }
            } else {
                fingerprints.insert(fingerprint, file_name.clone());
            }

            for row in 3..=last_row {
                let record_row = read_row(sheet, row, &columns, config, &file_name, &sheet_name)?;
                if let Value::Text(uuid) = &record_row.values[0]
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
        columns: output_columns(config.include_imei),
        rows,
    })
}

#[cfg(test)]
mod tests {
    use rust_xlsxwriter::Workbook;

    use super::*;
    use crate::test_support::unique_temp_path;

    fn write_workbook(path: &Path, title: &str, fields: &[&str], rows: &[Vec<&str>]) {
        let mut workbook = Workbook::new();
        let sheet = workbook.add_worksheet();
        sheet.write_string(0, 0, title).unwrap();
        for (col, name) in fields.iter().enumerate() {
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

    /// 按`COMMON_FIELDS`（+ 数码追加的 IMEI）顺序生成一行示例数据。
    fn sample_row<'a>(
        uuid: &'a str,
        merchant_no: &'a str,
        order_no: &'a str,
        with_imei: bool,
    ) -> Vec<&'a str> {
        let mut row = vec![
            uuid,
            merchant_no,
            order_no,
            "20260914",
            "100.00",
            "REF001",
            "已上传",
            "描述文本",
            "2026-09-14 10:18:09",
            "2026-09-14 10:20:00",
            "T001",
            "INV001",
            "购买方甲",
            "SN0001",
        ];
        if with_imei {
            row.push("IMEI0001");
            row.push("IMEI0002");
        }
        row
    }

    fn all_fields(with_imei: bool) -> Vec<&'static str> {
        let mut fields = COMMON_FIELDS.to_vec();
        if with_imei {
            fields.extend_from_slice(&IMEI_FIELDS);
        }
        fields
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
    fn digital_job_reads_16_columns_and_parses_date_time() {
        let dir = unique_temp_path("uploaded-digital-happy-path");
        std::fs::create_dir_all(&dir).unwrap();
        write_workbook(
            &dir.join("MER_89813014812B06R_20260914101809_yjhx.xlsx"),
            "以旧换新数据[手机/其他3C]",
            &all_fields(true),
            &[sample_row("UUID-1", "89813014812B06R", "ORDER-1", true)],
        );

        let table = UPLOADED_DIGITAL.run(&dir).unwrap();

        assert_eq!(table.columns.len(), 16);
        assert_eq!(table.rows.len(), 1);
        assert_eq!(table.rows[0].values[0], Value::Text("UUID-1".to_string()));
        assert!(matches!(table.rows[0].values[3], Value::Date(_)));
        assert!(matches!(table.rows[0].values[8], Value::DateTime(_)));
        assert_eq!(
            table.rows[0].values[14],
            Value::Text("IMEI0001".to_string())
        );
        assert_eq!(
            table.rows[0].values[15],
            Value::Text("IMEI0002".to_string())
        );

        std::fs::remove_dir_all(&dir).unwrap();
    }

    #[test]
    fn appliance_job_has_14_columns_and_ignores_digital_files() {
        let dir = unique_temp_path("uploaded-appliance-happy-path");
        std::fs::create_dir_all(&dir).unwrap();
        write_workbook(
            &dir.join("MER_89813015722APT1_20260914101809_yjhx.xlsx"),
            "家电",
            &all_fields(false),
            &[sample_row("UUID-2", "89813015722APT1", "ORDER-2", false)],
        );
        // 数码商户的文件必须被家电、电脑数据组忽略。
        write_workbook(
            &dir.join("MER_89813014812B06R_20260914101809_yjhx.xlsx"),
            "数码",
            &all_fields(true),
            &[sample_row("UUID-3", "89813014812B06R", "ORDER-3", true)],
        );

        let table = UPLOADED_APPLIANCE.run(&dir).unwrap();

        assert_eq!(table.columns.len(), 14);
        assert_eq!(table.rows.len(), 1);
        assert_eq!(table.rows[0].values[0], Value::Text("UUID-2".to_string()));

        std::fs::remove_dir_all(&dir).unwrap();
    }

    #[test]
    fn resolves_fields_by_name_regardless_of_column_order() {
        let dir = unique_temp_path("uploaded-shuffled-header");
        std::fs::create_dir_all(&dir).unwrap();

        // 表头顺序打乱（S/N码 与 实时清分UUID 互换位置），仍须按字段名正确取值。
        let mut fields = all_fields(false);
        fields.swap(0, 13);
        let mut row = sample_row("UUID-4", "89813015722APT1", "ORDER-4", false);
        row.swap(0, 13);

        write_workbook(
            &dir.join("MER_89813015722APT1_20260914101809_yjhx.xlsx"),
            "家电",
            &fields,
            &[row],
        );

        let table = UPLOADED_APPLIANCE.run(&dir).unwrap();
        assert_eq!(table.rows[0].values[0], Value::Text("UUID-4".to_string()));
        assert_eq!(table.rows[0].values[13], Value::Text("SN0001".to_string()));

        std::fs::remove_dir_all(&dir).unwrap();
    }

    #[test]
    fn rejects_merchant_mismatch_between_filename_and_row() {
        let dir = unique_temp_path("uploaded-merchant-mismatch");
        std::fs::create_dir_all(&dir).unwrap();
        write_workbook(
            &dir.join("MER_89813015722APT1_20260914101809_yjhx.xlsx"),
            "家电",
            &all_fields(false),
            &[sample_row("UUID-5", "89813014812B06R", "ORDER-5", false)],
        );

        let error = UPLOADED_APPLIANCE.run(&dir).unwrap_err();
        assert!(matches!(error, ProcessError::Data { .. }));

        std::fs::remove_dir_all(&dir).unwrap();
    }

    #[test]
    fn rejects_missing_required_field() {
        let dir = unique_temp_path("uploaded-missing-field");
        std::fs::create_dir_all(&dir).unwrap();
        let mut fields = all_fields(false);
        fields.remove(13); // 去掉 S/N码
        let mut row = sample_row("UUID-6", "89813015722APT1", "ORDER-6", false);
        row.remove(13);

        write_workbook(
            &dir.join("MER_89813015722APT1_20260914101809_yjhx.xlsx"),
            "家电",
            &fields,
            &[row],
        );

        let error = UPLOADED_APPLIANCE.run(&dir).unwrap_err();
        assert!(matches!(error, ProcessError::Structure { .. }));

        std::fs::remove_dir_all(&dir).unwrap();
    }

    #[test]
    fn rejects_ambiguous_duplicate_field_name() {
        let dir = unique_temp_path("uploaded-duplicate-field-name");
        std::fs::create_dir_all(&dir).unwrap();
        let mut fields = all_fields(false);
        fields[0] = "订单号"; // 与已有字段重名，制造无法唯一识别
        write_workbook(
            &dir.join("MER_89813015722APT1_20260914101809_yjhx.xlsx"),
            "家电",
            &fields,
            &[sample_row("UUID-7", "89813015722APT1", "ORDER-7", false)],
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
            &all_fields(false),
            &[sample_row("UUID-SAME", "89813015722APT1", "ORDER-8", false)],
        );
        write_workbook(
            &dir.join("MER_89813015722APT1_20260914101900_yjhx.xlsx"),
            "家电",
            &all_fields(false),
            &[sample_row("UUID-SAME", "89813015722APT1", "ORDER-9", false)],
        );

        let error = UPLOADED_APPLIANCE.run(&dir).unwrap_err();
        assert!(matches!(error, ProcessError::Duplicate { .. }));

        std::fs::remove_dir_all(&dir).unwrap();
    }

    #[test]
    fn rejects_whole_sheet_duplicate_export_across_files() {
        let dir = unique_temp_path("uploaded-duplicate-sheet");
        std::fs::create_dir_all(&dir).unwrap();
        let rows = [sample_row("UUID-DUP", "89813015722APT1", "ORDER-10", false)];
        write_workbook(
            &dir.join("MER_89813015722APT1_20260914101809_yjhx.xlsx"),
            "家电",
            &all_fields(false),
            &rows,
        );
        write_workbook(
            &dir.join("MER_89813015722APT1_20260914101900_yjhx.xlsx"),
            "家电",
            &all_fields(false),
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
        let mut row = sample_row("UUID-11", "89813015722APT1", "ORDER-11", false);
        row[3] = "不是日期";

        write_workbook(
            &dir.join("MER_89813015722APT1_20260914101809_yjhx.xlsx"),
            "家电",
            &all_fields(false),
            &[row],
        );

        let table = UPLOADED_APPLIANCE.run(&dir).unwrap();
        assert_eq!(table.rows[0].values[3], Value::Text("不是日期".to_string()));

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
