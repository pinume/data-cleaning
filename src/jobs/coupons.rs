use std::collections::HashSet;
use std::path::Path;
use std::sync::LazyLock;

use regex::Regex;
use rust_decimal::Decimal;

use crate::io::xlsx_reader::{RawCell, SheetGrid, open_sheets};
use crate::model::{Column, ColumnType, ProcessError, Row, Table, Value};
use crate::utils::{numbers, text};

use super::unionpay;
use super::{
    Category, Job, build_match_doc_no, cell_display, cell_text, data_error, parse_date_field,
    text_value,
};

const FILE_NAME: &str = "销售用券情况统计.xlsx";
const TITLE: &str = "销售用券情况统计";

const SOURCE_HEADERS: [&str; 28] = [
    "供应商",
    "原始供应商",
    "单据号",
    "单据日期",
    "收款员",
    "商品名称",
    "商品简码",
    "品牌",
    "销售部门",
    "销售员",
    "业绩成本",
    "业绩利润",
    "价格类型名称",
    "库存类型",
    "财务大类",
    "顾客姓名",
    "备注",
    "明细摘要",
    "收款日期",
    "销售成本",
    "不含券收入",
    "其它",
    "本期尾款",
    "销售单价",
    "含券收入",
    "2026家电国补（计入收入）",
    "2026数码国补（计入收入）",
    "合计",
];

// 源表列位置（1 基）。
const COL_DOC_NO: u32 = 3;
const COL_DOC_DATE: u32 = 4;
const COL_PRODUCT_NAME: u32 = 6;
const COL_BRAND: u32 = 8;
const COL_FINANCE_CATEGORY: u32 = 15;
const COL_SUMMARY: u32 = 18;
const COL_TOTAL: u32 = 28;

const TRIGGER_LETTERS: [char; 10] = ['N', 'n', 'W', 'w', 'M', 'm', 'H', 'h', 'B', 'b'];

struct CouponRecord {
    doc_no: String,
    doc_date: Value,
    product_name: String,
    brand: String,
    finance_category: String,
    subsidy: Decimal,
    summary: String,
}

pub struct CouponsJob;

impl Job for CouponsJob {
    fn category(&self) -> Category {
        Category::Coupons
    }

    fn title(&self) -> &'static str {
        "销售用券情况统计"
    }

    fn output_stem(&self) -> &'static str {
        "销售用券情况统计"
    }

    fn run(&self, input_dir: &Path) -> Result<Table, ProcessError> {
        // 10.6.1：权威校验集缺失、结构异常或无法完整读取时必须停止，不得绕过校验；
        // 直接复用 unionpay::load_records 的全部校验（文件发现、表头、首尾结构、重复导出）。
        let authority = build_authority(input_dir)?;

        let path = input_dir.join(FILE_NAME);
        if !path.is_file() {
            return Err(ProcessError::NoInput {
                pattern: FILE_NAME.to_string(),
            });
        }

        let sheets = open_sheets(&path)?;
        if sheets.len() != 1 {
            return Err(ProcessError::Structure {
                file: FILE_NAME.to_string(),
                sheet: String::new(),
                detail: format!("工作表数量异常：应为 1 个，实际为 {} 个", sheets.len()),
            });
        }
        let sheet = &sheets[0];
        let sheet_name = sheet.name().to_string();

        let title = cell_display(&sheet.cell(1, 1));
        if title != TITLE {
            return Err(ProcessError::Structure {
                file: FILE_NAME.to_string(),
                sheet: sheet_name,
                detail: format!("第1行第1列应为“{TITLE}”，实际为“{title}”"),
            });
        }

        let header = sheet.row_texts(2);
        if header.iter().map(String::as_str).collect::<Vec<_>>() != SOURCE_HEADERS {
            return Err(ProcessError::Structure {
                file: FILE_NAME.to_string(),
                sheet: sheet_name,
                detail: "第2行表头与规定的28个字段不一致".to_string(),
            });
        }

        let last_row = sheet.last_value_row().unwrap_or(2);
        let total_marker = cell_display(&sheet.cell(last_row, 1));
        if total_marker != "合计" {
            return Err(ProcessError::Structure {
                file: FILE_NAME.to_string(),
                sheet: sheet_name,
                detail: format!("最后一个实际有值行第1列应为“合计”，实际为“{total_marker}”"),
            });
        }

        let mut rows = Vec::new();
        for row in 3..last_row {
            let record = read_row(sheet, row, FILE_NAME, &sheet_name)?;
            rows.push(to_row(record, &authority));
        }

        Ok(Table {
            columns: output_columns(),
            rows,
        })
    }
}

fn build_authority(input_dir: &Path) -> Result<HashSet<String>, ProcessError> {
    let records = unionpay::load_records(input_dir)?;
    Ok(records
        .iter()
        .map(|r| r.retrieval_no.as_str())
        .filter(|v| is_valid_ref_no(v))
        .map(str::to_owned)
        .collect())
}

fn is_valid_ref_no(value: &str) -> bool {
    value.len() == 12
        && value.as_bytes()[11] == b'N'
        && value.as_bytes()[..11].iter().all(u8::is_ascii_digit)
}

/// 补贴额：取源字段`合计`，按 10.7 节尾差规则统一为两位小数；为空或超出容差时终止。
fn parse_subsidy(
    cell: &RawCell,
    file: &str,
    sheet: &str,
    row: u32,
) -> Result<Decimal, ProcessError> {
    let raw: Option<Decimal> = match cell {
        RawCell::Text(t) if t.trim().is_empty() => None,
        RawCell::Int(n) => Some(Decimal::from(*n)),
        RawCell::Float(f) => numbers::from_f64(*f),
        RawCell::Text(t) => t.trim().parse().ok(),
        _ => None,
    };
    let Some(raw) = raw else {
        return Err(data_error(
            file,
            sheet,
            row,
            "合计",
            cell_display(cell),
            "补贴额为空或无法解析为数值".to_string(),
        ));
    };
    numbers::to_cents(raw).ok_or_else(|| {
        data_error(
            file,
            sheet,
            row,
            "合计",
            cell_display(cell),
            "金额与最接近的两位小数之差超出0.000001元容差".to_string(),
        )
    })
}

fn read_row(
    sheet: &SheetGrid,
    row: u32,
    file: &str,
    sheet_name: &str,
) -> Result<CouponRecord, ProcessError> {
    let text_at = |col: u32, field: &'static str| -> Result<String, ProcessError> {
        let cell = sheet.cell(row, col);
        cell_text(&cell)
            .map_err(|detail| data_error(file, sheet_name, row, field, cell_display(&cell), detail))
    };

    Ok(CouponRecord {
        doc_no: text_at(COL_DOC_NO, "单据号")?,
        doc_date: parse_date_field(
            &sheet.cell(row, COL_DOC_DATE),
            "单据日期",
            file,
            sheet_name,
            row,
        )?,
        product_name: text_at(COL_PRODUCT_NAME, "商品名称")?,
        brand: text_at(COL_BRAND, "品牌")?,
        finance_category: text_at(COL_FINANCE_CATEGORY, "财务大类")?,
        subsidy: parse_subsidy(&sheet.cell(row, COL_TOTAL), file, sheet_name, row)?,
        summary: text_at(COL_SUMMARY, "明细摘要")?,
    })
}

fn quantity_of(subsidy: Decimal) -> i64 {
    use std::cmp::Ordering;
    match subsidy.cmp(&Decimal::ZERO) {
        Ordering::Greater => 1,
        Ordering::Equal => 0,
        Ordering::Less => -1,
    }
}

fn to_row(record: CouponRecord, authority: &HashSet<String>) -> Row {
    let quantity = quantity_of(record.subsidy);
    let match_doc_no = build_match_doc_no(&record.doc_date, &record.doc_no);
    let reference = extract_reference(&record.summary, authority);

    let values = vec![
        text_value(record.doc_no),
        record.doc_date,
        text_value(record.product_name),
        text_value(record.brand),
        text_value(record.finance_category),
        Value::Decimal(record.subsidy),
        Value::Integer(quantity),
        reference.map_or(Value::Empty, Value::Text),
        text_value(match_doc_no),
    ];
    Row { values, fill: None }
}

fn output_columns() -> Vec<Column> {
    vec![
        Column {
            name: "单据号",
            ty: ColumnType::Text,
        },
        Column {
            name: "单据日期",
            ty: ColumnType::Date,
        },
        Column {
            name: "商品名称",
            ty: ColumnType::Text,
        },
        Column {
            name: "品牌",
            ty: ColumnType::Text,
        },
        Column {
            name: "财务大类",
            ty: ColumnType::Text,
        },
        Column {
            name: "补贴额",
            ty: ColumnType::Decimal(crate::model::DecimalScale::Two),
        },
        Column {
            name: "数量",
            ty: ColumnType::Integer,
        },
        Column {
            name: "参考号",
            ty: ColumnType::Text,
        },
        Column {
            name: "匹配单据号",
            ty: ColumnType::Text,
        },
    ]
}

// ---------------------------------------------------------------------------
// 10.6 节：参考号提取（四级优先级，命中即停，同级内先去重再判定）。
// ---------------------------------------------------------------------------

enum PriorityOutcome {
    Unique(String),
    Ambiguous,
    NoHit,
}

fn resolve(hits: Vec<String>) -> PriorityOutcome {
    let distinct: HashSet<String> = hits.into_iter().collect();
    match distinct.len() {
        0 => PriorityOutcome::NoHit,
        1 => PriorityOutcome::Unique(distinct.into_iter().next().unwrap()),
        _ => PriorityOutcome::Ambiguous,
    }
}

fn extract_reference(summary: &str, authority: &HashSet<String>) -> Option<String> {
    if summary.trim().is_empty() {
        return None;
    }

    match resolve(as_is_hits(summary, authority)) {
        PriorityOutcome::Unique(value) => return Some(value),
        PriorityOutcome::Ambiguous => return None,
        PriorityOutcome::NoHit => {}
    }

    let digit_candidates = digit_candidates(summary);

    match resolve(zero_edit_hits(&digit_candidates, authority)) {
        PriorityOutcome::Unique(value) => return Some(value),
        PriorityOutcome::Ambiguous => return None,
        PriorityOutcome::NoHit => {}
    }

    match resolve(single_edit_hits(&digit_candidates, authority)) {
        PriorityOutcome::Unique(value) => Some(value),
        PriorityOutcome::Ambiguous | PriorityOutcome::NoHit => None,
    }
}

fn reference_pattern() -> &'static Regex {
    static RE: LazyLock<Regex> = LazyLock::new(|| Regex::new(r"[0-9]{11}N").unwrap());
    &RE
}

/// 第 1 级：原样提取完整`[0-9]{11}N`片段，前后不得紧邻数字或字母。
fn as_is_hits(summary: &str, authority: &HashSet<String>) -> Vec<String> {
    reference_pattern()
        .find_iter(summary)
        .filter(|m| text::has_isolated_boundaries(summary, m.start(), m.end()))
        .map(|m| m.as_str().to_string())
        .filter(|candidate| authority.contains(candidate))
        .collect()
}

/// 字节级扫描，返回所有连续 ASCII 数字片段的字节范围（Chinese 字符不会被误判）。
fn maximal_digit_runs(text: &str) -> Vec<(usize, usize)> {
    let bytes = text.as_bytes();
    let mut runs = Vec::new();
    let mut start = None;
    for (i, &b) in bytes.iter().enumerate() {
        if b.is_ascii_digit() {
            start.get_or_insert(i);
        } else if let Some(s) = start.take() {
            runs.push((s, i));
        }
    }
    if let Some(s) = start {
        runs.push((s, bytes.len()));
    }
    runs
}

/// 第 2 级候选：长度 10-12 的完整连续数字串；若摘要含触发字母，额外把全部数字顺序拼接为一个候选。
fn digit_candidates(summary: &str) -> Vec<String> {
    let mut candidates = Vec::new();

    for (start, end) in maximal_digit_runs(summary) {
        if (10..=12).contains(&(end - start)) {
            candidates.push(summary[start..end].to_string());
        }
    }

    if summary.chars().any(|c| TRIGGER_LETTERS.contains(&c)) {
        let all_digits: String = summary.chars().filter(char::is_ascii_digit).collect();
        if (10..=12).contains(&all_digits.len()) {
            candidates.push(all_digits);
        }
    }

    candidates
}

/// 第 3 级：仅对 11 位候选补上大写`N`后与权威校验集比对，不做任何数字改动。
fn zero_edit_hits(candidates: &[String], authority: &HashSet<String>) -> Vec<String> {
    candidates
        .iter()
        .filter(|c| c.len() == 11)
        .map(|c| format!("{c}N"))
        .filter(|candidate| authority.contains(candidate))
        .collect()
}

/// 第 4 级：对 10-12 位候选执行恰好一次插入/替换/删除，纠正为 11 位后补`N`比对。
fn single_edit_hits(candidates: &[String], authority: &HashSet<String>) -> Vec<String> {
    let mut hits = Vec::new();
    for candidate in candidates {
        for variant in single_edit_variants(candidate) {
            let with_suffix = format!("{variant}N");
            if authority.contains(&with_suffix) {
                hits.push(with_suffix);
            }
        }
    }
    hits
}

fn single_edit_variants(candidate: &str) -> Vec<String> {
    let digits = candidate.as_bytes();
    let mut variants = Vec::new();
    match digits.len() {
        10 => {
            for pos in 0..=digits.len() {
                for d in b'0'..=b'9' {
                    let mut v = digits.to_vec();
                    v.insert(pos, d);
                    variants.push(String::from_utf8(v).unwrap());
                }
            }
        }
        11 => {
            for pos in 0..digits.len() {
                for d in b'0'..=b'9' {
                    if d != digits[pos] {
                        let mut v = digits.to_vec();
                        v[pos] = d;
                        variants.push(String::from_utf8(v).unwrap());
                    }
                }
            }
        }
        12 => {
            for pos in 0..digits.len() {
                let mut v = digits.to_vec();
                v.remove(pos);
                variants.push(String::from_utf8(v).unwrap());
            }
        }
        _ => {}
    }
    variants
}

#[cfg(test)]
mod tests {
    use rust_xlsxwriter::Workbook;

    use super::*;
    use crate::jobs::unionpay;
    use crate::test_support::unique_temp_path;

    fn authority_of(values: &[&str]) -> HashSet<String> {
        values.iter().map(ToString::to_string).collect()
    }

    #[test]
    fn validates_reference_number_shape() {
        assert!(is_valid_ref_no("16867252734N"));
        assert!(!is_valid_ref_no("16867252734W")); // 结尾不是大写 N
        assert!(!is_valid_ref_no("1686725273N")); // 只有 10 位数字
        assert!(!is_valid_ref_no("16867252734n")); // 小写 n 不算
    }

    #[test]
    fn quantity_sign_matches_subsidy() {
        assert_eq!(quantity_of("10.00".parse().unwrap()), 1);
        assert_eq!(quantity_of("0.00".parse().unwrap()), 0);
        assert_eq!(quantity_of("-10.00".parse().unwrap()), -1);
    }

    #[test]
    fn as_is_hits_requires_isolated_match_in_authority() {
        let authority = authority_of(&["16867252734N"]);
        assert_eq!(
            as_is_hits("摘要16867252734N末尾", &authority),
            vec!["16867252734N".to_string()]
        );
        // 紧邻的后续数字使其成为更长数字串的一部分，不得原样提取。
        assert!(as_is_hits("摘要16867252734N9末尾", &authority).is_empty());
        // 格式正确但不在权威校验集中的值不得输出。
        assert!(as_is_hits("摘要99999999999N末尾", &authority).is_empty());
    }

    #[test]
    fn extracts_reference_via_as_is_match() {
        let authority = authority_of(&["16867252734N"]);
        assert_eq!(
            extract_reference("参考号：16867252734N。", &authority),
            Some("16867252734N".to_string())
        );
    }

    #[test]
    fn falls_back_to_zero_edit_when_n_suffix_missing() {
        let authority = authority_of(&["16867252734N"]);
        // 只有 11 位数字，缺少末尾大写 N；原样提取无结果，第三级补上 N 后命中。
        assert_eq!(
            extract_reference("单号16867252734完成", &authority),
            Some("16867252734N".to_string())
        );
    }

    #[test]
    fn trigger_letter_allows_digit_concatenation_across_gaps() {
        let authority = authority_of(&["16867252734N"]);
        // 数字被空格分割为两段（6 位+5 位，均不在 10-12 位范围内），
        // 但摘要含触发字母 N，可将全部数字按原顺序拼接为一个候选。
        assert_eq!(
            extract_reference("订单168672 52734N附言", &authority),
            Some("16867252734N".to_string())
        );
    }

    #[test]
    fn single_digit_insertion_recovers_reference() {
        let authority = authority_of(&["16867252734N"]);
        // 缺少末位数字 4（10 位），第四级允许一次插入。
        assert_eq!(
            extract_reference("单号1686725273结清", &authority),
            Some("16867252734N".to_string())
        );
    }

    #[test]
    fn single_digit_substitution_recovers_reference() {
        let authority = authority_of(&["16867252734N"]);
        // 末位数字错为 9（应为 4），第四级允许一次替换。
        assert_eq!(
            extract_reference("单号16867252739结清", &authority),
            Some("16867252734N".to_string())
        );
    }

    #[test]
    fn single_digit_deletion_recovers_reference() {
        let authority = authority_of(&["16867252734N"]);
        // 多出一位数字（12 位），第四级允许一次删除。
        assert_eq!(
            extract_reference("单号168672527340结清", &authority),
            Some("16867252734N".to_string())
        );
    }

    #[test]
    fn two_edits_are_not_attempted() {
        let authority = authority_of(&["16867252734N"]);
        // 两位数字都错误，超出“至多一次改动”的范围，不得纠正。
        assert_eq!(extract_reference("单号16867252799结清", &authority), None);
    }

    #[test]
    fn ambiguous_hits_at_same_priority_leave_reference_blank() {
        let authority = authority_of(&["11111111111N", "22222222222N"]);
        assert_eq!(
            extract_reference("含11111111111N及22222222222N两个编号", &authority),
            None
        );
    }

    #[test]
    fn blank_or_unrelated_summary_yields_none() {
        let authority = authority_of(&["16867252734N"]);
        assert_eq!(extract_reference("", &authority), None);
        assert_eq!(extract_reference("预售", &authority), None);
    }

    // --- 集成测试：构造一份门店银联样本建立权威校验集，再跑完整 CouponsJob。 ---

    fn write_unionpay_fixture(dir: &Path, retrieval_no: &str) {
        let mut workbook = Workbook::new();
        let sheet = workbook.add_worksheet();
        sheet.set_name("对账数据").unwrap();
        sheet.write_string(0, 0, "汇总").unwrap();
        for (col, header) in unionpay::HEADERS.iter().enumerate() {
            sheet.write_string(1, col as u16, *header).unwrap();
        }
        let row: [&str; 26] = [
            "20260914",
            "2026-09-14 10:18:09",
            "T001",
            "消费",
            "622***1234",
            "100.00",
            "100.00",
            "1.00",
            "0.50",
            "0.50",
            "SN0001",
            retrieval_no,
            "借记卡",
            "工商银行",
            "89813014812B06R",
            "某门店",
            "门店简",
            "ORDER1",
            "AC0001",
            "云闪付",
            "分店A",
            "0.00",
            "0.00",
            "备注文字",
            "",
            "buyer001",
        ];
        for (col, value) in row.iter().enumerate() {
            sheet.write_string(2, col as u16, *value).unwrap();
        }
        sheet.write_string(3, 0, unionpay::D1_NOTICE).unwrap();
        workbook
            .save(dir.join("89813014812B06R_MX_20260914101809_1.xlsx"))
            .unwrap();
    }

    fn write_coupons_workbook(dir: &Path, rows: &[(&str, &str, &str, &str, &str, &str, &str)]) {
        let mut workbook = Workbook::new();
        let sheet = workbook.add_worksheet();
        sheet.write_string(0, 0, TITLE).unwrap();
        for (col, header) in SOURCE_HEADERS.iter().enumerate() {
            sheet.write_string(1, col as u16, *header).unwrap();
        }
        for (index, (doc_no, doc_date, product_name, brand, finance_category, summary, total)) in
            rows.iter().enumerate()
        {
            let row = (2 + index) as u32;
            sheet
                .write_string(row, (COL_DOC_NO - 1) as u16, *doc_no)
                .unwrap();
            sheet
                .write_string(row, (COL_DOC_DATE - 1) as u16, *doc_date)
                .unwrap();
            sheet
                .write_string(row, (COL_PRODUCT_NAME - 1) as u16, *product_name)
                .unwrap();
            sheet
                .write_string(row, (COL_BRAND - 1) as u16, *brand)
                .unwrap();
            sheet
                .write_string(row, (COL_FINANCE_CATEGORY - 1) as u16, *finance_category)
                .unwrap();
            sheet
                .write_string(row, (COL_SUMMARY - 1) as u16, *summary)
                .unwrap();
            sheet
                .write_number(row, (COL_TOTAL - 1) as u16, total.parse::<f64>().unwrap())
                .unwrap();
        }
        sheet
            .write_string((2 + rows.len()) as u32, 0, "合计")
            .unwrap();
        workbook.save(dir.join(FILE_NAME)).unwrap();
    }

    #[test]
    fn end_to_end_extracts_reference_and_generates_match_doc_no() {
        let dir = unique_temp_path("coupons-happy-path");
        std::fs::create_dir_all(&dir).unwrap();

        write_unionpay_fixture(&dir, "16867252734N");
        write_coupons_workbook(
            &dir,
            &[
                (
                    "收款ZFFX000003",
                    "2026-08-29",
                    "商品甲",
                    "品牌甲",
                    "家电",
                    "参考号：16867252734N",
                    "299.85",
                ),
                (
                    "ZFFX000004",
                    "2026-08-30",
                    "商品乙",
                    "品牌乙",
                    "数码",
                    "无编号信息",
                    "-50.00",
                ),
            ],
        );

        let table = CouponsJob.run(&dir).unwrap();
        assert_eq!(table.columns.len(), 9);
        assert_eq!(table.rows.len(), 2);

        // 第一行：参考号原样命中，补贴额为正，数量为 1，匹配单据号保留“收款”前缀已剥离后拼接。
        assert_eq!(
            table.rows[0].values[0],
            Value::Text("收款ZFFX000003".to_string())
        );
        assert_eq!(
            table.rows[0].values[5],
            Value::Decimal("299.85".parse().unwrap())
        );
        assert_eq!(table.rows[0].values[6], Value::Integer(1));
        assert_eq!(
            table.rows[0].values[7],
            Value::Text("16867252734N".to_string())
        );
        assert_eq!(
            table.rows[0].values[8],
            Value::Text("260829ZFFX000003".to_string())
        );

        // 第二行：摘要中无编号，参考号留空；补贴额为负，数量为 -1。
        assert_eq!(table.rows[1].values[6], Value::Integer(-1));
        assert_eq!(table.rows[1].values[7], Value::Empty);
        assert_eq!(
            table.rows[1].values[8],
            Value::Text("260830ZFFX000004".to_string())
        );

        std::fs::remove_dir_all(&dir).unwrap();
    }

    #[test]
    fn rejects_title_mismatch() {
        let dir = unique_temp_path("coupons-bad-title");
        std::fs::create_dir_all(&dir).unwrap();
        write_unionpay_fixture(&dir, "16867252734N");

        let mut workbook = Workbook::new();
        let sheet = workbook.add_worksheet();
        sheet.write_string(0, 0, "错误标题").unwrap();
        for (col, header) in SOURCE_HEADERS.iter().enumerate() {
            sheet.write_string(1, col as u16, *header).unwrap();
        }
        sheet.write_string(2, 0, "合计").unwrap();
        workbook.save(dir.join(FILE_NAME)).unwrap();

        let error = CouponsJob.run(&dir).unwrap_err();
        assert!(matches!(error, ProcessError::Structure { .. }));

        std::fs::remove_dir_all(&dir).unwrap();
    }

    #[test]
    fn reports_no_input_when_authority_source_missing() {
        let dir = unique_temp_path("coupons-missing-authority");
        std::fs::create_dir_all(&dir).unwrap();
        write_coupons_workbook(
            &dir,
            &[(
                "Z1",
                "2026-08-29",
                "商品甲",
                "品牌甲",
                "家电",
                "无编号",
                "10.00",
            )],
        );

        // 门店银联样本文件缺失：必须停止，不得绕过权威校验集。
        let error = CouponsJob.run(&dir).unwrap_err();
        assert!(matches!(error, ProcessError::NoInput { .. }));

        std::fs::remove_dir_all(&dir).unwrap();
    }

    #[test]
    fn rejects_empty_subsidy_amount() {
        let dir = unique_temp_path("coupons-empty-subsidy");
        std::fs::create_dir_all(&dir).unwrap();
        write_unionpay_fixture(&dir, "16867252734N");

        let mut workbook = Workbook::new();
        let sheet = workbook.add_worksheet();
        sheet.write_string(0, 0, TITLE).unwrap();
        for (col, header) in SOURCE_HEADERS.iter().enumerate() {
            sheet.write_string(1, col as u16, *header).unwrap();
        }
        sheet
            .write_string(2, (COL_DOC_NO - 1) as u16, "Z1")
            .unwrap();
        // 合计（补贴额）留空。
        sheet.write_string(3, 0, "合计").unwrap();
        workbook.save(dir.join(FILE_NAME)).unwrap();

        let error = CouponsJob.run(&dir).unwrap_err();
        assert!(matches!(error, ProcessError::Data { .. }));

        std::fs::remove_dir_all(&dir).unwrap();
    }
}
