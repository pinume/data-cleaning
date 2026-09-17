use std::path::Path;

use rust_decimal::Decimal;

use crate::io::xlsx_reader::RawCell;
use crate::model::{ProcessError, Table, Value};
use crate::utils::{dates, numbers, text};

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum Category {
    Invoice = 1,
    UnionSubsidy,
    UploadedDigital,
    UploadedAppliance,
    UnionPay,
    RefundAppliance,
    RefundDigital,
    Receipts,
    Coupons,
}

pub trait Job {
    fn category(&self) -> Category;
    fn title(&self) -> &'static str;
    fn output_stem(&self) -> &'static str;
    fn run(&self, input_dir: &Path) -> Result<Table, ProcessError>;
}

/// 单元格的通用文本表示，仅用于错误报告与重复导出指纹等诊断用途，不代表最终输出值。
pub(crate) fn cell_display(cell: &RawCell) -> String {
    match cell {
        RawCell::Empty => String::new(),
        RawCell::Text(text) => text.clone(),
        RawCell::Int(n) => n.to_string(),
        RawCell::Float(f) => f.to_string(),
        RawCell::Bool(b) => b.to_string(),
        RawCell::DateTime(serial) => serial.to_string(),
        RawCell::Error(message) => message.clone(),
    }
}

/// 文本字段：保留原值（包括纯空白），仅数值型单元格需还原为完整整数文本。
pub(crate) fn cell_text(cell: &RawCell) -> Result<String, String> {
    match cell {
        RawCell::Empty => Ok(String::new()),
        RawCell::Text(text) => Ok(text.clone()),
        RawCell::Int(n) => Ok(n.to_string()),
        RawCell::Bool(b) => Ok(b.to_string()),
        RawCell::Float(f) => text::identifier_from_float(*f)
            .ok_or_else(|| format!("数值 {f} 带小数或超出精度范围，无法还原为完整文本")),
        RawCell::DateTime(_) => Err("此字段不应为日期类型".to_string()),
        RawCell::Error(message) => Err(format!("单元格为错误值：{message}")),
    }
}

/// 空字符串转换为`Value::Empty`，否则包装为`Value::Text`。
pub(crate) fn text_value(value: String) -> Value {
    if value.is_empty() {
        Value::Empty
    } else {
        Value::Text(value)
    }
}

/// `None`转换为`Value::Empty`，否则包装为`Value::Decimal`。
pub(crate) fn amount_value(value: Option<Decimal>) -> Value {
    value.map(Value::Decimal).unwrap_or(Value::Empty)
}

/// 数值字段：为空时保持为空，不做尾差取整（显示格式如`0.00`只影响展示，不改变实际值）。
/// 部分导出会把金额存成文本，按十进制文本解析（避免不必要的二进制浮点转换）。
pub(crate) fn cell_amount(cell: &RawCell) -> Result<Option<Decimal>, String> {
    match cell {
        RawCell::Empty => Ok(None),
        RawCell::Text(t) if t.trim().is_empty() => Ok(None),
        RawCell::Text(t) => t
            .trim()
            .parse::<Decimal>()
            .map(Some)
            .map_err(|_| format!("文本“{t}”无法解析为十进制金额")),
        RawCell::Int(n) => Ok(Some(Decimal::from(*n))),
        RawCell::Float(f) => numbers::from_f64(*f)
            .map(Some)
            .ok_or_else(|| format!("数值 {f} 无法转换为十进制金额")),
        other => Err(format!("金额字段出现非数值内容：{}", cell_display(other))),
    }
}

/// 日期字段：源值为 Excel 日期序列值或`yyyymmdd`文本；无法识别时保留原值，不猜测修正
/// （第 5、6 节的通用豁免）。
pub(crate) fn cell_date_or_text(cell: &RawCell) -> Value {
    if let RawCell::DateTime(serial) = cell {
        return dates::date_from_serial(*serial)
            .map(Value::Date)
            .unwrap_or_else(|| Value::Text(cell_display(cell)));
    }
    let text = cell_display(cell);
    if text.trim().is_empty() {
        return Value::Empty;
    }
    dates::parse_yyyymmdd(&text)
        .map(Value::Date)
        .unwrap_or(Value::Text(text))
}

/// 日期时间字段：保留到秒；无法识别时保留原值，不猜测修正（第 5、6 节的通用豁免）。
pub(crate) fn cell_datetime_or_text(cell: &RawCell) -> Value {
    if let RawCell::DateTime(serial) = cell {
        return dates::datetime_from_serial(*serial)
            .map(Value::DateTime)
            .unwrap_or_else(|| Value::Text(cell_display(cell)));
    }
    let text = cell_display(cell);
    if text.trim().is_empty() {
        return Value::Empty;
    }
    dates::parse_date_text(&text)
        .map(Value::DateTime)
        .unwrap_or(Value::Text(text))
}

pub mod coupons;
pub mod invoice;
pub mod receipts;
pub mod refund;
pub mod union_subsidy;
pub mod unionpay;
pub mod uploaded;

/// 按菜单编号 1–9 排列的全部任务；`app::runner`据此驱动单类或批量执行。
pub fn registry() -> Vec<Box<dyn Job>> {
    vec![
        Box::new(invoice::InvoiceJob),
        Box::new(union_subsidy::UnionSubsidyJob),
        Box::new(uploaded::UPLOADED_DIGITAL),
        Box::new(uploaded::UPLOADED_APPLIANCE),
        Box::new(unionpay::UnionPayJob),
        Box::new(refund::REFUND_APPLIANCE),
        Box::new(refund::REFUND_DIGITAL),
        Box::new(receipts::ReceiptsJob),
        Box::new(coupons::CouponsJob),
    ]
}

#[cfg(test)]
mod tests {
    use super::*;
    use std::str::FromStr;

    #[test]
    fn registry_is_ordered_by_menu_number_one_through_nine() {
        let numbers: Vec<u32> = registry().iter().map(|job| job.category() as u32).collect();
        assert_eq!(numbers, (1..=9).collect::<Vec<_>>());
    }

    #[test]
    fn cell_amount_parses_numeric_text() {
        assert_eq!(
            cell_amount(&RawCell::Text("1000.00".to_string())),
            Ok(Some(Decimal::from_str("1000.00").unwrap()))
        );
    }

    #[test]
    fn cell_amount_blank_text_is_none() {
        assert_eq!(cell_amount(&RawCell::Text("   ".to_string())), Ok(None));
        assert_eq!(cell_amount(&RawCell::Empty), Ok(None));
    }

    #[test]
    fn cell_amount_rejects_non_numeric_text() {
        assert!(cell_amount(&RawCell::Text("不是数字".to_string())).is_err());
    }

    #[test]
    fn cell_text_preserves_whitespace_only_text() {
        assert_eq!(
            cell_text(&RawCell::Text("  ".to_string())),
            Ok("  ".to_string())
        );
    }

    #[test]
    fn text_value_and_amount_value_map_empty_to_value_empty() {
        assert_eq!(text_value(String::new()), Value::Empty);
        assert_eq!(text_value("x".to_string()), Value::Text("x".to_string()));
        assert_eq!(amount_value(None), Value::Empty);
    }
}
