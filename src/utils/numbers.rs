use rust_decimal::Decimal;
use rust_decimal::prelude::FromPrimitive;

/// 把源单元格的`f64`转换为`Decimal`，使用最短往返表示，避免引入新的二进制误差
/// （`Decimal::from_f64`会去除多余的二进制尾差位，`from_f64_retain`则会保留，不适用）。
pub fn from_f64(value: f64) -> Option<Decimal> {
    Decimal::from_f64(value)
}

/// 第 7.3/8.4/10.7 节的浮点尾差规则：与最近两位小数的差不超过`0.000001`时，
/// 返回该两位小数值；否则返回`None`，由调用方按数据异常终止处理并报告上下文。
pub fn to_cents(value: Decimal) -> Option<Decimal> {
    let rounded = value.round_dp(2);
    let tolerance = Decimal::new(1, 6); // 0.000001
    if (value - rounded).abs() <= tolerance {
        Some(rounded)
    } else {
        None
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use std::str::FromStr;

    fn dec(text: &str) -> Decimal {
        Decimal::from_str(text).unwrap()
    }

    #[test]
    fn from_f64_avoids_binary_noise() {
        // 0.1 + 0.2 的二进制表示带有尾差；最短往返表示应直接得到 0.3。
        assert_eq!(from_f64(0.1 + 0.2), Some(dec("0.3")));
    }

    #[test]
    fn to_cents_accepts_exact_value() {
        assert_eq!(to_cents(dec("99.50")), Some(dec("99.50")));
    }

    #[test]
    fn to_cents_accepts_within_tolerance() {
        // project.md 10.4 节示例尾差值；作为 f64 字面量时与 299.85 是同一个比特模式。
        let value = from_f64(299.85).unwrap();
        assert_eq!(to_cents(value), Some(dec("299.85")));
    }

    #[test]
    fn to_cents_rejects_beyond_tolerance() {
        assert_eq!(to_cents(dec("99.505")), None);
    }
}
