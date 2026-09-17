use std::cmp::Ordering;
use std::iter::Peekable;
use std::str::Chars;

/// 文件名自然排序：连续数字按数值大小比较，其他字符按原始字符顺序比较。
/// 例如`银联国补明细2.xlsx`排在`银联国补明细10.xlsx`之前。
pub fn natural_cmp(a: &str, b: &str) -> Ordering {
    let mut a_chars = a.chars().peekable();
    let mut b_chars = b.chars().peekable();

    loop {
        return match (a_chars.peek(), b_chars.peek()) {
            (None, None) => Ordering::Equal,
            (None, Some(_)) => Ordering::Less,
            (Some(_), None) => Ordering::Greater,
            (Some(ac), Some(bc)) if ac.is_ascii_digit() && bc.is_ascii_digit() => {
                match take_number(&mut a_chars).cmp(&take_number(&mut b_chars)) {
                    Ordering::Equal => continue,
                    other => other,
                }
            }
            (Some(ac), Some(bc)) => match ac.cmp(bc) {
                Ordering::Equal => {
                    a_chars.next();
                    b_chars.next();
                    continue;
                }
                other => other,
            },
        };
    }
}

fn take_number(chars: &mut Peekable<Chars>) -> u64 {
    let mut number: u64 = 0;
    while let Some(digit) = chars.peek().and_then(|c| c.to_digit(10)) {
        number = number * 10 + u64::from(digit);
        chars.next();
    }
    number
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn numeric_runs_compare_by_value_not_by_text() {
        assert_eq!(
            natural_cmp("银联国补明细2.xlsx", "银联国补明细10.xlsx"),
            Ordering::Less
        );
    }

    #[test]
    fn non_numeric_text_compares_lexically() {
        assert_eq!(natural_cmp("a.xlsx", "b.xlsx"), Ordering::Less);
    }

    #[test]
    fn identical_names_are_equal() {
        assert_eq!(
            natural_cmp("银联国补明细2.xlsx", "银联国补明细2.xlsx"),
            Ordering::Equal
        );
    }

    #[test]
    fn sorts_a_full_file_list_in_natural_order() {
        let mut names = vec![
            "银联国补明细10.xlsx",
            "银联国补明细2.xlsx",
            "银联国补明细1.xlsx",
        ];
        names.sort_by(|a, b| natural_cmp(a, b));
        assert_eq!(
            names,
            vec![
                "银联国补明细1.xlsx",
                "银联国补明细2.xlsx",
                "银联国补明细10.xlsx",
            ]
        );
    }
}
