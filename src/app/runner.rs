use std::panic::{self, AssertUnwindSafe};
use std::path::{Path, PathBuf};

use crate::io::paths::ensure_output_dir;
use crate::io::publisher;
use crate::io::xlsx_writer::write_table;
use crate::jobs::{self, Job};
use crate::model::{ProcessError, Table};

/// 空输入：按编号 1–9 依次执行全部类别；单个类别失败或 panic 不中断其余类别，
/// 逐项报告各自的成功或失败结果。
pub fn run_all(input_dir: &Path) {
    for job in jobs::registry() {
        execute(*job, input_dir);
    }
}

/// 单一编号：只执行并发布该类别的结果。调用方（`cli.rs`）保证`category_number`
/// 已在`1..=9`范围内校验过，`registry()`固定收录这9个类别，查找必然命中。
pub fn run_one(input_dir: &Path, category_number: u32) {
    let job = jobs::registry()
        .iter()
        .find(|job| job.category() as u32 == category_number)
        .expect("category_number 应已由调用方校验在 registry() 收录的范围内");
    execute(*job, input_dir);
}

fn execute(job: &dyn Job, input_dir: &Path) {
    let title = job.title();
    // 用 catch_unwind 隔离单个任务的 panic，避免其中断批量执行或整个程序。
    let outcome = panic::catch_unwind(AssertUnwindSafe(|| job.run(input_dir)))
        .unwrap_or_else(|payload| {
            Err(ProcessError::Read(format!(
                "处理过程中发生程序内部错误：{}",
                panic_message(&payload)
            )))
        })
        .and_then(|table| publish(job, input_dir, &table));

    match outcome {
        Ok(path) => println!("[{title}] 处理成功：{}", path.display()),
        Err(error) => println!("[{title}] 处理失败：{error}"),
    }
}

fn publish(job: &dyn Job, input_dir: &Path, table: &Table) -> Result<PathBuf, ProcessError> {
    let output_dir = ensure_output_dir(input_dir)?;
    let temp = publisher::temp_path(&output_dir, job.output_stem());
    write_table(table, &temp)?;
    publisher::publish(&output_dir, job.output_stem(), &temp)
}

fn panic_message(payload: &(dyn std::any::Any + Send)) -> String {
    if let Some(message) = payload.downcast_ref::<&str>() {
        message.to_string()
    } else if let Some(message) = payload.downcast_ref::<String>() {
        message.clone()
    } else {
        "未知错误".to_string()
    }
}

#[cfg(test)]
mod tests {
    use rust_xlsxwriter::Workbook;

    use super::*;
    use crate::io::xlsx_reader::open_sheets;
    use crate::jobs::unionpay;
    use crate::test_support::unique_temp_path;

    /// 端到端验证 `execute()` 真正把 job 的结果写盘并发布到正确的兄弟 `source_data/` 目录，
    /// 而不仅仅是业务逻辑本身（业务逻辑已由各 job 自己的单元测试覆盖）。
    #[test]
    fn execute_publishes_real_file_to_sibling_output_dir() {
        let base = unique_temp_path("runner-end-to-end");
        let input_dir = base.join("输入目录");
        std::fs::create_dir_all(&input_dir).unwrap();

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
            "16867252734N",
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
            .save(input_dir.join("89813014812B06R_MX_20260914101809_1.xlsx"))
            .unwrap();

        run_one(&input_dir, 5);

        let output_path = base.join("source_data").join("银联交易明细门店.xlsx");
        assert!(
            output_path.is_file(),
            "应在输入目录同级的 source_data/ 下生成结果文件"
        );

        // 确认写出的文件是一个真正可被读回的 XLSX（而不仅仅是存在同名文件）。
        let sheets = open_sheets(&output_path).unwrap();
        assert_eq!(sheets.len(), 1);
        assert_eq!(sheets[0].row_texts(1), unionpay::HEADERS.to_vec());

        std::fs::remove_dir_all(&base).unwrap();
    }
}
