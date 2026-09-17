use std::io::{self, Write};
use std::path::Path;

use crate::io::paths::resolve_input_dir;
use crate::jobs;
use crate::model::ProcessError;

use super::runner;

/// 程序入口：读取源文件夹路径与类别编号后执行一次，处理完成即退出。
pub fn run() -> Result<(), ProcessError> {
    let Some(input_dir) = read_input_dir()? else {
        return Ok(());
    };

    execute_once(&input_dir)
}

/// 读取源文件夹路径，无效时重新提示；读到输入结束（EOF）时返回`None`。
fn read_input_dir() -> Result<Option<std::path::PathBuf>, ProcessError> {
    loop {
        let Some(raw) = read_line("请输入源文件夹路径：")? else {
            return Ok(None);
        };

        match resolve_input_dir(&raw) {
            Ok(dir) => return Ok(Some(dir)),
            Err(error) => println!("{error}"),
        }
    }
}

/// 读取菜单编号并执行一次；无效编号重新提示，执行完成或读到输入结束后返回。
fn execute_once(input_dir: &Path) -> Result<(), ProcessError> {
    loop {
        print_menu();
        let Some(choice) = read_line("请输入编号（直接回车执行全部）：")? else {
            return Ok(());
        };

        if choice.is_empty() {
            runner::run_all(input_dir);
            return Ok(());
        }

        match choice.parse::<u32>() {
            Ok(number @ 1..=9) => {
                runner::run_one(input_dir, number);
                return Ok(());
            }
            _ => println!("无效编号：{choice}"),
        }
    }
}

fn print_menu() {
    println!("请选择要处理的数据类别：");
    for job in jobs::registry() {
        println!("  {} - {}", job.category() as u32, job.title());
    }
}

/// 读取一行输入，去除首尾空白；`read_line()`返回 0 字节（EOF）时返回`None`。
fn read_line(prompt: &str) -> Result<Option<String>, ProcessError> {
    print!("{prompt}");
    io::stdout().flush()?;

    let mut buffer = String::new();
    let bytes_read = io::stdin().read_line(&mut buffer)?;
    if bytes_read == 0 {
        return Ok(None);
    }
    Ok(Some(buffer.trim().to_string()))
}
