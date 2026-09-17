use std::path::{Path, PathBuf};
use std::time::{SystemTime, UNIX_EPOCH};

use crate::model::ProcessError;

/// 生成本次发布使用的临时文件路径：`.<基名>.<随机后缀>.xlsx.tmp`。
pub fn temp_path(output_dir: &Path, stem: &str) -> PathBuf {
    let nanos = SystemTime::now()
        .duration_since(UNIX_EPOCH)
        .map(|elapsed| elapsed.as_nanos())
        .unwrap_or_default();
    output_dir.join(format!(".{stem}.{}.{nanos:x}.xlsx.tmp", std::process::id()))
}

/// 确认临时文件已生成且非空后，备份并替换正式文件；任一步失败时清理临时文件、
/// 恢复原有结果并返回错误。全部成功后删除备份，返回正式文件路径。
pub fn publish(output_dir: &Path, stem: &str, temp: &Path) -> Result<PathBuf, ProcessError> {
    let target = output_dir.join(format!("{stem}.xlsx"));

    let result = (|| -> Result<(), ProcessError> {
        let metadata = std::fs::metadata(temp)?;
        if metadata.len() == 0 {
            return Err(ProcessError::Io(std::io::Error::other(format!(
                "临时文件为空：{}",
                temp.display()
            ))));
        }

        let backup = output_dir.join(format!(".{stem}.xlsx.bak"));
        let had_target = target.exists();
        if had_target {
            std::fs::rename(&target, &backup)?;
        }

        if let Err(error) = std::fs::rename(temp, &target) {
            if had_target {
                let _ = std::fs::rename(&backup, &target);
            }
            return Err(ProcessError::Io(error));
        }

        if had_target {
            let _ = std::fs::remove_file(&backup);
        }

        Ok(())
    })();

    if result.is_err() {
        let _ = std::fs::remove_file(temp);
    }

    result.map(|()| target)
}
