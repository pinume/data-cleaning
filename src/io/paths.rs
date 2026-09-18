use std::path::{Path, PathBuf};

use crate::model::ProcessError;

/// 去除首尾空白，并在结果首尾为成对单引号或双引号时去除该对引号。
fn normalize_raw_path(raw: &str) -> String {
    let trimmed = raw.trim();
    let unquoted = trimmed
        .strip_prefix('"')
        .and_then(|s| s.strip_suffix('"'))
        .or_else(|| trimmed.strip_prefix('\'').and_then(|s| s.strip_suffix('\'')))
        .unwrap_or(trimmed);
    unquoted.trim().to_string()
}

/// 解析用户输入的源文件夹路径：去首尾空白与成对引号，相对路径以当前工作目录为基准
/// 解析为绝对路径，并确认该路径存在、是文件夹且可读取。
pub fn resolve_input_dir(raw: &str) -> Result<PathBuf, ProcessError> {
    let normalized = normalize_raw_path(raw);
    if normalized.is_empty() {
        return Err(ProcessError::InvalidPath {
            path: PathBuf::new(),
            reason: "路径为空".to_string(),
        });
    }

    let candidate = Path::new(&normalized);
    let absolute = if candidate.is_absolute() {
        candidate.to_path_buf()
    } else {
        std::env::current_dir()?.join(candidate)
    };

    if !absolute.exists() {
        return Err(ProcessError::InvalidPath {
            path: absolute,
            reason: "路径不存在".to_string(),
        });
    }
    if !absolute.is_dir() {
        return Err(ProcessError::InvalidPath {
            path: absolute,
            reason: "不是文件夹".to_string(),
        });
    }
    std::fs::read_dir(&absolute).map_err(|source| ProcessError::InvalidPath {
        path: absolute.clone(),
        reason: format!("不可读取：{source}"),
    })?;

    Ok(absolute)
}

/// 列出`输入目录`直接包含的 `.xlsx` 文件，不递归，排除 `~$` 开头的临时文件。
pub fn list_xlsx_files(dir: &Path) -> Result<Vec<PathBuf>, ProcessError> {
    let mut files = Vec::new();
    for entry in std::fs::read_dir(dir)? {
        let entry = entry?;
        let path = entry.path();
        if !path.is_file() {
            continue;
        }
        let Some(name) = path.file_name().and_then(|n| n.to_str()) else {
            continue;
        };
        if name.starts_with("~$") {
            continue;
        }
        if path.extension().and_then(|ext| ext.to_str()) == Some("xlsx") {
            files.push(path);
        }
    }
    Ok(files)
}

/// 计算并确保`输入目录`的父目录下的`输出目录`（`source_data/`）存在。
pub fn ensure_output_dir(input_dir: &Path) -> Result<PathBuf, ProcessError> {
    let parent = input_dir
        .parent()
        .ok_or_else(|| ProcessError::InvalidPath {
            path: input_dir.to_path_buf(),
            reason: "无法确定父目录".to_string(),
        })?;
    let output_dir = parent.join("source_data");
    std::fs::create_dir_all(&output_dir)?;
    Ok(output_dir)
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::test_support::unique_temp_path;

    fn scratch_dir(label: &str) -> PathBuf {
        let dir = unique_temp_path(&format!("paths-test-{label}"));
        std::fs::create_dir_all(&dir).unwrap();
        dir
    }

    #[test]
    fn normalize_trims_whitespace() {
        assert_eq!(normalize_raw_path("  /a/b  "), "/a/b");
    }

    #[test]
    fn normalize_strips_matching_single_quotes() {
        assert_eq!(normalize_raw_path("'/a/b'"), "/a/b");
    }

    #[test]
    fn normalize_strips_matching_double_quotes() {
        assert_eq!(normalize_raw_path("  \"/a/b\"  "), "/a/b");
    }

    #[test]
    fn normalize_keeps_unmatched_quote() {
        assert_eq!(normalize_raw_path("'/a/b"), "'/a/b");
    }

    #[test]
    fn resolve_input_dir_rejects_empty_input() {
        let error = resolve_input_dir("   ").unwrap_err();
        assert!(matches!(error, ProcessError::InvalidPath { .. }));
    }

    #[test]
    fn resolve_input_dir_rejects_missing_path() {
        let dir = scratch_dir("missing-base");
        let missing = dir.join("does-not-exist");
        let error = resolve_input_dir(missing.to_str().unwrap()).unwrap_err();
        assert!(matches!(error, ProcessError::InvalidPath { .. }));
        std::fs::remove_dir_all(&dir).unwrap();
    }

    #[test]
    fn resolve_input_dir_rejects_file_path() {
        let dir = scratch_dir("file-not-dir");
        let file = dir.join("file.txt");
        std::fs::write(&file, b"x").unwrap();
        let error = resolve_input_dir(file.to_str().unwrap()).unwrap_err();
        assert!(matches!(error, ProcessError::InvalidPath { .. }));
        std::fs::remove_dir_all(&dir).unwrap();
    }

    #[test]
    fn resolve_input_dir_accepts_quoted_absolute_dir() {
        let dir = scratch_dir("ok-dir");
        let quoted = format!("  \"{}\"  ", dir.display());
        let resolved = resolve_input_dir(&quoted).unwrap();
        assert_eq!(resolved, dir);
        std::fs::remove_dir_all(&dir).unwrap();
    }

    #[test]
    fn list_xlsx_files_filters_by_extension_and_temp_prefix() {
        let dir = scratch_dir("list-xlsx");
        std::fs::write(dir.join("发票_20260914.xlsx"), b"x").unwrap();
        std::fs::write(dir.join("~$发票_20260914.xlsx"), b"x").unwrap();
        std::fs::write(dir.join("说明.txt"), b"x").unwrap();
        std::fs::create_dir(dir.join("子目录.xlsx")).unwrap();

        let mut files: Vec<String> = list_xlsx_files(&dir)
            .unwrap()
            .into_iter()
            .map(|p| p.file_name().unwrap().to_string_lossy().into_owned())
            .collect();
        files.sort();

        assert_eq!(files, vec!["发票_20260914.xlsx".to_string()]);
        std::fs::remove_dir_all(&dir).unwrap();
    }

    #[test]
    fn ensure_output_dir_creates_sibling_output_folder() {
        let base = scratch_dir("output-base");
        let input_dir = base.join("源数据");
        std::fs::create_dir_all(&input_dir).unwrap();

        let output_dir = ensure_output_dir(&input_dir).unwrap();

        assert_eq!(output_dir, base.join("source_data"));
        assert!(output_dir.is_dir());
        std::fs::remove_dir_all(&base).unwrap();
    }
}
