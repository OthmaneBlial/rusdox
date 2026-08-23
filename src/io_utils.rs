use std::fs;
use std::io::Write;
use std::path::Path;

use tempfile::NamedTempFile;

use crate::{DocxError, Result};

/// Atomically replaces a file after fully writing and syncing a same-directory temporary file.
pub fn atomic_write(path: impl AsRef<Path>, bytes: &[u8]) -> Result<()> {
    atomic_write_with(path.as_ref(), |file| {
        file.write_all(bytes)?;
        Ok(())
    })
}

pub(crate) fn atomic_write_with<F>(path: &Path, write: F) -> Result<()>
where
    F: FnOnce(&mut fs::File) -> Result<()>,
{
    let parent = path.parent().unwrap_or_else(|| Path::new("."));
    fs::create_dir_all(parent)?;
    let mut temporary = NamedTempFile::new_in(parent)?;

    write(temporary.as_file_mut())?;
    temporary.as_file_mut().flush()?;
    temporary.as_file().sync_all()?;
    temporary
        .persist(path)
        .map_err(|error| DocxError::Io(error.error))?;

    sync_parent_directory(parent)?;
    Ok(())
}

#[cfg(unix)]
fn sync_parent_directory(parent: &Path) -> Result<()> {
    fs::File::open(parent)?.sync_all()?;
    Ok(())
}

#[cfg(not(unix))]
fn sync_parent_directory(_parent: &Path) -> Result<()> {
    Ok(())
}

#[cfg(test)]
mod tests {
    use std::fs;
    use std::io::{Error, ErrorKind, Write};

    use tempfile::tempdir;

    use super::atomic_write_with;

    #[test]
    fn interrupted_atomic_write_preserves_previous_file() {
        let temp = tempdir().expect("temp dir");
        let destination = temp.path().join("recoverable.bin");
        fs::write(&destination, b"known-good").expect("seed destination");

        let result = atomic_write_with(&destination, |writer| {
            writer.write_all(b"partial")?;
            Err(Error::new(ErrorKind::Interrupted, "simulated interruption").into())
        });

        assert!(result.is_err());
        assert_eq!(
            fs::read(&destination).expect("read destination"),
            b"known-good"
        );
        let leftovers = fs::read_dir(temp.path())
            .expect("read temp dir")
            .filter_map(std::result::Result::ok)
            .collect::<Vec<_>>();
        assert_eq!(leftovers.len(), 1, "temporary output should be cleaned");
    }
}
