#![no_main]

use std::io::Cursor;

use libfuzzer_sys::fuzz_target;
use rusdox::{Document, DocumentMode, InputLimits};

fuzz_target!(|data: &[u8]| {
    let limits = InputLimits {
        max_docx_archive_bytes: 2 * 1024 * 1024,
        max_docx_entry_bytes: 2 * 1024 * 1024,
        max_docx_total_bytes: 8 * 1024 * 1024,
        ..InputLimits::default()
    };
    let _ = Document::open_from_reader_with_limits(
        Cursor::new(data),
        DocumentMode::ReadOnly,
        limits,
    );
});
