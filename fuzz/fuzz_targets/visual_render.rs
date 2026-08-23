#![no_main]

use std::io::Cursor;

use libfuzzer_sys::fuzz_target;
use rusdox::{Document, Visual, VisualFormat};

fuzz_target!(|data: &[u8]| {
    let Some((&selector, payload)) = data.split_first() else {
        return;
    };
    let format = match selector % 3 {
        0 => VisualFormat::Png,
        1 => VisualFormat::Jpeg,
        _ => VisualFormat::Svg,
    };
    let document = Document::new().add_visual(Visual::from_bytes(payload.to_vec(), format));
    let _ = document.save_to_writer(Cursor::new(Vec::new()));
});
