#![no_main]

use libfuzzer_sys::fuzz_target;
use rusdox::spec::DocumentSpec;
use rusdox::InputLimits;

fuzz_target!(|data: &[u8]| {
    let Ok(text) = std::str::from_utf8(data) else {
        return;
    };
    let limits = InputLimits {
        max_spec_bytes: 1024 * 1024,
        ..InputLimits::default()
    };
    let _ = DocumentSpec::from_yaml_str_with_limits(text, limits);
    let _ = DocumentSpec::from_json_str_with_limits(text, limits);
    let _ = DocumentSpec::from_toml_str_with_limits(text, limits);
});
