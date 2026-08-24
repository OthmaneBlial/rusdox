//! Generated authoring schema for the versioned RusDox document spec.

use serde_json::Value;

use crate::spec::{DocumentSpec, SPEC_VERSION};
use crate::{DocxError, Result};

/// Canonical public identifier for the version 1 document-spec schema.
pub const DOCUMENT_SPEC_SCHEMA_ID: &str =
    "https://othmaneblial.github.io/rusdox/schema/rusdox-spec-v1.schema.json";

/// Generates the JSON Schema directly from the Rust/serde document model.
pub fn document_spec_schema() -> Result<Value> {
    let schema = schemars::schema_for!(DocumentSpec);
    let mut value = serde_json::to_value(schema).map_err(|error| {
        DocxError::parse(format!("failed to generate document schema: {error}"))
    })?;
    let object = value
        .as_object_mut()
        .ok_or_else(|| DocxError::parse("generated document schema root is not an object"))?;
    object.insert(
        "$id".to_string(),
        Value::String(DOCUMENT_SPEC_SCHEMA_ID.to_string()),
    );
    object.insert(
        "title".to_string(),
        Value::String("RusDox document specification".to_string()),
    );
    object.insert(
        "description".to_string(),
        Value::String(
            "Schema for YAML, JSON, and TOML RusDox authoring. Files are local-first and deterministic."
                .to_string(),
        ),
    );
    object.insert(
        "x-rusdox-spec-version".to_string(),
        Value::from(SPEC_VERSION),
    );
    object.insert(
        "x-rusdox-authoring-formats".to_string(),
        serde_json::json!(["yaml", "json", "toml"]),
    );
    let properties = object
        .get_mut("properties")
        .and_then(Value::as_object_mut)
        .ok_or_else(|| DocxError::parse("generated document schema has no properties object"))?;
    properties.insert(
        "variables".to_string(),
        serde_json::json!({
            "type": "object",
            "description": "Local values available to deterministic {{ path | filter }} expressions, repeat blocks, and when blocks.",
            "additionalProperties": true,
            "default": {}
        }),
    );
    if let Some(version) = properties.get_mut("version").and_then(Value::as_object_mut) {
        version.insert("const".to_string(), Value::from(SPEC_VERSION));
    }
    object.insert(
        "required".to_string(),
        serde_json::json!(["version", "blocks"]),
    );

    let variants = object
        .get_mut("$defs")
        .and_then(Value::as_object_mut)
        .and_then(|definitions| definitions.get_mut("BlockSpec"))
        .and_then(Value::as_object_mut)
        .and_then(|block| block.get_mut("oneOf"))
        .and_then(Value::as_array_mut)
        .ok_or_else(|| DocxError::parse("generated document schema has no BlockSpec variants"))?;
    variants.extend(authoring_block_schemas());
    Ok(value)
}

fn authoring_block_schemas() -> Vec<Value> {
    vec![
        serde_json::json!({
            "type": "object",
            "description": "Inline a local block fragment.",
            "properties": {
                "type": {"const": "include"},
                "path": {"type": "string"},
                "variables": {"type": "object", "additionalProperties": true}
            },
            "required": ["type", "path"],
            "additionalProperties": false
        }),
        serde_json::json!({
            "type": "object",
            "description": "Repeat blocks for a local array. Exactly one of variable or items is required.",
            "properties": {
                "type": {"const": "repeat"},
                "variable": {"type": "string"},
                "items": {"type": "array"},
                "as": {"type": "string", "default": "item"},
                "blocks": {"type": "array", "items": {"$ref": "#/$defs/BlockSpec"}}
            },
            "required": ["type", "blocks"],
            "oneOf": [
                {"required": ["variable"], "not": {"required": ["items"]}},
                {"required": ["items"], "not": {"required": ["variable"]}}
            ],
            "additionalProperties": false
        }),
        serde_json::json!({
            "type": "object",
            "description": "Select one deterministic block branch from a data path and optional scalar equality.",
            "properties": {
                "type": {"const": "when"},
                "path": {"type": "string"},
                "equals": {},
                "blocks": {"type": "array", "items": {"$ref": "#/$defs/BlockSpec"}},
                "otherwise": {"type": "array", "items": {"$ref": "#/$defs/BlockSpec"}, "default": []}
            },
            "required": ["type", "path", "blocks"],
            "additionalProperties": false
        }),
    ]
}

/// Generates stable, pretty-printed JSON suitable for checked-in tooling assets.
pub fn document_spec_schema_pretty() -> Result<String> {
    let mut rendered = serde_json::to_string_pretty(&document_spec_schema()?).map_err(|error| {
        DocxError::parse(format!("failed to serialize document schema: {error}"))
    })?;
    rendered.push('\n');
    Ok(rendered)
}

#[cfg(test)]
mod tests {
    use super::{document_spec_schema, document_spec_schema_pretty, DOCUMENT_SPEC_SCHEMA_ID};

    #[test]
    fn schema_is_versioned_and_exposes_block_variants() {
        let schema = document_spec_schema().expect("schema");
        assert_eq!(schema["$id"], DOCUMENT_SPEC_SCHEMA_ID);
        assert_eq!(schema["x-rusdox-spec-version"], 1);
        let rendered = document_spec_schema_pretty().expect("rendered schema");
        for variant in ["title", "paragraph", "table", "image", "table_of_contents"] {
            assert!(
                rendered.contains(&format!("\"{variant}\"")),
                "missing {variant}"
            );
        }
    }

    #[test]
    fn checked_in_and_editor_schemas_match_the_generator() {
        let generated = document_spec_schema_pretty().expect("generated schema");
        assert_eq!(
            generated,
            include_str!("../schema/rusdox-spec-v1.schema.json")
        );
        assert_eq!(
            generated,
            include_str!("../editors/vscode/schema/rusdox-spec-v1.schema.json")
        );
    }
}
