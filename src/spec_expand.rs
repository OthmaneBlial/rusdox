use std::fs;
use std::path::{Path, PathBuf};

use serde_yaml::{Mapping, Number, Value};

use crate::{DocxError, Result};

pub(crate) fn expand_yaml_document_spec(
    content: &str,
    source_path: Option<&Path>,
) -> Result<Value> {
    let root: Value = serde_yaml::from_str(content)
        .map_err(|error| DocxError::parse(format!("invalid YAML document spec: {error}")))?;
    let mut state = ExpansionState::default();
    if let Some(source_path) = source_path {
        state
            .include_stack
            .push(canonical_or_absolute(source_path)?);
    }
    expand_document_root(
        root,
        source_path.and_then(Path::parent),
        &Mapping::new(),
        None,
        &mut state,
    )
}

#[derive(Default)]
struct ExpansionState {
    include_stack: Vec<PathBuf>,
}

fn expand_document_root(
    root: Value,
    current_dir: Option<&Path>,
    parent_vars: &Mapping,
    override_vars: Option<Mapping>,
    state: &mut ExpansionState,
) -> Result<Value> {
    let mut mapping = expect_mapping(root, "YAML document spec root must be a mapping")?;
    let variables = build_variable_context(&mut mapping, parent_vars, override_vars)?;

    let blocks = match mapping.remove(string_key("blocks")) {
        Some(value) => expand_block_sequence(value, current_dir, &variables, state)?,
        None => Vec::new(),
    };

    let mut expanded = Mapping::new();
    for (key, value) in mapping {
        if key == string_key("variables") {
            continue;
        }
        expanded.insert(key, expand_value(value, &variables)?);
    }
    expanded.insert(string_key("blocks"), Value::Sequence(blocks));
    Ok(Value::Mapping(expanded))
}

fn build_variable_context(
    mapping: &mut Mapping,
    parent_vars: &Mapping,
    override_vars: Option<Mapping>,
) -> Result<Mapping> {
    let mut merged = parent_vars.clone();

    if let Some(local_vars) = mapping.remove(string_key("variables")) {
        let local_mapping = expect_mapping(local_vars, "`variables` must be a mapping")?;
        for (key, value) in local_mapping {
            let name = expect_string_key(&key, "variable names must be strings")?;
            let expanded = expand_value(value, &merged)?;
            merged.insert(Value::String(name), expanded);
        }
    }

    if let Some(overrides) = override_vars {
        for (key, value) in overrides {
            let name = expect_string_key(&key, "override variable names must be strings")?;
            let expanded = expand_value(value, &merged)?;
            merged.insert(Value::String(name), expanded);
        }
    }

    Ok(merged)
}

fn expand_block_sequence(
    value: Value,
    current_dir: Option<&Path>,
    variables: &Mapping,
    state: &mut ExpansionState,
) -> Result<Vec<Value>> {
    let blocks = expect_sequence(value, "`blocks` must be a sequence")?;
    let mut expanded = Vec::new();

    for block in blocks {
        let Some(block_type) = block_type(&block) else {
            expanded.push(expand_value(block, variables)?);
            continue;
        };

        match block_type.as_str() {
            "include" => {
                expanded.extend(expand_include_block(block, current_dir, variables, state)?)
            }
            "repeat" => expanded.extend(expand_repeat_block(block, current_dir, variables, state)?),
            _ => expanded.push(expand_value(block, variables)?),
        }
    }

    Ok(expanded)
}

fn expand_include_block(
    block: Value,
    current_dir: Option<&Path>,
    variables: &Mapping,
    state: &mut ExpansionState,
) -> Result<Vec<Value>> {
    let mut mapping = expect_mapping(block, "`include` block must be a mapping")?;
    let path_value = mapping
        .remove(string_key("path"))
        .ok_or_else(|| DocxError::parse("`include` block requires a `path` field"))?;
    let path_value = expand_value(path_value, variables)?;
    let include_path = scalar_to_string(&path_value)
        .ok_or_else(|| DocxError::parse("`include.path` must resolve to a scalar string"))?;

    let override_vars = match mapping.remove(string_key("variables")) {
        Some(value) => Some(expect_mapping(
            expand_value(value, variables)?,
            "`include.variables` must be a mapping",
        )?),
        None => None,
    };

    let resolved = resolve_include_path(current_dir, &include_path)?;
    let canonical = canonical_or_absolute(&resolved)?;
    if state.include_stack.iter().any(|path| path == &canonical) {
        let mut cycle = state
            .include_stack
            .iter()
            .map(|path| path.display().to_string())
            .collect::<Vec<_>>();
        cycle.push(canonical.display().to_string());
        return Err(DocxError::parse(format!(
            "YAML include cycle detected: {}",
            cycle.join(" -> ")
        )));
    }

    let content = fs::read_to_string(&canonical)?;
    let root: Value = serde_yaml::from_str(&content).map_err(|error| {
        DocxError::parse(format!(
            "invalid included YAML fragment '{}': {error}",
            canonical.display()
        ))
    })?;

    state.include_stack.push(canonical.clone());
    let result = expand_include_root(root, canonical.parent(), variables, override_vars, state);
    state.include_stack.pop();
    result
}

fn expand_include_root(
    root: Value,
    current_dir: Option<&Path>,
    parent_vars: &Mapping,
    override_vars: Option<Mapping>,
    state: &mut ExpansionState,
) -> Result<Vec<Value>> {
    match root {
        Value::Sequence(_) => {
            let variables = if let Some(overrides) = override_vars {
                build_variable_context(&mut Mapping::new(), parent_vars, Some(overrides))?
            } else {
                parent_vars.clone()
            };
            expand_block_sequence(root, current_dir, &variables, state)
        }
        Value::Mapping(mapping) => {
            if mapping.contains_key(string_key("blocks")) {
                let expanded = expand_document_root(
                    Value::Mapping(mapping),
                    current_dir,
                    parent_vars,
                    override_vars,
                    state,
                )?;
                let mut expanded_mapping =
                    expect_mapping(expanded, "expanded include root must stay a mapping")?;
                let blocks = expanded_mapping
                    .remove(string_key("blocks"))
                    .ok_or_else(|| DocxError::parse("expanded include root lost `blocks`"))?;
                expect_sequence(blocks, "expanded include blocks must stay a sequence")
            } else if mapping.contains_key(string_key("type")) {
                let variables = if let Some(overrides) = override_vars {
                    build_variable_context(&mut Mapping::new(), parent_vars, Some(overrides))?
                } else {
                    parent_vars.clone()
                };
                Ok(vec![expand_value(Value::Mapping(mapping), &variables)?])
            } else {
                Err(DocxError::parse(
                    "included YAML fragment must be a block, a block sequence, or a mapping with `blocks`",
                ))
            }
        }
        _ => Err(DocxError::parse(
            "included YAML fragment must be a block, a block sequence, or a mapping with `blocks`",
        )),
    }
}

fn expand_repeat_block(
    block: Value,
    current_dir: Option<&Path>,
    variables: &Mapping,
    state: &mut ExpansionState,
) -> Result<Vec<Value>> {
    let mut mapping = expect_mapping(block, "`repeat` block must be a mapping")?;
    let alias = mapping
        .remove(string_key("as"))
        .map(|value| {
            let value = expand_value(value, variables)?;
            scalar_to_string(&value)
                .ok_or_else(|| DocxError::parse("`repeat.as` must resolve to a scalar string"))
        })
        .transpose()?
        .unwrap_or_else(|| "item".to_string());

    let template = mapping
        .remove(string_key("blocks"))
        .ok_or_else(|| DocxError::parse("`repeat` block requires a `blocks` field"))?;
    let template = expect_sequence(template, "`repeat.blocks` must be a sequence")?;

    let items = match (
        mapping.remove(string_key("items")),
        mapping.remove(string_key("variable")),
    ) {
        (Some(items), None) => expect_sequence(
            expand_value(items, variables)?,
            "`repeat.items` must resolve to a sequence",
        )?,
        (None, Some(variable)) => {
            let variable = expand_value(variable, variables)?;
            let variable_name = scalar_to_string(&variable).ok_or_else(|| {
                DocxError::parse("`repeat.variable` must resolve to a scalar string")
            })?;
            let item_values =
                lookup_variable_value(variables, &variable_name).ok_or_else(|| {
                    DocxError::parse(format!(
                        "unknown repeat variable '{variable_name}' in YAML document spec"
                    ))
                })?;
            expect_sequence(
                item_values.clone(),
                "`repeat.variable` must point to a sequence value",
            )?
        }
        (Some(_), Some(_)) => {
            return Err(DocxError::parse(
                "`repeat` block accepts either `items` or `variable`, not both",
            ))
        }
        (None, None) => {
            return Err(DocxError::parse(
                "`repeat` block requires either `items` or `variable`",
            ))
        }
    };

    let mut expanded = Vec::new();
    for (index, item) in items.into_iter().enumerate() {
        let item = expand_value(item, variables)?;
        let mut repeat_vars = variables.clone();
        repeat_vars.insert(Value::String(alias.clone()), item);
        repeat_vars.insert(
            string_key("repeat_index"),
            Value::Number(Number::from(index as u64)),
        );
        repeat_vars.insert(
            string_key("repeat_number"),
            Value::Number(Number::from((index + 1) as u64)),
        );
        expanded.extend(expand_block_sequence(
            Value::Sequence(template.clone()),
            current_dir,
            &repeat_vars,
            state,
        )?);
    }

    Ok(expanded)
}

fn expand_value(value: Value, variables: &Mapping) -> Result<Value> {
    match value {
        Value::String(text) => interpolate_string_value(&text, variables),
        Value::Sequence(items) => Ok(Value::Sequence(
            items
                .into_iter()
                .map(|value| expand_value(value, variables))
                .collect::<Result<Vec<_>>>()?,
        )),
        Value::Mapping(mapping) => {
            let mut expanded = Mapping::new();
            for (key, value) in mapping {
                expanded.insert(key, expand_value(value, variables)?);
            }
            Ok(Value::Mapping(expanded))
        }
        other => Ok(other),
    }
}

fn interpolate_string_value(text: &str, variables: &Mapping) -> Result<Value> {
    if let Some(path) = exact_placeholder(text) {
        let value = lookup_variable_value(variables, path).ok_or_else(|| {
            DocxError::parse(format!("unknown variable '{path}' in YAML document spec"))
        })?;
        return Ok(value.clone());
    }

    let mut rendered = String::new();
    let mut rest = text;
    while let Some(start) = rest.find("{{") {
        rendered.push_str(&rest[..start]);
        let tail = &rest[start + 2..];
        let Some(end) = tail.find("}}") else {
            return Err(DocxError::parse(format!(
                "unterminated variable placeholder in '{text}'"
            )));
        };
        let path = tail[..end].trim();
        if path.is_empty() {
            return Err(DocxError::parse(format!(
                "empty variable placeholder in '{text}'"
            )));
        }
        let value = lookup_variable_value(variables, path).ok_or_else(|| {
            DocxError::parse(format!("unknown variable '{path}' in YAML document spec"))
        })?;
        let scalar = scalar_to_string(value).ok_or_else(|| {
            DocxError::parse(format!(
                "variable '{path}' resolves to a non-scalar value and cannot be interpolated into text"
            ))
        })?;
        rendered.push_str(&scalar);
        rest = &tail[end + 2..];
    }
    rendered.push_str(rest);
    Ok(Value::String(rendered))
}

fn exact_placeholder(text: &str) -> Option<&str> {
    if !text.starts_with("{{") || !text.ends_with("}}") {
        return None;
    }
    let inner = &text[2..text.len() - 2];
    let trimmed = inner.trim();
    if trimmed.is_empty() || inner.contains("}}") || inner.contains("{{") {
        None
    } else {
        Some(trimmed)
    }
}

fn lookup_variable_value<'a>(variables: &'a Mapping, path: &str) -> Option<&'a Value> {
    let mut current = variables.get(string_key(path.split('.').next()?))?;
    for segment in path.split('.').skip(1) {
        current = match current {
            Value::Mapping(mapping) => mapping.get(string_key(segment))?,
            Value::Sequence(items) => items.get(segment.parse::<usize>().ok()?)?,
            _ => return None,
        };
    }
    Some(current)
}

fn scalar_to_string(value: &Value) -> Option<String> {
    match value {
        Value::Null => Some(String::new()),
        Value::Bool(value) => Some(value.to_string()),
        Value::Number(value) => Some(value.to_string()),
        Value::String(value) => Some(value.clone()),
        Value::Sequence(_) | Value::Mapping(_) | Value::Tagged(_) => None,
    }
}

fn block_type(value: &Value) -> Option<String> {
    let Value::Mapping(mapping) = value else {
        return None;
    };
    mapping
        .get(string_key("type"))
        .and_then(Value::as_str)
        .map(str::to_string)
}

fn resolve_include_path(current_dir: Option<&Path>, raw_path: &str) -> Result<PathBuf> {
    let path = PathBuf::from(raw_path);
    if path.is_absolute() {
        Ok(path)
    } else {
        let Some(current_dir) = current_dir else {
            return Err(DocxError::parse(format!(
                "cannot resolve relative include path '{raw_path}' without a source file"
            )));
        };
        Ok(current_dir.join(path))
    }
}

fn canonical_or_absolute(path: &Path) -> Result<PathBuf> {
    match fs::canonicalize(path) {
        Ok(path) => Ok(path),
        Err(_) if path.is_absolute() => Ok(path.to_path_buf()),
        Err(_) => Ok(std::env::current_dir()?.join(path)),
    }
}

fn expect_mapping(value: Value, message: &str) -> Result<Mapping> {
    match value {
        Value::Mapping(mapping) => Ok(mapping),
        _ => Err(DocxError::parse(message)),
    }
}

fn expect_sequence(value: Value, message: &str) -> Result<Vec<Value>> {
    match value {
        Value::Sequence(values) => Ok(values),
        _ => Err(DocxError::parse(message)),
    }
}

fn expect_string_key(key: &Value, message: &str) -> Result<String> {
    key.as_str()
        .map(str::to_string)
        .ok_or_else(|| DocxError::parse(message))
}

fn string_key(value: &str) -> Value {
    Value::String(value.to_string())
}

#[cfg(test)]
mod tests {
    use std::fs;

    use tempfile::tempdir;

    use super::expand_yaml_document_spec;

    #[test]
    fn expands_variables_repeaters_and_includes() {
        let temp = tempdir().expect("temp dir");
        let fragment_path = temp.path().join("fragment.yaml");
        fs::write(
            &fragment_path,
            r#"variables:
  intro: Included summary for {{client}}
blocks:
  - type: body
    text: "{{intro}}"
"#,
        )
        .expect("write fragment");
        let source_path = temp.path().join("spec.yaml");
        let yaml = format!(
            r#"output_name: regional-summary
variables:
  client: Acme
  regions:
    - name: North America
      owner: Maya
    - name: EMEA
      owner: Leon
blocks:
  - type: title
    text: "{{{{client}}}} Executive Summary"
  - type: include
    path: {}
  - type: repeat
    variable: regions
    as: region
    blocks:
      - type: section
        text: "{{{{region.name}}}}"
      - type: body
        text: "Owner: {{{{region.owner}}}}"
"#,
            fragment_path
                .file_name()
                .and_then(|name| name.to_str())
                .unwrap_or("fragment.yaml")
        );

        let expanded = expand_yaml_document_spec(&yaml, Some(&source_path)).expect("expand yaml");
        let spec: crate::spec::DocumentSpec =
            serde_yaml::from_value(expanded).expect("deserialize expanded spec");

        assert_eq!(spec.output_name.as_deref(), Some("regional-summary"));
        assert_eq!(spec.blocks.len(), 6);
        assert_eq!(
            serde_yaml::to_string(&spec.blocks[0]).expect("serialize block"),
            "type: title\ntext: Acme Executive Summary\n"
        );
        assert_eq!(
            serde_yaml::to_string(&spec.blocks[1]).expect("serialize block"),
            "type: body\ntext: Included summary for Acme\n"
        );
    }

    #[test]
    fn rejects_include_cycles() {
        let temp = tempdir().expect("temp dir");
        let a = temp.path().join("a.yaml");
        let b = temp.path().join("b.yaml");
        fs::write(
            &a,
            r#"blocks:
  - type: include
    path: b.yaml
"#,
        )
        .expect("write a");
        fs::write(
            &b,
            r#"blocks:
  - type: include
    path: a.yaml
"#,
        )
        .expect("write b");

        let error =
            expand_yaml_document_spec(&fs::read_to_string(&a).expect("read a"), Some(a.as_path()))
                .expect_err("cycle must fail");
        assert!(error.to_string().contains("include cycle"));
    }
}
