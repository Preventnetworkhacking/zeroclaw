use super::traits::{Tool, ToolResult};
use crate::security::SecurityPolicy;
use async_trait::async_trait;
use serde_json::json;
use std::io::{Cursor, Read};
use std::sync::Arc;

/// Maximum PPTX file size (50 MB).
const MAX_PPTX_BYTES: u64 = 50 * 1024 * 1024;
/// Default character limit returned to the LLM.
const DEFAULT_MAX_CHARS: usize = 50_000;
/// Hard ceiling regardless of what the caller requests.
const MAX_OUTPUT_CHARS: usize = 200_000;

/// Extract plain text from a PowerPoint (PPTX) file in the workspace.
///
/// PPTX files are ZIP archives containing XML. This tool extracts text
/// from all slides by parsing the `ppt/slides/slide*.xml` files.
pub struct PptxReadTool {
    security: Arc<SecurityPolicy>,
}

impl PptxReadTool {
    pub fn new(security: Arc<SecurityPolicy>) -> Self {
        Self { security }
    }
}

#[async_trait]
impl Tool for PptxReadTool {
    fn name(&self) -> &str {
        "pptx_read"
    }

    fn description(&self) -> &str {
        "Extract plain text from a PowerPoint (PPTX) file in the workspace. \
         Returns all readable text from all slides, separated by slide markers. \
         Useful for analyzing presentations without manual copy-paste."
    }

    fn parameters_schema(&self) -> serde_json::Value {
        json!({
            "type": "object",
            "properties": {
                "path": {
                    "type": "string",
                    "description": "Path to the PPTX file. Relative paths resolve from workspace; outside paths require policy allowlist."
                },
                "max_chars": {
                    "type": "integer",
                    "description": "Maximum characters to return (default: 50000, max: 200000)",
                    "minimum": 1,
                    "maximum": 200_000
                }
            },
            "required": ["path"]
        })
    }

    async fn execute(&self, args: serde_json::Value) -> anyhow::Result<ToolResult> {
        let path = args
            .get("path")
            .and_then(|v| v.as_str())
            .ok_or_else(|| anyhow::anyhow!("Missing 'path' parameter"))?;

        let max_chars = args
            .get("max_chars")
            .and_then(|v| v.as_u64())
            .map(|n| {
                usize::try_from(n)
                    .unwrap_or(MAX_OUTPUT_CHARS)
                    .min(MAX_OUTPUT_CHARS)
            })
            .unwrap_or(DEFAULT_MAX_CHARS);

        if self.security.is_rate_limited() {
            return Ok(ToolResult {
                success: false,
                output: String::new(),
                error: Some("Rate limit exceeded: too many actions in the last hour".into()),
            });
        }

        if !self.security.is_path_allowed(path) {
            return Ok(ToolResult {
                success: false,
                output: String::new(),
                error: Some(format!("Path not allowed by security policy: {path}")),
            });
        }

        if !self.security.record_action() {
            return Ok(ToolResult {
                success: false,
                output: String::new(),
                error: Some("Rate limit exceeded: action budget exhausted".into()),
            });
        }

        let full_path = self.security.workspace_dir.join(path);

        let resolved_path = match tokio::fs::canonicalize(&full_path).await {
            Ok(p) => p,
            Err(e) => {
                return Ok(ToolResult {
                    success: false,
                    output: String::new(),
                    error: Some(format!("Failed to resolve file path: {e}")),
                });
            }
        };

        if !self.security.is_resolved_path_allowed(&resolved_path) {
            return Ok(ToolResult {
                success: false,
                output: String::new(),
                error: Some(
                    self.security
                        .resolved_path_violation_message(&resolved_path),
                ),
            });
        }

        tracing::debug!("Reading PPTX: {}", resolved_path.display());

        match tokio::fs::metadata(&resolved_path).await {
            Ok(meta) => {
                if meta.len() > MAX_PPTX_BYTES {
                    return Ok(ToolResult {
                        success: false,
                        output: String::new(),
                        error: Some(format!(
                            "PPTX too large: {} bytes (limit: {MAX_PPTX_BYTES} bytes)",
                            meta.len()
                        )),
                    });
                }
            }
            Err(e) => {
                return Ok(ToolResult {
                    success: false,
                    output: String::new(),
                    error: Some(format!("Failed to read file metadata: {e}")),
                });
            }
        }

        let bytes = match tokio::fs::read(&resolved_path).await {
            Ok(b) => b,
            Err(e) => {
                return Ok(ToolResult {
                    success: false,
                    output: String::new(),
                    error: Some(format!("Failed to read PPTX file: {e}")),
                });
            }
        };

        // PPTX extraction is CPU-bound; run in blocking task
        let text = match tokio::task::spawn_blocking(move || extract_pptx_text(&bytes)).await {
            Ok(Ok(t)) => t,
            Ok(Err(e)) => {
                return Ok(ToolResult {
                    success: false,
                    output: String::new(),
                    error: Some(format!("PPTX extraction failed: {e}")),
                });
            }
            Err(e) => {
                return Ok(ToolResult {
                    success: false,
                    output: String::new(),
                    error: Some(format!("PPTX extraction task panicked: {e}")),
                });
            }
        };

        if text.trim().is_empty() {
            return Ok(ToolResult {
                success: true,
                output: "PPTX contains no extractable text (may be image-only)".into(),
                error: None,
            });
        }

        let output = if text.chars().count() > max_chars {
            let mut truncated: String = text.chars().take(max_chars).collect();
            truncated.push_str("\n\n[... truncated, use max_chars to read more ...]");
            truncated
        } else {
            text
        };

        Ok(ToolResult {
            success: true,
            output,
            error: None,
        })
    }
}

/// Extract text from PPTX bytes by parsing slide XML files.
fn extract_pptx_text(bytes: &[u8]) -> anyhow::Result<String> {
    let cursor = Cursor::new(bytes);
    let mut archive = zip::ZipArchive::new(cursor)?;

    // Collect slide files and sort them numerically
    let mut slide_files: Vec<String> = archive
        .file_names()
        .filter(|name| {
            name.starts_with("ppt/slides/slide") && name.ends_with(".xml")
        })
        .map(String::from)
        .collect();

    // Sort by slide number (slide1.xml, slide2.xml, etc.)
    slide_files.sort_by(|a, b| {
        let num_a = extract_slide_number(a).unwrap_or(0);
        let num_b = extract_slide_number(b).unwrap_or(0);
        num_a.cmp(&num_b)
    });

    let mut result = String::new();

    for (idx, slide_name) in slide_files.iter().enumerate() {
        let mut file = archive.by_name(slide_name)?;
        let mut xml_content = String::new();
        file.read_to_string(&mut xml_content)?;

        let slide_text = extract_text_from_xml(&xml_content);

        if !slide_text.trim().is_empty() {
            result.push_str(&format!("--- Slide {} ---\n", idx + 1));
            result.push_str(&slide_text);
            result.push_str("\n\n");
        }
    }

    Ok(result)
}

/// Extract slide number from filename like "ppt/slides/slide1.xml"
fn extract_slide_number(name: &str) -> Option<u32> {
    let filename = name.rsplit('/').next()?;
    let num_str = filename
        .strip_prefix("slide")?
        .strip_suffix(".xml")?;
    num_str.parse().ok()
}

/// Extract text content from OOXML by finding <a:t> elements.
/// This is a simple regex-based extraction that handles most cases.
fn extract_text_from_xml(xml: &str) -> String {
    let mut text = String::new();
    let mut in_text_element = false;
    let mut current_text = String::new();

    let mut chars = xml.chars().peekable();

    while let Some(c) = chars.next() {
        if c == '<' {
            // Check if this is <a:t> or </a:t>
            let mut tag = String::new();
            while let Some(&next_c) = chars.peek() {
                if next_c == '>' {
                    chars.next();
                    break;
                }
                tag.push(chars.next().unwrap());
            }

            if tag == "a:t" || tag.starts_with("a:t ") {
                in_text_element = true;
                current_text.clear();
            } else if tag == "/a:t" {
                if in_text_element && !current_text.is_empty() {
                    text.push_str(&current_text);
                    text.push(' ');
                }
                in_text_element = false;
            } else if tag == "/a:p" || tag == "a:br" || tag == "a:br/" {
                // Paragraph break or line break
                if !text.is_empty() && !text.ends_with('\n') {
                    text.push('\n');
                }
            }
        } else if in_text_element {
            current_text.push(c);
        }
    }

    // Decode common XML entities
    text.replace("&amp;", "&")
        .replace("&lt;", "<")
        .replace("&gt;", ">")
        .replace("&quot;", "\"")
        .replace("&apos;", "'")
        .replace("&#39;", "'")
}
