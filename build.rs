use std::env;
use std::process::Command;

fn git_short_hash() -> Option<String> {
    let output = Command::new("git")
        .args(["rev-parse", "--short", "HEAD"])
        .output()
        .ok()?;
    if !output.status.success() {
        return None;
    }

    let hash = String::from_utf8(output.stdout).ok()?;
    let hash = hash.trim();
    if hash.is_empty() {
        return None;
    }

    Some(hash.to_string())
}

fn format_version_display(pkg_version: &str, git_hash: Option<&str>) -> String {
    match git_hash.map(str::trim) {
        Some(hash) if !hash.is_empty() => format!("{pkg_version} ({hash})"),
        _ => pkg_version.to_string(),
    }
}

fn main() {
    let pkg_version = env::var("CARGO_PKG_VERSION").unwrap_or_else(|_| "0.0.0".to_string());
    let display = format_version_display(&pkg_version, git_short_hash().as_deref());

    println!("cargo:rustc-env=ZEROCLAW_VERSION_DISPLAY={display}");
    // Recompute version display when git HEAD (or refs under it) changes.
    println!("cargo:rerun-if-changed=.git/HEAD");
    println!("cargo:rerun-if-changed=.git/refs");
}
