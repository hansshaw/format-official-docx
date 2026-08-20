#!/usr/bin/env python3
"""Small repository-local validator for the public skill package."""

from __future__ import annotations

import re
import sys
from pathlib import Path

import yaml


ALLOWED_FRONTMATTER_KEYS = {
    "name",
    "description",
    "license",
    "allowed-tools",
    "metadata",
}


def validate(root: Path) -> list[str]:
    errors: list[str] = []
    skill = root / "SKILL.md"
    if not skill.is_file():
        return ["SKILL.md is missing"]
    text = skill.read_text(encoding="utf-8")
    match = re.match(r"^---\r?\n(.*?)\r?\n---", text, re.DOTALL)
    if not match:
        return ["SKILL.md frontmatter is missing or malformed"]
    try:
        frontmatter = yaml.safe_load(match.group(1))
    except yaml.YAMLError as exc:
        return [f"SKILL.md frontmatter is invalid YAML: {exc}"]
    if not isinstance(frontmatter, dict):
        return ["SKILL.md frontmatter must be a mapping"]

    unexpected = set(frontmatter) - ALLOWED_FRONTMATTER_KEYS
    if unexpected:
        errors.append("unexpected SKILL.md frontmatter keys: " + ", ".join(sorted(unexpected)))
    name = frontmatter.get("name")
    if name != "format-official-docx":
        errors.append("SKILL.md name must be format-official-docx")
    if not isinstance(name, str) or not re.fullmatch(r"[a-z0-9]+(?:-[a-z0-9]+)*", name):
        errors.append("SKILL.md name must use lowercase hyphen-case")
    description = frontmatter.get("description")
    if not isinstance(description, str) or not description.strip():
        errors.append("SKILL.md description is missing")
    elif len(description) > 1024:
        errors.append("SKILL.md description exceeds 1024 characters")
    for relative in (
        "agents/openai.yaml",
        "references/format-spec.md",
        "scripts/official_docx.py",
        "scripts/verify_official_docx.py",
        "README.md",
        "LICENSE",
        "requirements.txt",
    ):
        if not (root / relative).is_file():
            errors.append(f"required file is missing: {relative}")
    if list(root.rglob("*.pyc")):
        errors.append("compiled Python files must not be committed")
    for yaml_path in sorted([*root.rglob("*.yml"), *root.rglob("*.yaml")]):
        try:
            yaml.safe_load(yaml_path.read_text(encoding="utf-8"))
        except (OSError, UnicodeError, yaml.YAMLError) as exc:
            errors.append(f"invalid YAML file {yaml_path.relative_to(root)}: {exc}")
    return errors


def main() -> int:
    root = Path(sys.argv[1] if len(sys.argv) > 1 else ".").resolve()
    errors = validate(root)
    if errors:
        for error in errors:
            print(f"ERROR: {error}")
        return 1
    print(f"Skill package is valid: {root}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
