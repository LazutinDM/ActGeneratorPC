"""Audit and optionally add docstrings to ActGeneratorPC functions.

Run without arguments in CI to fail when a new function has no explanation.
Run with ``--fix`` after adding a function to insert a concise starter docstring;
then refine that text when the function has non-obvious business rules.
"""

from __future__ import annotations

import argparse
import ast
import re
from pathlib import Path


PROJECT_DIR = Path(__file__).resolve().parents[1]
TARGETS = (
    "app.py",
    "app_paths.py",
    "docx_preview.py",
    "equipment_registry_update.py",
    "excel_export.py",
    "history_store.py",
    "update_manager.py",
)

# COLLEAGUE EDIT POINT: use explicit wording for business-critical functions.
EXPLICIT_DESCRIPTIONS = {
    "ensure_dirs": "Create required portable folders and migrate supported legacy data.",
    "load_equipment_registry": "Load validated serial-to-station and model bindings from JSON.",
    "replace_placeholders_docx": "Replace placeholders throughout paragraphs, tables, headers and footers.",
    "create_surname_only_act": "Create the selected blank act template containing only the executor name.",
    "generate_act_document": "Fill one DOCX template and return its final output path.",
    "create_batch_acts": "Generate every queued act without opening individual previews.",
    "create_act": "Validate the form, optionally preview the act, and save the confirmed document.",
    "refresh_serials_for_station": "Filter factory numbers by the selected station and equipment model.",
    "apply_equipment_for_serial": "Fill station and equipment after an exact factory-number match.",
    "update_equipment_registry_from_tables": "Refresh equipment bindings from the approved remote workbook.",
    "update_registry": "Download, validate, back up and atomically install the equipment registry.",
    "extract_registry": "Extract approved station, model and serial bindings from the XLSX workbook.",
    "append_act_to_workbook": "Append one act to the configured Excel workbook while preserving ignored columns.",
    "render_docx_to_pdf": "Render a DOCX to PDF using Word first and LibreOffice as a fallback.",
    "safe_extract": "Extract an update archive while rejecting traversal and unsafe paths.",
    "confirm_healthy_startup": "Confirm a successful launch so a pending update is not rolled back.",
    "main": "Initialize the application, update service and main window event loop.",
}


def _human_name(name: str) -> str:
    """Convert a Python identifier into short readable words."""
    return re.sub(r"_+", " ", name.strip("_")).strip() or "callback"


def description_for(name: str, owner: str | None) -> str:
    """Return a concise starter explanation for one undocumented function."""
    if name in EXPLICIT_DESCRIPTIONS:
        return EXPLICIT_DESCRIPTIONS[name]
    readable = _human_name(name)
    if name == "__init__":
        return f"Initialize the {owner or 'object'} and its runtime state."
    prefixes = {
        "load": "Load",
        "save": "Save",
        "read": "Read",
        "write": "Write",
        "build": "Build",
        "create": "Create",
        "generate": "Generate",
        "update": "Update",
        "refresh": "Refresh",
        "apply": "Apply",
        "validate": "Validate",
        "normalize": "Normalize",
        "parse": "Parse",
        "resolve": "Resolve",
        "ensure": "Ensure",
        "make": "Create",
        "open": "Open",
        "show": "Show",
        "choose": "Choose",
        "set": "Set",
        "get": "Return",
        "selected": "Return",
        "current": "Return",
        "clear": "Clear",
        "delete": "Delete",
        "remove": "Remove",
        "add": "Add",
        "edit": "Edit",
        "filter": "Filter",
        "toggle": "Toggle",
        "check": "Check",
        "run": "Run",
    }
    first = readable.split()[0]
    verb = prefixes.get(first)
    if verb:
        remainder = readable[len(first):].strip()
        return f"{verb} {remainder or 'the related value'} for this workflow."
    if name.startswith("_on_") or name.startswith("event"):
        return f"Handle the {readable} interface event."
    scope = f" in {owner}" if owner else ""
    return f"Handle {readable}{scope}."


class FunctionCollector(ast.NodeVisitor):
    """Collect undocumented functions together with their owning class."""

    def __init__(self):
        """Initialize the collector stack and result list."""
        self.class_stack: list[str] = []
        self.missing: list[tuple[ast.FunctionDef | ast.AsyncFunctionDef, str | None]] = []

    def visit_ClassDef(self, node: ast.ClassDef):
        """Track the class name while visiting its methods and nested callbacks."""
        self.class_stack.append(node.name)
        self.generic_visit(node)
        self.class_stack.pop()

    def visit_FunctionDef(self, node: ast.FunctionDef):
        """Record a synchronous function when it has no docstring."""
        if ast.get_docstring(node) is None:
            self.missing.append((node, self.class_stack[-1] if self.class_stack else None))
        self.generic_visit(node)

    def visit_AsyncFunctionDef(self, node: ast.AsyncFunctionDef):
        """Record an asynchronous function when it has no docstring."""
        if ast.get_docstring(node) is None:
            self.missing.append((node, self.class_stack[-1] if self.class_stack else None))
        self.generic_visit(node)


def undocumented_functions(path: Path):
    """Return every function without a docstring in one Python source file."""
    source = path.read_text(encoding="utf-8-sig")
    tree = ast.parse(source, filename=str(path))
    collector = FunctionCollector()
    collector.visit(tree)
    return source, collector.missing


def add_docstrings(path: Path) -> int:
    """Insert starter docstrings without changing executable statements."""
    source, missing = undocumented_functions(path)
    if not missing:
        return 0
    lines = source.splitlines(keepends=True)
    insertions = []
    for node, owner in missing:
        if not node.body:
            continue
        first_statement = node.body[0]
        indent = " " * first_statement.col_offset
        description = description_for(node.name, owner).replace('"""', "'''")
        insertions.append((first_statement.lineno - 1, f'{indent}"""{description}"""\n'))
    for index, text in sorted(insertions, reverse=True):
        lines.insert(index, text)
    path.write_text("".join(lines), encoding="utf-8")
    return len(insertions)


def main() -> int:
    """Audit target modules or insert missing starter documentation."""
    parser = argparse.ArgumentParser()
    parser.add_argument("--fix", action="store_true")
    args = parser.parse_args()
    missing_total = 0
    for relative_path in TARGETS:
        path = PROJECT_DIR / relative_path
        if args.fix:
            inserted = add_docstrings(path)
            print(f"{relative_path}: inserted {inserted}")
        _source, missing = undocumented_functions(path)
        if missing:
            missing_total += len(missing)
            print(f"{relative_path}: missing {len(missing)}")
            for node, owner in missing:
                print(f"  {owner + '.' if owner else ''}{node.name}:{node.lineno}")
    return 1 if missing_total else 0


if __name__ == "__main__":
    raise SystemExit(main())
