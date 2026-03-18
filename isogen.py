"""
isogen.py — ISO Document Generator: Unified CLI
================================================
One entry point for all document generation and audit workflows.

Usage:
    python isogen.py <command> [options]

Commands:
    generate   Generate ISO documents for a project
    validate   Run QA validation on generated documents
    audit      Create audit snapshot from existing .docx files
    patch      Apply AI-generated patches to documents
    new        Scaffold a new project config
    list       List all known projects
    demo       Quick-start with built-in demo project

Run 'python isogen.py <command> --help' for per-command help.
"""

import argparse
import sys
from pathlib import Path

# Ensure generator/ is importable
_ROOT = Path(__file__).resolve().parent
sys.path.insert(0, str(_ROOT))

from generator.api import (
    apply_patches,
    audit_snapshot,
    create_config_scaffold,
    generate_documents,
    list_projects,
    run_demo,
    validate_config,
)


# ---------------------------------------------------------------------------
# Output helpers
# ---------------------------------------------------------------------------

_GREEN = "\033[92m"
_RED   = "\033[91m"
_CYAN  = "\033[96m"
_BOLD  = "\033[1m"
_RESET = "\033[0m"


def _ok(msg: str):
    print(f"{_GREEN}✅ {msg}{_RESET}")


def _fail(msg: str):
    print(f"{_RED}❌ {msg}{_RESET}", file=sys.stderr)


def _info(msg: str):
    print(f"{_CYAN}{msg}{_RESET}")


def _print_result(result: dict, raw: bool = False):
    """Print a consistent result dict, showing raw output only if requested."""
    if result["success"]:
        _ok(result["message"])
    else:
        _fail(result["message"])

    data = result.get("data") or {}
    if raw and isinstance(data, dict) and "output" in data:
        print()
        print(data["output"].strip())
    elif isinstance(data, dict) and "output" in data and not result["success"]:
        # Always show output on failure
        print()
        print(data["output"].strip())


# ---------------------------------------------------------------------------
# Command: generate
# ---------------------------------------------------------------------------

def cmd_generate(args):
    """
    Generate ISO documents for a project.

    Examples:
        python isogen.py generate --config configs/LOGPRO/LOGPRO_config.json
        python isogen.py generate --config configs/LOGPRO/LOGPRO_config.json --folder 05
        python isogen.py generate --config configs/LOGPRO/LOGPRO_config.json --no-validate
    """
    _info(f"Generating documents from: {args.config}")
    if args.folder:
        _info(f"  Folder: {args.folder}")

    result = generate_documents(
        config_path=args.config,
        folder=args.folder,
        validate=not args.no_validate,
        verbose_validate=args.verbose_validate,
    )

    _print_result(result, raw=True)

    data = result.get("data") or {}
    if result["success"] and data.get("output_dir"):
        _info(f"\n  Output: {data['output_dir']}")

    sys.exit(0 if result["success"] else 1)


# ---------------------------------------------------------------------------
# Command: validate
# ---------------------------------------------------------------------------

def cmd_validate(args):
    """
    Validate generated documents for QA compliance.

    Examples:
        python isogen.py validate --config configs/LOGPRO/LOGPRO_config.json
        python isogen.py validate --config configs/LOGPRO/LOGPRO_config.json --verbose
    """
    _info(f"Validating: {args.config}")
    result = validate_config(args.config, verbose=args.verbose)
    _print_result(result, raw=True)
    sys.exit(0 if result["success"] else 1)


# ---------------------------------------------------------------------------
# Command: audit
# ---------------------------------------------------------------------------

def cmd_audit(args):
    """
    Create audit snapshot from existing .docx files.

    Examples:
        python isogen.py audit --project LOGPRO
        python isogen.py audit --project LOGPRO --report
    """
    _info(f"Auditing project: {args.project}")
    result = audit_snapshot(
        project_code=args.project,
        report=args.report,
        output_path=args.output or None,
    )

    _print_result(result, raw=True)

    data = result.get("data") or {}
    if result["success"] and data.get("snapshot_path"):
        _info(f"\n  Snapshot: {data['snapshot_path']}")
        if data.get("summary"):
            s = data["summary"]
            _info(f"  Documents: {s.get('total_docs', '?')}")
            _info(f"  Placeholders: {s.get('total_placeholders', '?')}")

    sys.exit(0 if result["success"] else 1)


# ---------------------------------------------------------------------------
# Command: patch
# ---------------------------------------------------------------------------

def cmd_patch(args):
    """
    Apply AI-generated patches to project documents.

    Examples:
        python isogen.py patch --project LOGPRO --patches configs/LOGPRO/LOGPRO_patches.json --dry-run
        python isogen.py patch --project LOGPRO --patches configs/LOGPRO/LOGPRO_patches.json
    """
    mode = "[DRY RUN] " if args.dry_run else ""
    _info(f"{mode}Patching project: {args.project}")
    _info(f"  Patches: {args.patches}")

    result = apply_patches(
        project_code=args.project,
        patches_path=args.patches,
        dry_run=args.dry_run,
        no_backup=args.no_backup,
    )

    _print_result(result, raw=True)
    sys.exit(0 if result["success"] else 1)


# ---------------------------------------------------------------------------
# Command: new
# ---------------------------------------------------------------------------

def cmd_new(args):
    """
    Scaffold a new project config from the template.

    Examples:
        python isogen.py new --code MYPROJ
        python isogen.py new --code MYPROJ --out configs/MYPROJ/MYPROJ_config.json
    """
    _info(f"Scaffolding config for: {args.code}")
    result = create_config_scaffold(
        project_code=args.code,
        output_dir=args.out or None,
    )

    _print_result(result)

    if result["success"]:
        data = result.get("data") or {}
        config_path = data.get("config_path", "")
        print()
        _info("Next steps:")
        _info(f"  1. Edit config:      {config_path}")
        _info(f"  2. Generate docs:    python isogen.py generate --config \"{config_path}\"")

    sys.exit(0 if result["success"] else 1)


# ---------------------------------------------------------------------------
# Command: list
# ---------------------------------------------------------------------------

def cmd_list(args):
    """
    List all projects with their document counts.

    Examples:
        python isogen.py list
    """
    result = list_projects()

    if not result["success"]:
        _fail(result["message"])
        sys.exit(1)

    projects = result["data"]
    if not projects:
        _info("No projects found in configs/")
        sys.exit(0)

    # Header
    print()
    print(f"  {'CODE':<18} {'NAME':<32} {'DOCS':>5}  STATUS")
    print(f"  {'-'*18} {'-'*32} {'-'*5}  {'-'*10}")

    for p in projects:
        status_icon = "✅" if p["has_docs"] else "⬜"
        doc_str = str(p["doc_count"]) if p["has_docs"] else "—"
        print(f"  {p['code']:<18} {p['name']:<32} {doc_str:>5}  {status_icon} {p['status']}")

    print()
    _info(f"Total: {len(projects)} project(s)")
    sys.exit(0)


# ---------------------------------------------------------------------------
# Command: demo
# ---------------------------------------------------------------------------

def cmd_demo(args):
    """
    Generate all documents using the built-in demo project (no config needed).

    Examples:
        python isogen.py demo
    """
    _info("Running demo generation...")
    result = run_demo()
    _print_result(result, raw=True)

    if result["success"]:
        data = result.get("data") or {}
        _info(f"\n  Output: {data.get('output_dir', '')}")

    sys.exit(0 if result["success"] else 1)


# ---------------------------------------------------------------------------
# CLI parser
# ---------------------------------------------------------------------------

def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        prog="isogen",
        description=(
            "ISO Document Generator — CLI\n"
            "Generate, validate, audit, and patch ISO/IEC 29110 document sets."
        ),
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog=(
            "Examples:\n"
            "  python isogen.py list\n"
            "  python isogen.py generate --config configs/LOGPRO/LOGPRO_config.json\n"
            "  python isogen.py generate --config configs/LOGPRO/LOGPRO_config.json --folder 05\n"
            "  python isogen.py validate --config configs/LOGPRO/LOGPRO_config.json\n"
            "  python isogen.py audit    --project LOGPRO --report\n"
            "  python isogen.py patch    --project LOGPRO --patches configs/LOGPRO/LOGPRO_patches.json --dry-run\n"
            "  python isogen.py new      --code MYPROJ\n"
            "  python isogen.py demo\n"
        ),
    )
    sub = parser.add_subparsers(dest="command", metavar="command")
    sub.required = True

    # ── generate ──────────────────────────────────────────────────────────
    p_gen = sub.add_parser("generate", help="Generate ISO documents", description=cmd_generate.__doc__,
                           formatter_class=argparse.RawDescriptionHelpFormatter)
    p_gen.add_argument("--config", required=True,
                       help="Path to project config JSON (e.g. configs/LOGPRO/LOGPRO_config.json)")
    p_gen.add_argument("--folder", default=None, metavar="NUM",
                       help="Generate only this folder (01-10). Default: all folders")
    p_gen.add_argument("--no-validate", action="store_true",
                       help="Skip QA validation after generation")
    p_gen.add_argument("--verbose-validate", action="store_true",
                       help="Show INFO-level issues in QA report")
    p_gen.set_defaults(func=cmd_generate)

    # ── validate ──────────────────────────────────────────────────────────
    p_val = sub.add_parser("validate", help="Validate existing documents (QA)",
                            description=cmd_validate.__doc__,
                            formatter_class=argparse.RawDescriptionHelpFormatter)
    p_val.add_argument("--config", required=True,
                       help="Path to project config JSON")
    p_val.add_argument("--verbose", action="store_true",
                       help="Show INFO-level issues in report")
    p_val.set_defaults(func=cmd_validate)

    # ── audit ─────────────────────────────────────────────────────────────
    p_aud = sub.add_parser("audit", help="Create audit snapshot from .docx files",
                            description=cmd_audit.__doc__,
                            formatter_class=argparse.RawDescriptionHelpFormatter)
    p_aud.add_argument("--project", required=True,
                       help="Project code (e.g. LOGPRO)")
    p_aud.add_argument("--report", action="store_true",
                       help="Print human-readable summary report")
    p_aud.add_argument("--output", default=None, metavar="PATH",
                       help="Custom output path for snapshot JSON")
    p_aud.set_defaults(func=cmd_audit)

    # ── patch ─────────────────────────────────────────────────────────────
    p_patch = sub.add_parser("patch", help="Apply patches to documents",
                              description=cmd_patch.__doc__,
                              formatter_class=argparse.RawDescriptionHelpFormatter)
    p_patch.add_argument("--project", required=True,
                         help="Project code (e.g. LOGPRO)")
    p_patch.add_argument("--patches", required=True, metavar="PATH",
                         help="Path to patches JSON file")
    p_patch.add_argument("--dry-run", action="store_true",
                         help="Preview changes without modifying files")
    p_patch.add_argument("--no-backup", action="store_true",
                         help="Skip .backup.docx (not recommended)")
    p_patch.set_defaults(func=cmd_patch)

    # ── new ───────────────────────────────────────────────────────────────
    p_new = sub.add_parser("new", help="Scaffold a new project config",
                            description=cmd_new.__doc__,
                            formatter_class=argparse.RawDescriptionHelpFormatter)
    p_new.add_argument("--code", required=True,
                       help="New project code (e.g. MYPROJ)")
    p_new.add_argument("--out", default=None, metavar="DIR",
                       help="Output directory. Default: configs/[CODE]/")
    p_new.set_defaults(func=cmd_new)

    # ── list ──────────────────────────────────────────────────────────────
    p_list = sub.add_parser("list", help="List all projects",
                             description=cmd_list.__doc__,
                             formatter_class=argparse.RawDescriptionHelpFormatter)
    p_list.set_defaults(func=cmd_list)

    # ── demo ──────────────────────────────────────────────────────────────
    p_demo = sub.add_parser("demo", help="Generate documents with built-in demo config",
                             description=cmd_demo.__doc__,
                             formatter_class=argparse.RawDescriptionHelpFormatter)
    p_demo.set_defaults(func=cmd_demo)

    return parser


# ---------------------------------------------------------------------------
# Entry point
# ---------------------------------------------------------------------------

def main():
    # Enable ANSI colors on Windows terminals
    if sys.platform == "win32":
        import os
        os.system("")  # Activates ANSI processing in Windows console

    parser = build_parser()
    args = parser.parse_args()
    args.func(args)


if __name__ == "__main__":
    main()
