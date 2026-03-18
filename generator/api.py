"""
api.py — ISO Document Generator: Python Function API
=====================================================
Clean, importable API for programmatic use and AI tool integration.

Every function returns a consistent dict:
    {
        "success": bool,
        "message": str,
        "data":    <depends on function>
    }

Usage (Python):
    from generator.api import generate_documents, list_projects

    result = generate_documents("configs/LOGPRO/LOGPRO_config.json")
    if result["success"]:
        print(f"Generated {result['data']['count']} documents")

    projects = list_projects()
    for p in projects["data"]:
        print(p["code"], p["status"])
"""

import json
import os
import shutil
import subprocess
import sys
from pathlib import Path

# Resolve workspace root (this file lives at <root>/generator/api.py)
_API_DIR = Path(__file__).resolve().parent       # generator/
_WORKSPACE = _API_DIR.parent                      # DocISOGen/
_PYTHON = sys.executable


# ---------------------------------------------------------------------------
# Helper
# ---------------------------------------------------------------------------

def _ok(message: str, data=None) -> dict:
    return {"success": True, "message": message, "data": data}


def _err(message: str, data=None) -> dict:
    return {"success": False, "message": message, "data": data}


def _run(args: list[str], cwd: str = None) -> tuple[int, str]:
    """Run a subprocess, return (returncode, combined stdout+stderr)."""
    result = subprocess.run(
        args,
        capture_output=True,
        text=True,
        encoding="utf-8",
        errors="replace",
        cwd=cwd or str(_WORKSPACE),
    )
    combined = result.stdout + result.stderr
    return result.returncode, combined


# ---------------------------------------------------------------------------
# 1. generate_documents
# ---------------------------------------------------------------------------

def generate_documents(
    config_path: str,
    folder: str = None,
    validate: bool = True,
    verbose_validate: bool = False,
) -> dict:
    """
    Generate ISO documents for a project.

    Args:
        config_path:      Path to project config JSON (absolute or relative to workspace root).
        folder:           Optional — generate only this folder number, e.g. "05".
        validate:         Run QA validation after generation (default True).
        verbose_validate: Include INFO-level issues in QA report.

    Returns:
        {
            "success": bool,
            "message": str,
            "data": {
                "count": int,          # docs generated
                "total": int,          # docs attempted
                "output_dir": str,
                "output": str          # raw generator stdout
            }
        }

    Example:
        from generator.api import generate_documents
        result = generate_documents("configs/LOGPRO/LOGPRO_config.json", folder="05")
    """
    config_path = _resolve_path(config_path)
    if not Path(config_path).exists():
        return _err(f"Config file not found: {config_path}")

    args = [_PYTHON, str(_API_DIR / "generate_iso_docs.py"), "--config", config_path]
    if folder:
        args += ["--folder", str(folder).zfill(2)]
    if not validate:
        args.append("--no-validate")
    if verbose_validate:
        args.append("--verbose-validate")

    code, output = _run(args)

    # Parse "Complete: X/Y" from output
    count, total = _parse_count(output)

    # Determine output dir from config
    try:
        with open(config_path, encoding="utf-8") as f:
            cfg = json.load(f)
        project_name = cfg.get("project", {}).get("code") or cfg.get("project", {}).get("short_name", "Unknown")
        output_dir = str(_WORKSPACE / "projects" / project_name)
    except Exception:
        output_dir = ""

    if code == 0 or count > 0:
        return _ok(
            f"Generated {count}/{total} documents successfully.",
            {"count": count, "total": total, "output_dir": output_dir, "output": output},
        )
    return _err(
        f"Generation failed (exit code {code}).",
        {"count": 0, "total": total, "output_dir": output_dir, "output": output},
    )


# ---------------------------------------------------------------------------
# 2. validate_config
# ---------------------------------------------------------------------------

def validate_config(config_path: str, verbose: bool = False) -> dict:
    """
    Run QA validation on an existing project's documents.

    Args:
        config_path: Path to project config JSON.
        verbose:     Include INFO-level issues in report.

    Returns:
        {
            "success": bool,
            "message": str,
            "data": {"output": str}
        }

    Example:
        from generator.api import validate_config
        result = validate_config("configs/LOGPRO/LOGPRO_config.json")
    """
    config_path = _resolve_path(config_path)
    if not Path(config_path).exists():
        return _err(f"Config file not found: {config_path}")

    args = [_PYTHON, str(_API_DIR / "generate_iso_docs.py"),
            "--validate-only", "--config", config_path]
    if verbose:
        args.append("--verbose-validate")

    code, output = _run(args)
    success = code == 0
    msg = "Validation passed." if success else "Validation found issues."
    return ((_ok if success else _err)(msg, {"output": output}))


# ---------------------------------------------------------------------------
# 3. audit_snapshot
# ---------------------------------------------------------------------------

def audit_snapshot(project_code: str, report: bool = False, output_path: str = None) -> dict:
    """
    Extract audit-relevant data from project .docx files → compact JSON snapshot.

    Args:
        project_code: Project code matching folder name under projects/ (e.g. "LOGPRO").
        report:       Also print a human-readable summary to stdout.
        output_path:  Custom output path for snapshot JSON.

    Returns:
        {
            "success": bool,
            "message": str,
            "data": {
                "snapshot_path": str,
                "output": str,
                "summary": dict   # parsed snapshot summary if available
            }
        }

    Example:
        from generator.api import audit_snapshot
        result = audit_snapshot("LOGPRO", report=True)
        print(result["data"]["snapshot_path"])
    """
    args = [_PYTHON, str(_WORKSPACE / "tools" / "audit_snapshot.py"),
            "--project", project_code]
    if report:
        args.append("--report")
    if output_path:
        args += ["--output", output_path]

    code, output = _run(args)

    # Determine default snapshot path
    snapshot_path = str(_WORKSPACE / "configs" / project_code / f"{project_code}_snapshot.json")
    if output_path:
        snapshot_path = output_path

    # Try to load summary from snapshot
    summary = {}
    try:
        if Path(snapshot_path).exists():
            with open(snapshot_path, encoding="utf-8") as f:
                snap = json.load(f)
            summary = snap.get("summary", {})
    except Exception:
        pass

    if code == 0:
        return _ok(
            f"Snapshot saved: {snapshot_path}",
            {"snapshot_path": snapshot_path, "output": output, "summary": summary},
        )
    return _err(
        f"Snapshot failed (exit code {code}).",
        {"snapshot_path": snapshot_path, "output": output, "summary": summary},
    )


# ---------------------------------------------------------------------------
# 4. apply_patches
# ---------------------------------------------------------------------------

def apply_patches(
    project_code: str,
    patches_path: str,
    dry_run: bool = False,
    no_backup: bool = False,
) -> dict:
    """
    Apply AI-generated patches to project .docx files.

    Args:
        project_code: Project code (e.g. "LOGPRO").
        patches_path: Path to patches JSON file.
        dry_run:      Preview changes without modifying files.
        no_backup:    Skip creating .backup.docx files (not recommended).

    Returns:
        {
            "success": bool,
            "message": str,
            "data": {"output": str}
        }

    Example:
        from generator.api import apply_patches
        result = apply_patches("LOGPRO", "configs/LOGPRO/LOGPRO_patches.json", dry_run=True)
    """
    patches_path = _resolve_path(patches_path)
    if not Path(patches_path).exists():
        return _err(f"Patches file not found: {patches_path}")

    args = [_PYTHON, str(_WORKSPACE / "tools" / "doc_patcher.py"),
            "--project", project_code,
            "--patches", patches_path]
    if dry_run:
        args.append("--dry-run")
    if no_backup:
        args.append("--no-backup")

    code, output = _run(args)
    mode = "Dry run" if dry_run else "Patches applied"
    if code == 0:
        return _ok(f"{mode} completed.", {"output": output})
    return _err(f"{mode} failed (exit code {code}).", {"output": output})


# ---------------------------------------------------------------------------
# 5. list_projects
# ---------------------------------------------------------------------------

def list_projects(configs_dir: str = None) -> dict:
    """
    List all projects found in configs/ directory.

    Args:
        configs_dir: Override configs directory path.

    Returns:
        {
            "success": bool,
            "message": str,
            "data": [
                {
                    "code":        str,
                    "name":        str,
                    "config_path": str,
                    "has_docs":    bool,
                    "doc_count":   int,
                    "status":      str   # "ready" | "no_docs" | "config_only"
                },
                ...
            ]
        }

    Example:
        from generator.api import list_projects
        result = list_projects()
        for p in result["data"]:
            print(f"{p['code']:15} {p['name']:30} docs={p['doc_count']}")
    """
    base_configs = Path(configs_dir) if configs_dir else (_WORKSPACE / "configs")

    if not base_configs.exists():
        return _err(f"Configs directory not found: {base_configs}")

    projects = []
    for entry in sorted(base_configs.iterdir()):
        if not entry.is_dir():
            continue

        code = entry.name
        config_file = entry / f"{code}_config.json"

        if not config_file.exists():
            # Check for any *_config.json in the folder
            candidates = list(entry.glob("*_config.json"))
            if not candidates:
                continue
            config_file = candidates[0]

        # Load name from config
        name = code
        try:
            with open(config_file, encoding="utf-8") as f:
                cfg = json.load(f)
            name = cfg.get("project", {}).get("name", code)
        except Exception:
            pass

        # Count docs
        project_folder = _WORKSPACE / "projects" / code
        doc_count = 0
        has_docs = False
        if project_folder.exists():
            doc_count = len(list(project_folder.rglob("*.docx")))
            has_docs = doc_count > 0

        status = "ready" if has_docs else ("config_only" if config_file.exists() else "no_docs")

        projects.append({
            "code":        code,
            "name":        name,
            "config_path": str(config_file),
            "has_docs":    has_docs,
            "doc_count":   doc_count,
            "status":      status,
        })

    return _ok(f"Found {len(projects)} project(s).", projects)


# ---------------------------------------------------------------------------
# 6. create_config_scaffold
# ---------------------------------------------------------------------------

def create_config_scaffold(project_code: str, output_dir: str = None) -> dict:
    """
    Scaffold a new project config from the template.

    Args:
        project_code: New project code (e.g. "MYPROJ"). Used for folder name.
        output_dir:   Custom output directory. Defaults to configs/[project_code]/.

    Returns:
        {
            "success": bool,
            "message": str,
            "data": {"config_path": str}
        }

    Example:
        from generator.api import create_config_scaffold
        result = create_config_scaffold("MYPROJ")
        print(f"Edit: {result['data']['config_path']}")
    """
    template_path = _API_DIR / "config_template.json"
    if not template_path.exists():
        return _err("config_template.json not found in generator/")

    if output_dir:
        dest_dir = Path(output_dir)
    else:
        dest_dir = _WORKSPACE / "configs" / project_code

    dest_dir.mkdir(parents=True, exist_ok=True)
    dest_file = dest_dir / f"{project_code}_config.json"

    if dest_file.exists():
        return _err(
            f"Config already exists: {dest_file}",
            {"config_path": str(dest_file)},
        )

    # Copy template and inject project_code as initial value
    try:
        with open(template_path, encoding="utf-8") as f:
            template = json.load(f)

        # Pre-fill project code so user doesn't have to hunt for it
        template["project"]["code"] = project_code
        template["project"]["short_name"] = project_code
        template["output_path"] = str(_WORKSPACE / "projects")

        with open(dest_file, "w", encoding="utf-8") as f:
            json.dump(template, f, ensure_ascii=False, indent=2)

    except Exception as e:
        return _err(f"Failed to create config: {e}")

    return _ok(
        f"Config scaffolded at: {dest_file}\nEdit the file and run: python isogen.py generate --config \"{dest_file}\"",
        {"config_path": str(dest_file)},
    )


# ---------------------------------------------------------------------------
# 7. run_demo
# ---------------------------------------------------------------------------

def run_demo() -> dict:
    """
    Run generation with built-in demo config (no config file needed).

    Returns:
        {
            "success": bool,
            "message": str,
            "data": {"count": int, "output_dir": str, "output": str}
        }

    Example:
        from generator.api import run_demo
        result = run_demo()
    """
    args = [_PYTHON, str(_API_DIR / "generate_iso_docs.py"), "--demo"]
    code, output = _run(args)
    count, total = _parse_count(output)
    output_dir = str(_WORKSPACE / "projects" / "HRMDemo")

    if code == 0 or count > 0:
        return _ok(
            f"Demo generated {count}/{total} documents.",
            {"count": count, "total": total, "output_dir": output_dir, "output": output},
        )
    return _err("Demo generation failed.", {"output": output})


# ---------------------------------------------------------------------------
# Internal helpers
# ---------------------------------------------------------------------------

def _resolve_path(path: str) -> str:
    """Resolve path relative to workspace root if not absolute."""
    p = Path(path)
    if p.is_absolute():
        return str(p)
    resolved = _WORKSPACE / path
    return str(resolved)


def _parse_count(output: str) -> tuple[int, int]:
    """Parse 'Complete: X/Y' from generator output."""
    import re
    m = re.search(r"Complete:\s*(\d+)/(\d+)", output)
    if m:
        return int(m.group(1)), int(m.group(2))
    return 0, 21
