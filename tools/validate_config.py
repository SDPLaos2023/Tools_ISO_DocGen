#!/usr/bin/env python3
"""
validate_config.py — Scan project config files for mockup/placeholder person names.

Usage:
    python tools/validate_config.py --config configs/LOGPRO/LOGPRO_config.json
    python tools/validate_config.py --all                   # scan all projects
    python tools/validate_config.py --all --fix             # interactive fix mode
    python tools/validate_config.py --config <path> --fix   # fix one project
"""
import argparse
import json
import re
import sys
from pathlib import Path

# ════════════════════════════════════════════════════════════════════════════════
# Detection Rules
# ════════════════════════════════════════════════════════════════════════════════

# Exact strings that are clearly placeholders (matched case-insensitively)
PLACEHOLDER_EXACT = {
    "tbd", "n/a", "na", "none", "unknown", "-", "—", "–", "null",
    "tba", "todo", "pending", "your name", "full name", "first last",
    "firstname lastname", "enter name", "name here", "name",
    "ชื่อ-สกุล", "ชื่อ นามสกุล", "ชื่อและนามสกุล",
    "placeholder", "sample", "template",
}

# Words that — if a name ENDS with one — indicate it is a group/dept, not a person
TRAILING_GROUP_WORDS = {
    "team", "group", "dept", "department", "squad", "unit",
    "division", "committee", "office",
}

# Words that are purely role/function titles — if the whole name is only these words
# it is a role label, not a person name
PURE_ROLE_WORDS = {
    "manager", "analyst", "engineer", "developer", "admin", "administrator",
    "lead", "senior", "junior", "head", "director", "coordinator",
    "tester", "supervisor", "specialist", "consultant", "architect",
    "officer", "exec", "executive",
    "qa", "ba", "dba", "sa", "pm", "dev", "ops", "it", "hr",
    "development", "technical", "system", "software", "project", "business",
    "database", "security", "network",
}


def is_mock_person_name(name) -> tuple:
    """
    Returns (is_mock: bool, reason: str).

    Detects values that are clearly NOT real person names (e.g. "TBD", "QA Team",
    "DevTeamYa", "TDP BA", "Development TDP", etc.).
    Returns (False, "") for values that look plausibly like real First-Last names.
    """
    if not name:
        return True, "empty value"
    name = str(name).strip()
    if not name:
        return True, "empty value"

    name_lower = name.lower()

    # 1. Exact known placeholder
    if name_lower in PLACEHOLDER_EXACT:
        return True, f"known placeholder value: '{name}'"

    # 2. Single character (cannot be a full name)
    if len(name) <= 2:
        return True, "too short to be a valid full name"

    # 3. CamelCase single word without space  →  DevTeamYa, QAEngineer, BusinessAnalyst1
    if " " not in name and re.search(r"[a-z][A-Z]", name):
        return True, "CamelCase single word suggests a template variable (e.g. 'DevTeamYa')"

    # 4. Name ends with a group/department word
    words = name.split()
    if words and words[-1].lower() in TRAILING_GROUP_WORDS:
        return True, f"ends with '{words[-1]}' — looks like a department/group, not a person"

    # 5. Entire name consists only of role/function keywords
    words_lower = {w.lower().rstrip("s") for w in words}  # simple singular
    if words_lower and words_lower.issubset(PURE_ROLE_WORDS):
        return True, "name consists entirely of role/function keywords (not a person's name)"

    # 6. Starts with org-code acronym (2-5 ALL-CAPS) + role keyword
    #    e.g. "TDP BA", "TDP QA", "XYZ DBA Team"
    if len(words) >= 2 and re.match(r"^[A-Z]{2,5}$", words[0]):
        if words[1].lower() in PURE_ROLE_WORDS or words[1].lower() in TRAILING_GROUP_WORDS:
            return True, f"'{words[0]}' looks like an org acronym followed by a role/group (not a person)"

    # 7. Single English word that matches a well-known role title exactly
    if len(words) == 1:
        single = words[0].lower()
        if single in PURE_ROLE_WORDS or single in TRAILING_GROUP_WORDS:
            return True, f"single word '{words[0]}' is a role/function title, not a person's name"

    # 8. Name contains digits (e.g. "Developer1", "User123")
    if re.search(r"\d", name):
        return True, "contains digits — not typical for a personal name"

    return False, ""


# ════════════════════════════════════════════════════════════════════════════════
# Config Field Extraction
# ════════════════════════════════════════════════════════════════════════════════

def _collect_person_fields(config: dict) -> list:
    """
    Walk the config structure and yield (json_path, value) for every field
    that is supposed to hold a real person name.
    """
    fields = []
    team = config.get("team", {})

    # Core named roles
    for role in ["project_manager", "lead_developer", "qa_engineer",
                 "business_analyst", "system_analyst", "dba"]:
        person = team.get(role)
        if isinstance(person, dict):
            fields.append((f"team.{role}.name", person.get("name", "")))

    # Additional team members
    for i, m in enumerate(team.get("members", [])):
        if isinstance(m, dict):
            fields.append((f"team.members[{i}].name", m.get("name", "")))

    # Stakeholders
    for i, s in enumerate(config.get("stakeholders", [])):
        if isinstance(s, dict):
            fields.append((f"stakeholders[{i}].name", s.get("name", "")))

    # Meetings — chair and attendees
    for i, mtg in enumerate(config.get("meetings", [])):
        if not isinstance(mtg, dict):
            continue
        if mtg.get("chair"):
            fields.append((f"meetings[{i}].chair", mtg["chair"]))
        for j, att in enumerate(mtg.get("attendees", [])):
            if isinstance(att, str) and att:
                fields.append((f"meetings[{i}].attendees[{j}]", att))
            elif isinstance(att, dict) and att.get("name"):
                fields.append((f"meetings[{i}].attendees[{j}].name", att["name"]))

    # Training sessions — trainer
    for i, tr in enumerate(config.get("training_sessions", [])):
        if isinstance(tr, dict) and tr.get("trainer"):
            fields.append((f"training_sessions[{i}].trainer", tr["trainer"]))

    # Audit — auditor
    audit = config.get("audit", {})
    if isinstance(audit, dict) and audit.get("auditor"):
        fields.append(("audit.auditor", audit["auditor"]))

    # Incidents — reported_by, assigned_to
    for i, inc in enumerate(config.get("incidents", [])):
        if not isinstance(inc, dict):
            continue
        for field in ("reported_by", "assigned_to"):
            if inc.get(field):
                fields.append((f"incidents[{i}].{field}", inc[field]))

    # CAPAs — owner
    for i, capa in enumerate(config.get("capas", [])):
        if isinstance(capa, dict) and capa.get("owner"):
            fields.append((f"capas[{i}].owner", capa["owner"]))

    # Risks — owner
    for i, risk in enumerate(config.get("risks", [])):
        if isinstance(risk, dict) and risk.get("owner"):
            fields.append((f"risks[{i}].owner", risk["owner"]))

    # Versions — deployed_by
    for i, v in enumerate(config.get("versions", [])):
        if isinstance(v, dict) and v.get("deployed_by"):
            fields.append((f"versions[{i}].deployed_by", v["deployed_by"]))

    # Bug log — reported_by, assigned_to
    for i, bug in enumerate(config.get("bugs", [])):
        if not isinstance(bug, dict):
            continue
        for field in ("reported_by", "assigned_to"):
            if bug.get(field):
                fields.append((f"bugs[{i}].{field}", bug[field]))

    return fields


def detect_mock_values(config: dict) -> list:
    """
    Returns list of (json_path, current_value, reason) for all suspicious name
    fields found in the config.
    """
    issues = []
    for path, value in _collect_person_fields(config):
        is_mock, reason = is_mock_person_name(value)
        if is_mock:
            issues.append((path, value, reason))
    return issues


# ════════════════════════════════════════════════════════════════════════════════
# Reporting
# ════════════════════════════════════════════════════════════════════════════════

def print_report(project_code: str, issues: list, verbose: bool = True):
    if not issues:
        print(f"  OK  {project_code}: No mockup names detected.")
        return

    print(f"\n  WARN  {project_code} — {len(issues)} suspicious name field(s):")
    for path, value, reason in issues:
        display_val = f"'{value}'" if value else "(empty)"
        print(f"        {path}")
        print(f"          Current : {display_val}")
        print(f"          Reason  : {reason}")


# ════════════════════════════════════════════════════════════════════════════════
# Interactive Fixer
# ════════════════════════════════════════════════════════════════════════════════

def _set_nested_value(obj, path: str, value: str):
    """Set a value in a nested dict/list using path like 'team.members[0].name'."""
    # Tokenise: split on "." but keep [n] with the preceding key
    parts = []
    for token in re.split(r"\.", path):
        m = re.match(r"^(\w+)\[(\d+)\]$", token)
        if m:
            parts.append(m.group(1))
            parts.append(int(m.group(2)))
        else:
            parts.append(token)

    current = obj
    for p in parts[:-1]:
        current = current[p]

    last = parts[-1]
    current[last] = value


def interactive_fix(config: dict, issues: list) -> bool:
    """
    Prompt user to supply real person names for each flagged field.
    Returns True if at least one value was updated.
    """
    updated = False
    print()
    print("  Enter the real First Last name for each field (press Enter to skip).")
    print("  Example:  Somchai Jaidee  /  สมชาย ใจดี  /  Apisit Yimboa")
    print()

    for path, old_value, reason in issues:
        display = f"'{old_value}'" if old_value else "(empty)"
        print(f"  {path}  [currently {display}]")
        print(f"    Reason: {reason}")
        try:
            new_val = input("    New name (Enter = skip): ").strip()
        except (EOFError, KeyboardInterrupt):
            print("\n  Aborted.")
            break

        if new_val:
            _set_nested_value(config, path, new_val)
            print(f"    Updated: '{new_val}'")
            updated = True
        else:
            print("    Skipped.")

    return updated


# ════════════════════════════════════════════════════════════════════════════════
# CLI Entry Point
# ════════════════════════════════════════════════════════════════════════════════

def main() -> int:
    parser = argparse.ArgumentParser(
        description="Validate ISO project configs for mockup / placeholder person names."
    )
    src = parser.add_mutually_exclusive_group(required=True)
    src.add_argument("--config", metavar="PATH",
                     help="Path to a single project config JSON")
    src.add_argument("--all", action="store_true",
                     help="Scan all project configs under configs/*/")
    parser.add_argument("--fix", action="store_true",
                        help="After reporting, interactively replace mockup values")
    args = parser.parse_args()

    workspace = Path(__file__).parent.parent
    configs_dir = workspace / "configs"

    if args.all:
        config_paths = sorted(configs_dir.glob("*/*_config.json"))
        if not config_paths:
            print(f"No config files found under {configs_dir}")
            return 1
    else:
        config_paths = [Path(args.config)]

    total_issues = 0
    projects_with_issues = 0

    for config_path in config_paths:
        try:
            with open(config_path, encoding="utf-8") as f:
                config = json.load(f)
        except Exception as exc:
            print(f"  ERROR reading {config_path}: {exc}")
            continue

        project_code = config.get("project", {}).get("code", config_path.parent.name)
        folder_name = config_path.parent.name
        # Show folder name when it differs from the project code embedded in the config
        display_name = project_code if folder_name.upper() == project_code.upper() else f"{folder_name} (code: {project_code})"
        issues = detect_mock_values(config)
        total_issues += len(issues)
        if issues:
            projects_with_issues += 1

        print_report(display_name, issues)

        if args.fix and issues:
            if interactive_fix(config, issues):
                with open(config_path, "w", encoding="utf-8") as f:
                    json.dump(config, f, ensure_ascii=False, indent=2)
                print(f"\n  Saved: {config_path}\n")

    # Summary
    if len(config_paths) > 1:
        print()
        print("=" * 60)
        if total_issues == 0:
            print(f"  All {len(config_paths)} project(s) have no mockup names.")
        else:
            print(f"  {total_issues} issue(s) in {projects_with_issues}/{len(config_paths)} project(s).")
            if not args.fix:
                print("  Run with --fix to interactively replace mockup values.")
        print("=" * 60)

    return 0 if total_issues == 0 else 1


if __name__ == "__main__":
    sys.exit(main())
