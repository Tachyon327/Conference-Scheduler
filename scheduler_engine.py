#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
scheduler_engine.py
====================
Core, GUI-independent engine for the Conference Presentation Scheduling & Venue
Optimization System.

Input validation, the MILP model, cross-platform CBC resolution and Excel
export live here. No GUI imports, so the engine can be unit tested and
reused headlessly.

Where settings live
-------------------
The workbook holds **conference data only**: venues, timeslots, presentations
and which fields clash with each other. How hard to search (time limit,
MIP gap), how many alternatives to return, and what to prioritise (the
weights) are decisions made per run and are passed in by the caller.
Keeping each setting in exactly one place means a result can never be
ambiguous about which value produced it; the output workbook records the
values that were actually used.

The conflict term
-----------------
``B1`` charges once when two *different* clashing topics land in the same
timeslot. ``B2`` charges per *excess* presentation when one topic is scheduled
in parallel with itself. The earlier implementation summed the two
topics' presentation counts, which conflated the two situations and turned a
soft constraint into an implicit hard one -- three presentations of a single
topic in one timeslot made the model infeasible.

Auxiliary variables are continuous: the objective pushes each in one
direction only, so relaxing them leaves the optimal value unchanged while
removing them from branching.
"""

from __future__ import annotations

import datetime as dt
import functools
import math
import os
import platform
import re
import shutil
import subprocess
import sys
import time
from dataclasses import dataclass, field
from typing import Dict, List, Optional, Sequence, Set, Tuple

import pandas as pd
import pulp

# ``_t`` rather than ``t``: the model code uses ``t`` for timeslots.
from scheduler_messages import t as _t

# =====================================================================
# Constants
# =====================================================================

#: Sheets that must be present in every input workbook.
REQUIRED_SHEETS: Tuple[str, ...] = ("Rooms", "TimeSlots", "Presentations")

#: Sheets used when present and skipped when absent.
OPTIONAL_SHEETS: Tuple[str, ...] = ("Conflicts",)

#: Every sheet ``validate_input_workbook`` returns.
ALL_SHEETS: Tuple[str, ...] = REQUIRED_SHEETS + OPTIONAL_SHEETS

#: Columns required in each required sheet.
REQUIRED_COLUMNS: Dict[str, Tuple[str, ...]] = {
    "Rooms": ("Room_ID", "Capacity"),
    "TimeSlots": ("TimeSlot_ID", "Day"),
    "Presentations": ("Presentation_ID", "Topic", "Expected_Pax"),
}

#: Columns required in an optional sheet *if that sheet is present*.
OPTIONAL_SHEET_COLUMNS: Dict[str, Tuple[str, ...]] = {
    "Conflicts": ("Topic1", "Topic2", "Conflict_Level"),
}

#: Sheets that older workbooks carried and that are no longer read.
RETIRED_SHEETS: Tuple[str, ...] = ("Settings",)

#: At most this many input problems are listed at once.
MAX_REPORTED_PROBLEMS = 20

#: The three soft-constraint terms are measured in different units --
#: continuity and conflicts are counts, empty seats are seats -- so a
#: fixed set of base weights can only balance them for one particular
#: dataset. Each term is therefore divided by the range it can actually
#: span for the data at hand, which puts all three on a 0-1 scale and
#: makes a weight of 1 mean the same thing everywhere. The result is
#: multiplied by SCORE_SCALE purely so the score reads as a familiar
#: two- or three-digit number rather than a fraction.
SCORE_SCALE = 100.0

#: Conflict level applied to a topic competing with *itself* (B2).
#: Fixed rather than configurable: running two talks of one field in
#: parallel is at least as bad for an attendee as the strongest
#: cross-field clash, and keeping it out of the input keeps the workbook
#: small.
SELF_CONFLICT_LEVEL = 3

#: Defaults for the per-run settings.
DEFAULT_WEIGHT = 1.0
DEFAULT_MIP_GAP = 0.0          # 0 solves to proven optimality
DEFAULT_TIME_LIMIT: Optional[int] = None
DEFAULT_SOLUTIONS = 1
MAX_SOLUTIONS = 3

#: Minimum number of assignments by which candidate *k+1* must differ
#: from every earlier candidate. Without this, "the next best schedule"
#: is often the same schedule with two interchangeable rooms swapped.
MIN_SOLUTION_DIFFERENCE = 3
SOLUTION_DIFFERENCE_RATIO = 0.05


class InputValidationError(ValueError):
    """Raised when the input workbook cannot be used as-is."""


class InfeasibleModelError(RuntimeError):
    """Raised when the MILP has no feasible solution, or CBC found none."""


# =====================================================================
# 1. Small helpers
# =====================================================================

def _is_blank(value) -> bool:
    """True for None, NaN and whitespace-only strings."""
    if value is None:
        return True
    try:
        if pd.isna(value):
            return True
    except (TypeError, ValueError):
        pass
    return str(value).strip() == ""


def truthy(value) -> bool:
    """Interpret a spreadsheet cell as a yes/no flag.

    Accepts 1/0, True/False, "Yes"/"No", "TRUE"/"FALSE", "有"/"無".
    Comparing against the literal ``1`` (as the earlier version did) meant
    a workbook written with ``TRUE`` silently lost the constraint.
    """
    if _is_blank(value):
        return False
    if isinstance(value, bool):
        return value
    if isinstance(value, (int, float)):
        return float(value) != 0.0
    text = str(value).strip().lower()
    if text in ("0", "false", "no", "n", "none", "-", "x", "×", "無", "なし"):
        return False
    return True


def parse_ng_slots(val) -> Set[str]:
    """Parse a (possibly messy) ``NG_TimeSlot_ID`` cell into a set of IDs.

    Accepts comma, full-width comma, Japanese "、" and whitespace as
    separators, and tolerates NaN/blank input.
    """
    if _is_blank(val):
        return set()
    safe_str = re.sub(r"[，、\s]+", ",", str(val).strip())
    return {t for t in safe_str.split(",") if t}


def parse_clock(value) -> Optional[int]:
    """Convert a Start_Time cell into minutes since midnight, or None."""
    if _is_blank(value):
        return None
    if isinstance(value, dt.time):
        return value.hour * 60 + value.minute
    if isinstance(value, (dt.datetime, pd.Timestamp)):
        return value.hour * 60 + value.minute
    match = re.match(r"^(\d{1,2}):(\d{2})", str(value).strip())
    if not match:
        return None
    hour, minute = int(match.group(1)), int(match.group(2))
    lowered = str(value).lower()
    if "pm" in lowered and hour < 12:
        hour += 12
    elif "am" in lowered and hour == 12:
        hour = 0
    if not (0 <= hour <= 23 and 0 <= minute <= 59):
        return None
    return hour * 60 + minute


def sort_timeslots(df_timeslots: pd.DataFrame) -> pd.DataFrame:
    """Return the timeslots in chronological order.

    The continuity bonus is awarded to *adjacent* timeslots, and the
    earlier implementation decided adjacency purely from the sheet's row
    order -- ``Start_Time`` was never read, so a workbook whose rows were
    not already chronological earned bonuses between slots that are not
    adjacent, with no error or warning.

    Days keep their order of first appearance (so "Day2" and "Day10"
    behave sensibly) and rows within a day are ordered by ``Start_Time``.
    If any ``Start_Time`` cannot be parsed, the sheet's own order is kept
    for the whole sheet rather than half-sorting it.
    """
    if df_timeslots.empty:
        return df_timeslots.reset_index(drop=True)

    day_order = {day: i for i, day in enumerate(dict.fromkeys(df_timeslots["Day"].tolist()))}
    ordered = df_timeslots.copy()
    ordered["_day_rank"] = [day_order[d] for d in ordered["Day"]]

    minutes: List[int] = []
    if "Start_Time" in ordered.columns:
        for raw in ordered["Start_Time"]:
            parsed = parse_clock(raw)
            if parsed is None:
                minutes = []
                break
            minutes.append(parsed)

    if minutes:
        ordered["_minutes"] = minutes
        ordered = ordered.sort_values(["_day_rank", "_minutes"], kind="stable")
        ordered = ordered.drop(columns=["_minutes"])
    else:
        ordered = ordered.sort_values(["_day_rank"], kind="stable")

    return ordered.drop(columns=["_day_rank"]).reset_index(drop=True)


# =====================================================================
# 2. Dynamic path resolution & solver management
# =====================================================================

def get_base_dir() -> str:
    """Base directory for bundled resources (solvers/, sample data, ...)."""
    if getattr(sys, "frozen", False):
        return getattr(sys, "_MEIPASS", os.path.dirname(sys.executable))
    return os.path.dirname(os.path.abspath(__file__))


#: How long a CBC binary gets to answer ``cbc -quit`` before it counts as
#: unusable. Generous on purpose: the first launch of an unsigned binary,
#: or of an Intel binary under Rosetta 2, can take several seconds.
CBC_PROBE_TIMEOUT_SEC = 30


class SolverUnavailableError(RuntimeError):
    """Raised when no CBC binary on this machine can be started."""


@dataclass(frozen=True)
class CbcInfo:
    """Which CBC binary will solve the model, and where it came from.

    ``source`` is one of ``"Bundled"`` (shipped in ``solvers/``),
    ``"System"`` (installed separately, e.g. with Homebrew) or ``"PuLP"``
    (the copy inside the PuLP package).
    """

    path: str
    source: str
    version: str


def _cbc_candidates() -> List[Tuple[str, str]]:
    """(path, source) pairs to try, most preferred first."""
    base_dir = get_base_dir()
    current_os = platform.system().lower()
    if "windows" in current_os:
        folder, exe = "windows", "cbc.exe"
        system_paths = [
            r"C:\Program Files\Cbc\bin\cbc.exe",
            r"C:\Program Files (x86)\Cbc\bin\cbc.exe",
            shutil.which("cbc.exe"),
            shutil.which("cbc"),
        ]
    elif "darwin" in current_os:
        folder, exe = "mac", "cbc"
        system_paths = [
            shutil.which("cbc"),
            "/opt/homebrew/bin/cbc",
            "/usr/local/bin/cbc",
            "/opt/local/bin/cbc",
        ]
    else:  # Linux / other Unix
        # NOTE: the earlier version wrote
        #   shutil.which("cbc") or "/usr/bin/cbc" or "/usr/local/bin/cbc"
        # which always evaluates to "/usr/bin/cbc" when `which` fails,
        # because a non-empty string literal is always truthy.
        folder, exe = "linux", "cbc"
        system_paths = [shutil.which("cbc"), "/usr/bin/cbc", "/usr/local/bin/cbc"]

    candidates = [
        (os.path.join(base_dir, "solvers", folder, exe), "Bundled"),
        (os.path.join(os.getcwd(), "solvers", folder, exe), "Bundled"),
        (os.path.join(base_dir, exe), "Bundled"),
    ]
    candidates += [(p, "System") for p in system_paths if p]
    try:
        pulp_path = pulp.PULP_CBC_CMD(msg=0).path
    except Exception:
        pulp_path = None
    if pulp_path:
        candidates.append((pulp_path, "PuLP"))

    unique: List[Tuple[str, str]] = []
    seen: Set[str] = set()
    for path, source in candidates:
        key = os.path.realpath(path)
        if key not in seen:
            seen.add(key)
            unique.append((path, source))
    return unique


@functools.lru_cache(maxsize=None)
def _probe_cbc(path: str) -> Optional[str]:
    """Start ``path`` once and return its version, or None if it will not run.

    Checking that the file *exists* is not enough: a CBC copied out of a
    Homebrew installation is present on every machine it is shipped to
    but only starts where its libraries are installed too. The earlier
    version found the file, chose it, and then failed mid-solve instead
    of moving on to a CBC that works.
    """
    if not os.path.isfile(path):
        return None
    if platform.system().lower() != "windows":
        try:
            os.chmod(path, 0o755)
        except OSError:
            pass
    kwargs = {}
    if platform.system().lower() == "windows":
        kwargs["creationflags"] = 0x08000000  # CREATE_NO_WINDOW
    try:
        completed = subprocess.run(
            [path, "-quit"], capture_output=True, text=True,
            timeout=CBC_PROBE_TIMEOUT_SEC, **kwargs,
        )
    except (OSError, subprocess.SubprocessError):
        return None
    if completed.returncode != 0:
        return None
    match = re.search(r"Version:\s*([0-9][0-9.]*)", completed.stdout or "")
    return match.group(1) if match else "unknown"


def resolve_cbc() -> CbcInfo:
    """Pick the first CBC that actually starts on this machine."""
    for path, source in _cbc_candidates():
        if not os.path.isfile(path):
            continue
        version = _probe_cbc(path)
        if version is not None:
            return CbcInfo(path=path, source=source, version=version)
        print(f"[Solver] CBC at '{path}' ({source}) could not be started; trying the next one.")
    raise SolverUnavailableError(_t("solver.unavailable"))


def get_cbc_solver(
    time_limit: Optional[int] = None,
    msg: int = 1,
    gap_rel: Optional[float] = None,
):
    """A PuLP solver object for the CBC chosen by ``resolve_cbc``."""
    info = resolve_cbc()
    gap_kwargs = {"gapRel": gap_rel} if gap_rel else {}
    print(f"[Solver] Using CBC {info.version} ({info.source}): {info.path}")
    return pulp.COIN_CMD(path=info.path, timeLimit=time_limit, msg=msg, **gap_kwargs)


def open_file_or_folder(target_path: str) -> None:
    """Open a file or directory with the OS's native handler."""
    if not os.path.exists(target_path):
        raise FileNotFoundError(f"Path does not exist: {target_path}")

    current_os = platform.system().lower()
    if "windows" in current_os:
        os.startfile(target_path)  # type: ignore[attr-defined]
    elif "darwin" in current_os:
        subprocess.run(["open", target_path], check=True)
    else:
        subprocess.run(["xdg-open", target_path], check=True)


# =====================================================================
# 3. Input validation
# =====================================================================

def _excel_row(index: int) -> int:
    """Spreadsheet row number for a zero-based DataFrame index."""
    return index + 2  # +1 for the header row, +1 because Excel counts from 1


def _check_ids(frame: pd.DataFrame, sheet: str, column: str) -> List[str]:
    problems: List[str] = []
    values = frame[column]
    for idx, value in enumerate(values):
        if _is_blank(value):
            problems.append(_t("input.blank", sheet=sheet, column=column, row=_excel_row(idx)))
    cleaned = [str(v).strip() for v in values if not _is_blank(v)]
    seen: Dict[str, int] = {}
    for value in cleaned:
        seen[value] = seen.get(value, 0) + 1
    duplicates = sorted(v for v, n in seen.items() if n > 1)
    if duplicates:
        values = ", ".join(duplicates[:5]) + ("…" if len(duplicates) > 5 else "")
        problems.append(_t("input.duplicate", sheet=sheet, column=column, values=values))
    return problems


def _check_numeric(frame: pd.DataFrame, sheet: str, column: str) -> List[str]:
    problems: List[str] = []
    for idx, value in enumerate(frame[column]):
        if _is_blank(value):
            problems.append(_t("input.blank", sheet=sheet, column=column, row=_excel_row(idx)))
            continue
        try:
            number = float(value)
        except (TypeError, ValueError):
            problems.append(
                _t("input.not_number", sheet=sheet, column=column, row=_excel_row(idx), value=value)
            )
            continue
        if number < 0:
            problems.append(
                _t("input.negative", sheet=sheet, column=column, row=_excel_row(idx), value=value)
            )
    return problems


def _validate_data(sheets: Dict[str, pd.DataFrame]) -> List[str]:
    """Check the *contents* of an otherwise well-formed workbook.

    Sheet and column checks alone let real mistakes through: a duplicated
    ``Presentation_ID`` used to produce a wrong schedule without any error at
    all, and a non-numeric ``Capacity`` surfaced as a Python type error
    from deep inside the model.
    """
    problems: List[str] = []
    rooms, slots, presentations = sheets["Rooms"], sheets["TimeSlots"], sheets["Presentations"]

    for frame, sheet in ((rooms, "Rooms"), (slots, "TimeSlots"), (presentations, "Presentations")):
        if frame.empty:
            problems.append(_t("input.empty_sheet", sheet=sheet))

    if not rooms.empty:
        problems += _check_ids(rooms, "Rooms", "Room_ID")
        problems += _check_numeric(rooms, "Rooms", "Capacity")
    if not slots.empty:
        problems += _check_ids(slots, "TimeSlots", "TimeSlot_ID")
        for idx, value in enumerate(slots["Day"]):
            if _is_blank(value):
                problems.append(
                    _t("input.blank", sheet="TimeSlots", column="Day", row=_excel_row(idx)))
    if not presentations.empty:
        problems += _check_ids(presentations, "Presentations", "Presentation_ID")
        problems += _check_numeric(presentations, "Presentations", "Expected_Pax")
        for idx, value in enumerate(presentations["Topic"]):
            if _is_blank(value):
                problems.append(
                    _t("input.blank", sheet="Presentations", column="Topic", row=_excel_row(idx)))

    conflicts = sheets.get("Conflicts")
    if conflicts is not None and not conflicts.empty:
        problems += _check_numeric(conflicts, "Conflicts", "Conflict_Level")

    # A presentation can only ever occupy one room in one timeslot, so more
    # presentations than room-slots is impossible however they are arranged.
    if not rooms.empty and not slots.empty and not presentations.empty:
        capacity = len(rooms) * len(slots)
        if len(presentations) > capacity:
            problems.append(_t("input.too_many_presentations", presentations=len(presentations),
                              rooms=len(rooms), slots=len(slots), capacity=capacity))

    # Referring to a timeslot that does not exist is almost always a typo.
    if "NG_TimeSlot_ID" in presentations.columns and not slots.empty:
        known = {str(t).strip() for t in slots["TimeSlot_ID"] if not _is_blank(t)}
        unknown: Set[str] = set()
        for value in presentations["NG_TimeSlot_ID"]:
            unknown |= {t for t in parse_ng_slots(value) if t not in known}
        if unknown:
            values = ", ".join(sorted(unknown)[:5]) + ("…" if len(unknown) > 5 else "")
            problems.append(_t("input.unknown_ng_slot", values=values))

    return problems


def validate_input_workbook(input_file: str) -> Dict[str, pd.DataFrame]:
    """Load and check the input workbook, returning its sheets.

    Only ``Rooms``, ``TimeSlots`` and ``Presentations`` are required, and
    only six columns between them. ``Conflicts`` is optional and comes back
    as an empty frame when absent; an optional sheet that *is* present
    must still carry its own columns, because a misspelt header is a
    mistake rather than an intention to switch the feature off.

    Every problem found is reported at once so a user fixing their
    spreadsheet doesn't have to re-run the tool once per mistake.
    """
    if not os.path.exists(input_file):
        raise FileNotFoundError(_t("input.file_not_found", path=input_file))

    try:
        workbook = pd.read_excel(input_file, sheet_name=None)
    except Exception as exc:
        raise InputValidationError(
            _t("input.unreadable", name=os.path.basename(input_file), detail=exc)
        ) from exc

    problems: List[str] = []
    for sheet in REQUIRED_SHEETS:
        if sheet not in workbook:
            problems.append(_t("input.missing_sheet", sheet=sheet))
            continue
        missing = [c for c in REQUIRED_COLUMNS[sheet] if c not in workbook[sheet].columns]
        if missing:
            problems.append(_t("input.missing_columns", sheet=sheet, columns=", ".join(missing)))

    for sheet in OPTIONAL_SHEETS:
        if sheet not in workbook:
            continue
        frame = workbook[sheet]
        if len(frame.columns) == 0:
            continue
        missing = [c for c in OPTIONAL_SHEET_COLUMNS[sheet] if c not in frame.columns]
        if missing:
            problems.append(
                _t("input.missing_columns_optional", sheet=sheet, columns=", ".join(missing))
            )

    sheets: Dict[str, pd.DataFrame] = {}
    if not problems:
        sheets = {name: workbook[name] for name in REQUIRED_SHEETS}
        for name in OPTIONAL_SHEETS:
            if name in workbook and len(workbook[name].columns) > 0:
                sheets[name] = workbook[name]
            else:
                sheets[name] = pd.DataFrame(columns=list(OPTIONAL_SHEET_COLUMNS[name]))
        problems += _validate_data(sheets)

    if problems:
        shown = problems[:MAX_REPORTED_PROBLEMS]
        more = len(problems) - len(shown)
        message = _t("input.header") + "\n" + "\n".join(f"  - {p}" for p in shown)
        if more > 0:
            message += "\n" + _t("input.more", count=more)
        message += "\n\n" + _t("input.readme_hint")
        raise InputValidationError(message)

    return sheets


def summarize_input(input_file: str) -> Dict:
    """Validate a workbook and report its shape, without solving anything.

    Lets the GUI tell the user what it read -- and what is wrong with the
    file -- the moment they pick it, rather than after a failed run.
    """
    sheets = validate_input_workbook(input_file)
    presentations = sheets["Presentations"]
    speakers = 0
    if "Speaker" in presentations.columns:
        speakers = len({str(s).strip() for s in presentations["Speaker"] if not _is_blank(s)})
    return {
        "presentations": len(presentations),
        "rooms": len(sheets["Rooms"]),
        "timeslots": len(sheets["TimeSlots"]),
        "topics": presentations["Topic"].astype(str).nunique(),
        "speakers": speakers,
        "conflict_pairs": len(sheets["Conflicts"]),
        "slots_available": len(sheets["Rooms"]) * len(sheets["TimeSlots"]),
    }


# =====================================================================
# 4. Excel export
# =====================================================================

def write_workbook_safe(output_file: str, sheets: Dict[str, pd.DataFrame]) -> str:
    """Write ``sheets`` to ``output_file``, falling back to a timestamped
    filename if the target is locked (e.g. already open in Excel)."""
    target_file = output_file
    output_dir = os.path.dirname(os.path.abspath(target_file))
    if output_dir and not os.path.exists(output_dir):
        os.makedirs(output_dir, exist_ok=True)

    def _write(path: str) -> None:
        with pd.ExcelWriter(path, engine="openpyxl") as writer:
            for name, frame in sheets.items():
                frame.to_excel(writer, sheet_name=name, index=False)

    try:
        _write(target_file)
        return target_file
    except Exception as exc:
        stamp = time.strftime("%Y%m%d_%H%M%S")
        base, ext = os.path.splitext(target_file)
        fallback_file = f"{base}_{stamp}{ext}"
        print(f"[Warning] '{target_file}' is in use or locked ({exc}). "
              f"Saving to '{fallback_file}' instead...")
        _write(fallback_file)
        return fallback_file


def save_excel_safe(
    output_file: str,
    df_out_matrix: pd.DataFrame,
    df_out_list: pd.DataFrame,
    df_report: pd.DataFrame,
) -> str:
    """Backwards-compatible single-candidate export."""
    return write_workbook_safe(
        output_file,
        {
            "Result_Timetable": df_out_matrix,
            "Result_List": df_out_list,
            "Simulation_Report": df_report,
        },
    )


# =====================================================================
# 5. Result containers
# =====================================================================

@dataclass
class Candidate:
    """One schedule produced by the solver."""

    rank: int
    status: str
    solution_quality: str
    objective_score: float
    calc_time: float
    assignments: List[Dict] = field(default_factory=list)
    continuity_bonuses: int = 0
    cross_topic_conflicts: int = 0
    cross_topic_level_sum: float = 0.0
    same_topic_overlaps: int = 0
    empty_seats: float = 0.0
    score_continuity: float = 0.0
    score_conflict: float = 0.0
    score_same_topic: float = 0.0
    score_empty_seats: float = 0.0

    def as_dict(self) -> Dict:
        return {
            "rank": self.rank,
            "objective_score": self.objective_score,
            "calc_time": self.calc_time,
            "solution_quality": self.solution_quality,
            "continuity_bonuses": self.continuity_bonuses,
            "cross_topic_conflicts": self.cross_topic_conflicts,
            "same_topic_overlaps": self.same_topic_overlaps,
            "empty_seats": self.empty_seats,
            "score_continuity": self.score_continuity,
            "score_conflict": self.score_conflict,
            "score_same_topic": self.score_same_topic,
            "score_empty_seats": self.score_empty_seats,
        }


@dataclass
class OptimizationResult:
    status: str
    objective_score: float
    calc_time: float
    output_file: str
    assigned_count: int
    model_variables: int = 0
    model_constraints: int = 0
    solution_quality: str = ""
    total_calc_time: float = 0.0
    solutions_requested: int = 1
    solutions_returned: int = 1
    candidates: List[Dict] = field(default_factory=list)

    def as_dict(self) -> Dict:
        return {
            "status": self.status,
            "objective_score": self.objective_score,
            "calc_time": self.calc_time,
            "output_file": self.output_file,
            "assigned_count": self.assigned_count,
            "model_variables": self.model_variables,
            "model_constraints": self.model_constraints,
            "solution_quality": self.solution_quality,
            "total_calc_time": self.total_calc_time,
            "solutions_requested": self.solutions_requested,
            "solutions_returned": self.solutions_returned,
            "candidates": self.candidates,
        }


def _describe_solution_quality(prob: pulp.LpProblem, mip_gap: float) -> str:
    """Say what the returned solution actually is.

    CBC reports ``Optimal`` even when a time limit cut the search short,
    which is why the earlier version could label a demonstrably
    suboptimal schedule "Optimal". PuLP's ``sol_status`` distinguishes
    the two.
    """
    if getattr(prob, "sol_status", None) == pulp.LpSolutionIntegerFeasible:
        return "Feasible (time limit reached; optimality not proven)"
    if mip_gap > 0:
        return f"Optimal within {mip_gap * 100:g}% gap"
    return "Optimal (proven)"


# =====================================================================
# 6. Optimization core
# =====================================================================

def run_optimization(
    input_file: str,
    output_file: str,
    *,
    weight_continuity: float = DEFAULT_WEIGHT,
    weight_conflict: float = DEFAULT_WEIGHT,
    weight_empty_seats: float = DEFAULT_WEIGHT,
    mip_gap: float = DEFAULT_MIP_GAP,
    time_limit: Optional[int] = DEFAULT_TIME_LIMIT,
    solutions: int = DEFAULT_SOLUTIONS,
) -> Dict:
    """Run the presentation-scheduling optimization and write the result.

    The workbook supplies the conference data; every setting above is
    supplied by the caller (GUI or command line). Returns a plain dict.
    """
    print("=" * 70)
    print("=== Conference Presentation Scheduling & Venue Optimization System ===")
    print("=" * 70)
    print(f"Input file  : {os.path.abspath(input_file)}")
    print(f"Output file : {os.path.abspath(output_file)}")

    # -------------------------------------------------------------
    # Step 1: read, validate and normalise the input
    # -------------------------------------------------------------
    print("\n[1/7] Reading and validating input workbook...")
    sheets = validate_input_workbook(input_file)
    df_rooms = sheets["Rooms"].copy()
    df_timeslots = sheets["TimeSlots"].copy()
    df_presentations = sheets["Presentations"].copy()
    df_conflicts = sheets["Conflicts"].copy()

    retired = [s for s in RETIRED_SHEETS if s in pd.ExcelFile(input_file).sheet_names]
    if retired:
        print(_t("run.settings_sheet_ignored", sheets=", ".join(retired)))

    use_equipment = "Screen" in df_rooms.columns and "Needs_Screen" in df_presentations.columns
    use_speaker = "Speaker" in df_presentations.columns

    if "Room_Name" not in df_rooms.columns:
        df_rooms["Room_Name"] = df_rooms["Room_ID"]
    if "NG_TimeSlot_ID" not in df_presentations.columns:
        df_presentations["NG_TimeSlot_ID"] = ""

    df_rooms["Room_ID"] = df_rooms["Room_ID"].astype(str).str.strip()
    df_rooms["Room_Name"] = df_rooms["Room_Name"].fillna(df_rooms["Room_ID"]).astype(str)
    df_rooms["Capacity"] = df_rooms["Capacity"].astype(float)
    df_presentations["Presentation_ID"] = df_presentations["Presentation_ID"].astype(str).str.strip()
    df_presentations["Topic"] = df_presentations["Topic"].astype(str).str.strip()
    df_presentations["Expected_Pax"] = df_presentations["Expected_Pax"].astype(float)
    df_timeslots["TimeSlot_ID"] = df_timeslots["TimeSlot_ID"].astype(str).str.strip()

    # Chronological order matters for the continuity bonus (see sort_timeslots).
    df_timeslots = sort_timeslots(df_timeslots)

    presentation_ng_slots = {
        row["Presentation_ID"]: parse_ng_slots(row.get("NG_TimeSlot_ID"))
        for _, row in df_presentations.iterrows()
    }

    presentation_speaker: Dict[str, str] = {}
    if use_speaker:
        for _, row in df_presentations.iterrows():
            speaker = row.get("Speaker")
            if not _is_blank(speaker):
                presentation_speaker[str(row["Presentation_ID"])] = str(speaker).strip()

    # -------------------------------------------------------------
    # Step 2: settings supplied by the caller
    # -------------------------------------------------------------
    print("[2/7] Applying weights and run settings...")
    w_a = float(weight_continuity)
    w_b = float(weight_conflict)
    w_c = float(weight_empty_seats)

    mip_gap = max(0.0, float(mip_gap))
    time_limit = int(time_limit) if time_limit is not None and time_limit > 0 else None
    n_solutions = max(1, min(MAX_SOLUTIONS, int(solutions)))

    presentation_pax = df_presentations.set_index("Presentation_ID")["Expected_Pax"].to_dict()
    presentation_topic = df_presentations.set_index("Presentation_ID")["Topic"].to_dict()
    presentation_needs_screen = (
        df_presentations.set_index("Presentation_ID")["Needs_Screen"].to_dict() if use_equipment else {}
    )

    room_cap = df_rooms.set_index("Room_ID")["Capacity"].to_dict()
    room_names = df_rooms.set_index("Room_ID")["Room_Name"].to_dict()
    room_screen = df_rooms.set_index("Room_ID")["Screen"].to_dict() if use_equipment else {}

    presentations_list = df_presentations["Presentation_ID"].tolist()
    rooms = df_rooms["Room_ID"].tolist()
    timeslots = df_timeslots["TimeSlot_ID"].tolist()
    timeslot_days = df_timeslots.set_index("TimeSlot_ID")["Day"].to_dict()
    topics = list(dict.fromkeys(df_presentations["Topic"].tolist()))

    presentations_by_topic = {g: [s for s in presentations_list if presentation_topic[s] == g] for g in topics}

    speakers_by_name: Dict[str, List[str]] = {}
    for pres, name in presentation_speaker.items():
        speakers_by_name.setdefault(name, []).append(pres)

    conflict_levels: Dict[Tuple[str, str], float] = {}
    if not df_conflicts.empty:
        for _, row in df_conflicts.iterrows():
            g1, g2 = str(row["Topic1"]).strip(), str(row["Topic2"]).strip()
            if _is_blank(g1) or _is_blank(g2):
                continue
            try:
                level = float(row["Conflict_Level"])
            except (TypeError, ValueError):
                continue
            if level <= 0 or g1 == g2:
                # Same-topic parallelism is handled by B2 with a fixed
                # level, so a self-pair row here is redundant, not wrong.
                continue
            conflict_levels[(g1, g2)] = level
            conflict_levels[(g2, g1)] = level

    print(f" -> Found {len(presentations_list)} presentations, {len(rooms)} rooms, "
          f"{len(timeslots)} timeslots across {len(topics)} topics.")
    print(f" -> Weights: Continuity={weight_continuity}, Conflict={weight_conflict}, "
          f"EmptySeats={weight_empty_seats}")
    print(f" -> Equipment constraint: {'enabled' if use_equipment else 'disabled (columns absent)'}")
    print(f" -> Speaker constraint  : {'enabled' if speakers_by_name else 'disabled (column absent/empty)'}")
    if mip_gap > 0:
        print(f" -> MIP gap enabled: solver may stop within {mip_gap * 100:.2f}% of proven optimal.")
    if n_solutions > 1:
        print(f" -> Requesting {n_solutions} candidate schedules "
              f"(each solved from scratch, so expect roughly {n_solutions}x the runtime).")

    # -------------------------------------------------------------
    # Step 3: feasible (presentation, room, timeslot) combinations
    # -------------------------------------------------------------
    # Infeasible combinations never become variables, which is
    # mathematically identical to forcing them to zero but yields a much
    # smaller model.
    feasible_rt: Dict[str, List[Tuple[str, str]]] = {}
    infeasible_presentations: List[str] = []
    for s in presentations_list:
        pax = presentation_pax[s]
        needs_screen = truthy(presentation_needs_screen.get(s)) if use_equipment else False
        ng_slots = presentation_ng_slots.get(s, set())
        combos = [
            (r, t)
            for r in rooms
            for t in timeslots
            if room_cap[r] >= pax
            and (not needs_screen or truthy(room_screen.get(r)))
            and t not in ng_slots
        ]
        feasible_rt[s] = combos
        if not combos:
            infeasible_presentations.append(s)

    if infeasible_presentations:
        details = "\n".join(
            _t("run.no_room_line", presentation=s, pax=f"{presentation_pax[s]:g}",
              screen=(_t("run.no_room_screen", value=presentation_needs_screen.get(s))
                      if use_equipment else ""))
            for s in infeasible_presentations[:10]
        )
        raise InfeasibleModelError(_t("run.no_room", details=details))

    # A speaker occupies at most one timeslot at a time, so a speaker with
    # more presentations than usable timeslots can never be scheduled. Saying
    # so by name beats an unexplained "Infeasible" from the solver.
    speaker_problems: List[str] = []
    for name, owned in speakers_by_name.items():
        if len(owned) < 2:
            continue
        usable = set()
        for s in owned:
            usable.update(t for _, t in feasible_rt[s])
        if len(owned) > len(usable):
            speaker_problems.append(
                _t("run.speaker_line", name=name, count=len(owned), usable=len(usable),
                  presentations=", ".join(owned))
            )
    if speaker_problems:
        raise InfeasibleModelError(
            _t("run.speaker_overlap", details="\n".join(speaker_problems)))

    # -------------------------------------------------------------
    # Step 4: build the MILP
    # -------------------------------------------------------------
    print("[3/7] Constructing MILP mathematical model...")
    prob = pulp.LpProblem("Conference_Scheduling", pulp.LpMaximize)

    x: Dict[Tuple[str, str, str], pulp.LpVariable] = {}
    x_by_presentation: Dict[str, List[pulp.LpVariable]] = {s: [] for s in presentations_list}
    x_by_room_slot: Dict[Tuple[str, str], List[pulp.LpVariable]] = {}
    x_by_topic_slot: Dict[Tuple[str, str], List[pulp.LpVariable]] = {}
    x_by_speaker_slot: Dict[Tuple[str, str], List[pulp.LpVariable]] = {}
    for s in presentations_list:
        topic = presentation_topic[s]
        speaker = presentation_speaker.get(s)
        for r, t in feasible_rt[s]:
            var = pulp.LpVariable(f"x_{s}_{r}_{t}", cat="Binary")
            x[(s, r, t)] = var
            x_by_presentation[s].append(var)
            x_by_room_slot.setdefault((r, t), []).append(var)
            x_by_topic_slot.setdefault((t, topic), []).append(var)
            if speaker:
                x_by_speaker_slot.setdefault((t, speaker), []).append(var)

    # Raw terms in their own natural units; normalised together below.
    continuity_terms: List = []   # counts
    conflict_terms: List = []     # level-weighted counts
    empty_seat_terms: List = []   # seats

    # --- Soft constraint A: topic continuity bonus -------------------
    # Rewards the same topic occupying the same room in consecutive
    # timeslots on the same day. ``z`` is pushed *up* by the objective and
    # bounded above by both endpoints, so a continuous variable lands on 0
    # or 1 on its own -- declaring it binary only adds branching.
    if w_a > 0:
        for g in topics:
            pres_g = presentations_by_topic[g]
            for r in rooms:
                for idx in range(len(timeslots) - 1):
                    t_curr, t_next = timeslots[idx], timeslots[idx + 1]
                    if timeslot_days[t_curr] != timeslot_days[t_next]:
                        continue
                    curr_vars = [x[(s, r, t_curr)] for s in pres_g if (s, r, t_curr) in x]
                    next_vars = [x[(s, r, t_next)] for s in pres_g if (s, r, t_next) in x]
                    if not curr_vars or not next_vars:
                        continue
                    z = pulp.LpVariable(
                        f"z_cons_{g}_{r}_{idx}", lowBound=0, upBound=1, cat="Continuous"
                    )
                    prob += z <= pulp.lpSum(curr_vars)
                    prob += z <= pulp.lpSum(next_vars)
                    continuity_terms.append(z)

    # --- Presence indicator, shared by B1 and the valid inequality ----
    # ``y[t][g]`` records whether topic g appears in timeslot t at all.
    # The ``y >= x`` direction is what the earlier implementation lacked:
    # without it the model compared raw presentation *counts*, so three talks
    # of a single topic in one timeslot made ``z >= 2`` collide with
    # ``z <= 1`` and the whole problem became infeasible.
    topic_present: Dict[Tuple[str, str], pulp.LpVariable] = {}
    if w_b > 0:
        for t in timeslots:
            for g in topics:
                group_vars = x_by_topic_slot.get((t, g))
                if not group_vars:
                    continue
                y = pulp.LpVariable(f"y_{t}_{g}", lowBound=0, upBound=1, cat="Continuous")
                for var in group_vars:
                    prob += y >= var
                topic_present[(t, g)] = y

    # --- Soft constraint B1: cross-topic conflict penalty -------------
    if w_b > 0 and conflict_levels:
        for t in timeslots:
            for i, g1 in enumerate(topics):
                for g2 in topics[i + 1:]:
                    level = conflict_levels.get((g1, g2), 0.0)
                    if level <= 0:
                        continue
                    y1, y2 = topic_present.get((t, g1)), topic_present.get((t, g2))
                    if y1 is None or y2 is None:
                        continue
                    z = pulp.LpVariable(
                        f"z_conf_{t}_{g1}_{g2}", lowBound=0, upBound=1, cat="Continuous"
                    )
                    prob += z >= y1 + y2 - 1
                    conflict_terms.append(level * z)

    # --- Soft constraint B2: same-topic parallel penalty --------------
    # Charged per *excess* presentation: n talks of one topic in one timeslot
    # cost (n - 1) units. A plain on/off flag would saturate -- with more
    # presentations in a topic than there are timeslots an overlap is
    # unavoidable by the pigeonhole principle, so the flag would be stuck
    # at 1 and give the solver no reason to spread the topic out.
    topic_excess: Dict[Tuple[str, str], pulp.LpVariable] = {}
    if w_b > 0:
        for t in timeslots:
            for g in topics:
                group_vars = x_by_topic_slot.get((t, g))
                if not group_vars or len(group_vars) < 2:
                    continue
                z_self = pulp.LpVariable(f"z_self_{t}_{g}", lowBound=0, cat="Continuous")
                prob += z_self >= pulp.lpSum(group_vars) - 1
                conflict_terms.append(SELF_CONFLICT_LEVEL * z_self)
                topic_excess[(t, g)] = z_self

        # --- Valid inequality that makes the relaxation usable ---------
        # In any integral schedule a topic's presentations in one timeslot
        # split into "the first one" (y = 1) and "the extras" (z_self), so
        # summing over timeslots gives exactly the topic's presentation count.
        # Every feasible schedule therefore satisfies this and nothing is
        # cut away -- but it is *not* implied for fractional points, where
        # many tiny x values keep both y and z_self near zero. Without it
        # the relaxation ignores both penalty terms almost entirely and
        # CBC cannot close the gap: the 40-presentation conflict-priority
        # instance ran over 80 seconds without proving optimality, and
        # finished in 0.2 seconds with it.
        for g in topics:
            presence_and_excess = []
            for t in timeslots:
                if (t, g) in topic_present:
                    presence_and_excess.append(topic_present[(t, g)])
                if (t, g) in topic_excess:
                    presence_and_excess.append(topic_excess[(t, g)])
            if presence_and_excess:
                prob += (
                    pulp.lpSum(presence_and_excess) >= len(presentations_by_topic[g]),
                    f"topic_coverage_{g}",
                )

    # --- Soft constraint C: empty seats -------------------------------
    if w_c > 0:
        for s in presentations_list:
            pax = presentation_pax[s]
            for r, t in feasible_rt[s]:
                cap = room_cap[r]
                if cap > pax:
                    empty_seat_terms.append((cap - pax) * x[(s, r, t)])

    # --- Normalise each term by the range it can span for this data ---
    # Upper bounds, all computable before solving and all deterministic,
    # so the same workbook always yields the same score.
    n_presentations, n_topics = len(presentations_list), len(topics)
    span_continuity = max(1.0, float(n_presentations - n_topics))
    unique_levels = sum(v for (g1, g2), v in conflict_levels.items() if g1 < g2)
    span_conflict = len(timeslots) * unique_levels + SELF_CONFLICT_LEVEL * (n_presentations - n_topics)
    # For empty seats the meaningful range is what the solver can
    # actually change: the spread between the largest and the smallest
    # room each presentation could be put in. Dividing by the raw maximum
    # would under-weight the term whenever the big rooms are unusable.
    span_empty = 0.0
    for s_id in presentations_list:
        # feasible_rt already accounts for capacity, equipment and
        # blackout slots, so this is the set of rooms the presentation can
        # really land in.
        usable = {room_cap[r] for r, _ in feasible_rt[s_id]}
        if usable:
            span_empty += max(usable) - min(usable)

    print(f" -> Normalisers: continuity {span_continuity:g}, "
          f"conflict {span_conflict:g}, empty seats {span_empty:g}")

    objective_terms: List = []
    if continuity_terms and span_continuity > 0:
        objective_terms.append(
            (SCORE_SCALE * w_a / span_continuity) * pulp.lpSum(continuity_terms))
    if conflict_terms and span_conflict > 0:
        objective_terms.append(
            -(SCORE_SCALE * w_b / span_conflict) * pulp.lpSum(conflict_terms))
    if empty_seat_terms and span_empty > 0:
        objective_terms.append(
            -(SCORE_SCALE * w_c / span_empty) * pulp.lpSum(empty_seat_terms))
    prob += pulp.lpSum(objective_terms)

    # -------------------------------------------------------------
    # Step 5: hard constraints
    # -------------------------------------------------------------
    # [1] Every presentation assigned exactly once.
    for s in presentations_list:
        prob += pulp.lpSum(x_by_presentation[s]) == 1

    # [2] At most one presentation per room per timeslot.
    for r in rooms:
        for t in timeslots:
            vars_here = x_by_room_slot.get((r, t))
            if vars_here:
                prob += pulp.lpSum(vars_here) <= 1

    # [3] A speaker presents at most once per timeslot.
    for (t, speaker), vars_here in x_by_speaker_slot.items():
        if len(vars_here) > 1:
            prob += pulp.lpSum(vars_here) <= 1

    # Capacity / equipment / NG-timeslot are enforced structurally: no
    # variable exists for a combination that would violate them.

    n_variables = len(prob.variables())
    n_constraints = len(prob.constraints)

    # -------------------------------------------------------------
    # Step 6: solve (once per requested candidate)
    # -------------------------------------------------------------
    time_limit_str = f"{time_limit} s" if time_limit is not None else "Unlimited"
    print(f"[4/7] Solving with CBC solver (Time Limit: {time_limit_str})...")
    cbc_info = resolve_cbc()

    difference_threshold = min(
        len(presentations_list),
        max(MIN_SOLUTION_DIFFERENCE,
            math.ceil(SOLUTION_DIFFERENCE_RATIO * len(presentations_list))),
    )

    candidates: List[Candidate] = []
    total_calc_time = 0.0

    for rank in range(1, n_solutions + 1):
        if rank > 1:
            print(f"[4/7] Searching for candidate #{rank} "
                  f"(at least {difference_threshold} assignments must differ)...")

        solver = get_cbc_solver(
            time_limit=time_limit, msg=1, gap_rel=mip_gap if mip_gap > 0 else None
        )
        start_time = time.time()
        status = prob.solve(solver)
        calc_time = time.time() - start_time
        total_calc_time += calc_time

        status_name = pulp.LpStatus.get(status, f"Status_{status}")
        objective_value = prob.objective.value() if prob.objective is not None else None
        solved = (
            status == pulp.LpStatusOptimal
            and getattr(prob, "sol_status", pulp.LpSolutionOptimal) not in (
                pulp.LpSolutionNoSolutionFound,
                pulp.LpSolutionInfeasible,
                pulp.LpSolutionUnbounded,
            )
        )

        print(f" -> Candidate #{rank}: {status_name} (Elapsed: {calc_time:.2f} s)")

        if not solved:
            if rank == 1:
                raise InfeasibleModelError(_t("run.infeasible", status=status_name))
            print(f" -> No further distinct candidate exists. Returning {rank - 1} schedule(s).")
            break

        assignments = _collect_assignments(
            presentations_list, feasible_rt, x, presentation_topic, room_names, presentation_speaker
        )
        _validate_assignments(assignments, presentations_list, speakers_by_name)

        diagnosis = _diagnose(
            assignments, timeslots, timeslot_days, conflict_levels, room_cap, presentation_pax
        )
        candidate = Candidate(
            rank=rank,
            status=status_name,
            solution_quality=_describe_solution_quality(prob, mip_gap),
            objective_score=round(objective_value, 2) if objective_value is not None else 0.0,
            calc_time=round(calc_time, 2),
            assignments=assignments,
            **diagnosis,
        )
        candidate.score_continuity = round(
            SCORE_SCALE * w_a * candidate.continuity_bonuses / span_continuity, 2)
        candidate.score_conflict = round(
            -SCORE_SCALE * w_b * candidate.cross_topic_level_sum / span_conflict, 2
        ) if span_conflict > 0 else 0.0
        candidate.score_same_topic = round(
            -SCORE_SCALE * w_b * SELF_CONFLICT_LEVEL * candidate.same_topic_overlaps
            / span_conflict, 2
        ) if span_conflict > 0 else 0.0
        candidate.score_empty_seats = round(
            -SCORE_SCALE * w_c * candidate.empty_seats / span_empty, 2
        ) if span_empty > 0 else 0.0

        # Independent check: the four components are recomputed from the
        # schedule itself, so if they disagree with the solver's own
        # objective something in the model or the reporting is wrong.
        recomputed = (candidate.score_continuity + candidate.score_conflict
                      + candidate.score_same_topic + candidate.score_empty_seats)
        if abs(recomputed - candidate.objective_score) > 0.5:
            print(_t("run.breakdown_mismatch", recomputed=round(recomputed, 2),
                    reported=candidate.objective_score))

        candidates.append(candidate)

        if rank < n_solutions:
            chosen = [x[(a["Presentation_ID"], a["Room_ID"], a["TimeSlot_ID"])] for a in assignments]
            prob += (
                pulp.lpSum(chosen) <= len(chosen) - difference_threshold,
                f"differ_from_candidate_{rank}",
            )

    print("[5/7] Hard constraints validated for every returned schedule.")

    # -------------------------------------------------------------
    # Step 7: reports
    # -------------------------------------------------------------
    print("[6/7] Building timetable matrix and summary reports...")
    sheets_out = _build_output_sheets(
        candidates=candidates,
        rooms=rooms,
        room_names=room_names,
        timeslots=timeslots,
        settings=dict(
            weight_continuity=weight_continuity,
            weight_conflict=weight_conflict,
            weight_empty_seats=weight_empty_seats,
            mip_gap=mip_gap,
            time_limit=time_limit,
            solutions_requested=n_solutions,
            n_variables=n_variables,
            n_constraints=n_constraints,
            use_equipment=use_equipment,
            use_speaker=bool(speakers_by_name),
            total_calc_time=total_calc_time,
            difference_threshold=difference_threshold,
            span_continuity=span_continuity,
            span_conflict=span_conflict,
            span_empty=span_empty,
            cbc_source=cbc_info.source,
            cbc_version=cbc_info.version,
        ),
    )

    print(f"[7/7] Writing output Excel file to: {output_file}...")
    final_output_path = write_workbook_safe(output_file, sheets_out)
    print(f"\n[Success] Timetable optimization results saved to '{final_output_path}'.\n")

    best = candidates[0]
    result = OptimizationResult(
        status=best.status,
        objective_score=best.objective_score,
        calc_time=best.calc_time,
        output_file=os.path.abspath(final_output_path),
        assigned_count=len(best.assignments),
        model_variables=n_variables,
        model_constraints=n_constraints,
        solution_quality=best.solution_quality,
        total_calc_time=round(total_calc_time, 2),
        solutions_requested=n_solutions,
        solutions_returned=len(candidates),
        candidates=[c.as_dict() for c in candidates],
    )
    return result.as_dict()


# =====================================================================
# 7. Solution post-processing
# =====================================================================

def _collect_assignments(
    presentations_list, feasible_rt, x, presentation_topic, room_names, presentation_speaker
) -> List[Dict]:
    assignments: List[Dict] = []
    for s in presentations_list:
        chosen = [
            (r, t)
            for (r, t) in feasible_rt[s]
            if pulp.value(x[(s, r, t)]) is not None and pulp.value(x[(s, r, t)]) > 0.5
        ]
        for r, t in chosen:
            assignments.append(
                {
                    "TimeSlot_ID": t,
                    "Room_ID": r,
                    "Room": f"{r}:{room_names[r]}",
                    "Presentation_ID": s,
                    "Topic": presentation_topic[s],
                    "Speaker": presentation_speaker.get(s, ""),
                }
            )
    return assignments


def _validate_assignments(assignments: List[Dict], presentations_list, speakers_by_name) -> None:
    """Re-check the hard constraints against the returned schedule.

    The solver should never violate them, but checking the *output*
    rather than trusting the model is what turns a silent modelling
    mistake into a visible error.
    """
    errors: List[str] = []
    per_presentation: Dict[str, List[Tuple[str, str]]] = {}
    for item in assignments:
        per_presentation.setdefault(item["Presentation_ID"], []).append(
            (item["Room_ID"], item["TimeSlot_ID"])
        )

    for s in presentations_list:
        placed = per_presentation.get(s, [])
        if not placed:
            errors.append(f"Unassigned Error: Presentation {s} was not assigned to any room/slot.")
        elif len(placed) > 1:
            errors.append(f"Double-Assignment Error: Presentation {s} was assigned to {placed}.")

    seen_room_slot: Dict[Tuple[str, str], str] = {}
    for item in assignments:
        key = (item["Room_ID"], item["TimeSlot_ID"])
        if key in seen_room_slot:
            errors.append(
                f"Room Clash Error: {item['Room_ID']} at {item['TimeSlot_ID']} holds both "
                f"{seen_room_slot[key]} and {item['Presentation_ID']}."
            )
        else:
            seen_room_slot[key] = item["Presentation_ID"]

    if speakers_by_name:
        seen_speaker_slot: Dict[Tuple[str, str], str] = {}
        for item in assignments:
            speaker = item.get("Speaker")
            if not speaker:
                continue
            key = (speaker, item["TimeSlot_ID"])
            if key in seen_speaker_slot:
                errors.append(
                    f"Speaker Clash Error: {speaker} presents both "
                    f"{seen_speaker_slot[key]} and {item['Presentation_ID']} at {item['TimeSlot_ID']}."
                )
            else:
                seen_speaker_slot[key] = item["Presentation_ID"]

    if errors:
        message = f"Validation failed with {len(errors)} error(s):\n" + "\n".join(
            f"  - {e}" for e in errors[:10]
        )
        print(f"[Error] {message}")
        raise ValueError(message)


def _diagnose(assignments, timeslots, timeslot_days, conflict_levels,
              room_cap, presentation_pax) -> Dict[str, float]:
    """Measure what the schedule actually achieved.

    Counted from the assignment itself rather than read back from the
    model, so these double as an independent check on the objective --
    and they are the numbers a user needs in order to judge a schedule.
    The old code printed "No conflicts found" unconditionally.
    """
    by_slot: Dict[str, List[Dict]] = {}
    by_room_slot: Dict[Tuple[str, str], Dict] = {}
    for item in assignments:
        by_slot.setdefault(item["TimeSlot_ID"], []).append(item)
        by_room_slot[(item["Room_ID"], item["TimeSlot_ID"])] = item

    cross_pairs = 0
    cross_level_sum = 0.0
    same_excess = 0
    for slot_items in by_slot.values():
        counts: Dict[str, int] = {}
        for item in slot_items:
            counts[item["Topic"]] = counts.get(item["Topic"], 0) + 1
        present = list(counts)
        for i, g1 in enumerate(present):
            for g2 in present[i + 1:]:
                level = conflict_levels.get((g1, g2), 0.0)
                if level > 0:
                    cross_pairs += 1
                    cross_level_sum += level
        same_excess += sum(max(0, n - 1) for n in counts.values())

    continuity = 0
    rooms_used = {item["Room_ID"] for item in assignments}
    for room in rooms_used:
        for idx in range(len(timeslots) - 1):
            t_curr, t_next = timeslots[idx], timeslots[idx + 1]
            if timeslot_days.get(t_curr) != timeslot_days.get(t_next):
                continue
            a = by_room_slot.get((room, t_curr))
            b = by_room_slot.get((room, t_next))
            if a and b and a["Topic"] == b["Topic"]:
                continuity += 1

    empty_seats = 0.0
    for item in assignments:
        cap = room_cap.get(item["Room_ID"], 0.0)
        pax = presentation_pax.get(item["Presentation_ID"], 0.0)
        empty_seats += max(0.0, cap - pax)

    return {
        "continuity_bonuses": continuity,
        "cross_topic_conflicts": cross_pairs,
        "cross_topic_level_sum": cross_level_sum,
        "same_topic_overlaps": same_excess,
        "empty_seats": empty_seats,
    }


def _timetable_frame(assignments, rooms, room_names, timeslots) -> pd.DataFrame:
    room_columns = [f"{r}:{room_names[r]}" for r in rooms]
    matrix = {col: {t: "---" for t in timeslots} for col in room_columns}
    for item in assignments:
        matrix[item["Room"]][item["TimeSlot_ID"]] = f"{item['Presentation_ID']} ({item['Topic']})"
    return (
        pd.DataFrame(matrix, index=timeslots)
        .reset_index()
        .rename(columns={"index": "TimeSlot_ID"})
    )


def _build_output_sheets(candidates, rooms, room_names, timeslots, settings) -> Dict[str, pd.DataFrame]:
    single = len(candidates) == 1
    sheets: Dict[str, pd.DataFrame] = {}

    list_rows: List[Dict] = []
    for cand in candidates:
        frame = _timetable_frame(cand.assignments, rooms, room_names, timeslots)
        # A single candidate keeps the historical sheet name so existing
        # workflows and downstream scripts keep working.
        sheets["Result_Timetable" if single else f"Result_Timetable_{cand.rank}"] = frame
        for item in cand.assignments:
            row = {} if single else {"Rank": cand.rank}
            row.update(
                {
                    "TimeSlot_ID": item["TimeSlot_ID"],
                    "Room": item["Room"],
                    "Presentation_ID": item["Presentation_ID"],
                    "Topic": item["Topic"],
                }
            )
            if settings["use_speaker"]:
                row["Speaker"] = item.get("Speaker", "")
            list_rows.append(row)

    sheets["Result_List"] = pd.DataFrame(list_rows)

    best = candidates[0]
    metrics: List[Tuple[str, object]] = [
        ("Total_Objective_Score", best.objective_score),
        ("Solver_Status", best.status),
        ("Solution_Quality", best.solution_quality),
        ("Computation_Time_Sec", best.calc_time),
        # The four components below sum to Total_Objective_Score.
        ("Score_Continuity_A", best.score_continuity),
        ("Score_Conflict_B1", best.score_conflict),
        ("Score_SameTopic_B2", best.score_same_topic),
        ("Score_EmptySeats_C", best.score_empty_seats),
        ("Continuity_Bonuses_A", best.continuity_bonuses),
        ("Cross_Topic_Conflicts_B1", best.cross_topic_conflicts),
        ("Same_Topic_Overlaps_B2", best.same_topic_overlaps),
        ("Empty_Seats_Total", best.empty_seats),
        ("Weight_Continuity", settings["weight_continuity"]),
        ("Weight_Conflict", settings["weight_conflict"]),
        ("Weight_EmptySeats", settings["weight_empty_seats"]),
        ("SelfConflict_Level", SELF_CONFLICT_LEVEL),
        # Without these the score cannot be reconstructed from the counts.
        ("Normalizer_Continuity", settings["span_continuity"]),
        ("Normalizer_Conflict", settings["span_conflict"]),
        ("Normalizer_EmptySeats", settings["span_empty"]),
        ("Score_Scale", SCORE_SCALE),
        ("MIP_Gap_Setting", settings["mip_gap"]),
        ("Time_Limit_Setting", settings["time_limit"] if settings["time_limit"] else "Unlimited"),
        ("Equipment_Constraint", "Enabled" if settings["use_equipment"] else "Disabled"),
        ("Speaker_Constraint", "Enabled" if settings["use_speaker"] else "Disabled"),
        ("Solutions_Requested", settings["solutions_requested"]),
        ("Solutions_Returned", len(candidates)),
        ("Model_Variables", settings["n_variables"]),
        ("Model_Constraints", settings["n_constraints"]),
        ("Total_Computation_Time_Sec", round(settings["total_calc_time"], 2)),
        # Which CBC produced this, since the one used depends on the machine.
        ("Solver_CBC_Source", settings.get("cbc_source", "")),
        ("Solver_CBC_Version", settings.get("cbc_version", "")),
    ]

    if not single:
        metrics.append(("Min_Assignment_Difference", settings["difference_threshold"]))
        for cand in candidates[1:]:
            metrics.extend(
                [
                    (f"Objective_Score_{cand.rank}", cand.objective_score),
                    (f"Computation_Time_Sec_{cand.rank}", cand.calc_time),
                    (f"Continuity_Bonuses_A_{cand.rank}", cand.continuity_bonuses),
                    (f"Cross_Topic_Conflicts_B1_{cand.rank}", cand.cross_topic_conflicts),
                    (f"Same_Topic_Overlaps_B2_{cand.rank}", cand.same_topic_overlaps),
                ]
            )

    sheets["Simulation_Report"] = pd.DataFrame(
        {"Metric": [m for m, _ in metrics], "Value": [v for _, v in metrics]}
    )
    return sheets
