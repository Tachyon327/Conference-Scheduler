"""
Unit, regression and property tests for scheduler_engine.

Run with:  pytest -q
(requires: pandas, openpyxl, pulp, pytest)

These tests avoid importing scheduler_gui / tkinter, so they run in
headless environments without Tk installed.

A note on the property tests near the bottom of this file: the 2026-09
revision fixed a conflict-penalty formulation that had been inert -- the
term could never take a non-zero value, so schedules with many clashing
parallel presentations were reported as optimal with a clean bill of health.
Every test passed throughout, because the suite only ever checked that
the code ran and that one fixed instance produced one fixed number. The
tests below assert *behaviour* instead: that turning a weight up changes
the objective, and that the solver avoids what it is told to avoid.
Those survive refactoring and catch a dead term.
"""
import os
import sys

import pandas as pd
import pytest

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

import scheduler_engine as engine
import scheduler_messages as msg


@pytest.fixture(autouse=True)
def _restore_language():
    """Every test starts in English and cannot leak a language change."""
    before = msg.get_language()
    msg.set_language("en")
    yield
    msg.set_language(before)


# ---------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------

def _write_workbook(path, sheets: dict):
    with pd.ExcelWriter(path, engine="openpyxl") as writer:
        for name, df in sheets.items():
            df.to_excel(writer, sheet_name=name, index=False)


def _valid_sheets():
    return {
        "Rooms": pd.DataFrame({"Room_ID": ["R1"], "Room_Name": ["A"], "Capacity": [10], "Screen": [1]}),
        "TimeSlots": pd.DataFrame({"TimeSlot_ID": ["T1"], "Day": ["D1"]}),
        "Presentations": pd.DataFrame(
            {"Presentation_ID": ["S1"], "Topic": ["A"], "Expected_Pax": [5],
             "Needs_Screen": [0], "NG_TimeSlot_ID": [""]}
        ),
        "Conflicts": pd.DataFrame({"Topic1": [], "Topic2": [], "Conflict_Level": []}),
    }


def _rooms(n: int, capacity: int = 100) -> pd.DataFrame:
    return pd.DataFrame(
        {
            "Room_ID": [f"R{i}" for i in range(1, n + 1)],
            "Room_Name": [f"Room {i}" for i in range(1, n + 1)],
            "Capacity": [capacity] * n,
        }
    )


def _timeslots(n: int) -> pd.DataFrame:
    return pd.DataFrame(
        {
            "TimeSlot_ID": [f"T{i}" for i in range(1, n + 1)],
            "Day": ["D1"] * n,
            "Start_Time": [f"{8 + i:02d}:00" for i in range(1, n + 1)],
        }
    )


def _presentations(topics, pax: int = 10, speakers=None) -> pd.DataFrame:
    data = {
        "Presentation_ID": [f"S{i}" for i in range(1, len(topics) + 1)],
        "Topic": list(topics),
        "Expected_Pax": [pax] * len(topics),
    }
    if speakers is not None:
        data["Speaker"] = list(speakers)
    return pd.DataFrame(data)


def _run(tmp_path, sheets, name="case", **kwargs):
    in_path = tmp_path / f"{name}.xlsx"
    out_path = tmp_path / f"{name}_out.xlsx"
    _write_workbook(in_path, sheets)
    return engine.run_optimization(str(in_path), str(out_path), **kwargs)


# ---------------------------------------------------------------------
# parse_ng_slots
# ---------------------------------------------------------------------

@pytest.mark.parametrize(
    "value, expected",
    [
        (None, set()), (float("nan"), set()), ("", set()), ("   ", set()),
        ("T1", {"T1"}), ("T1,T2", {"T1", "T2"}), ("T1, T2", {"T1", "T2"}),
        ("T1、T2", {"T1", "T2"}), ("T1，T2", {"T1", "T2"}), ("T1  T2", {"T1", "T2"}),
    ],
)
def test_parse_ng_slots(value, expected):
    assert engine.parse_ng_slots(value) == expected


# ---------------------------------------------------------------------
# truthy
# ---------------------------------------------------------------------

@pytest.mark.parametrize(
    "value, expected",
    [
        (1, True), (1.0, True), (True, True),
        ("1", True), ("TRUE", True), ("True", True), ("Yes", True), ("y", True), ("有", True),
        (0, False), (0.0, False), (False, False),
        ("0", False), ("FALSE", False), ("no", False), ("N", False), ("", False),
        ("-", False), ("×", False), ("無", False), (None, False), (float("nan"), False),
    ],
)
def test_truthy(value, expected):
    """A workbook written with TRUE/Yes used to lose the constraint silently."""
    assert engine.truthy(value) is expected


# ---------------------------------------------------------------------
# sort_timeslots
# ---------------------------------------------------------------------

def test_sort_timeslots_orders_rows_by_clock_within_each_day():
    """Adjacency for the continuity bonus must not depend on row order."""
    df = pd.DataFrame(
        {
            "TimeSlot_ID": ["T3", "T1", "T2", "T4"],
            "Day": ["D1", "D1", "D1", "D2"],
            "Start_Time": ["13:00", "09:00", "10:45", "09:00"],
        }
    )
    assert engine.sort_timeslots(df)["TimeSlot_ID"].tolist() == ["T1", "T2", "T3", "T4"]


def test_sort_timeslots_keeps_day_order_of_first_appearance():
    df = pd.DataFrame(
        {
            "TimeSlot_ID": ["A1", "B1", "A2"],
            "Day": ["Day10", "Day2", "Day10"],
            "Start_Time": ["10:00", "09:00", "09:00"],
        }
    )
    # Day10 appears first so it stays first. A plain string sort would put
    # "Day10" before "Day2" for the wrong reason; a numeric sort would
    # reorder the user's days.
    assert engine.sort_timeslots(df)["TimeSlot_ID"].tolist() == ["A2", "A1", "B1"]


def test_sort_timeslots_falls_back_to_sheet_order_when_times_unparsable():
    df = pd.DataFrame(
        {
            "TimeSlot_ID": ["T1", "T2", "T3"],
            "Day": ["D1", "D1", "D1"],
            "Start_Time": ["morning", "afternoon", "evening"],
        }
    )
    assert engine.sort_timeslots(df)["TimeSlot_ID"].tolist() == ["T1", "T2", "T3"]


# ---------------------------------------------------------------------
# Workbook structure
# ---------------------------------------------------------------------

def test_validate_input_workbook_missing_file(tmp_path):
    with pytest.raises(FileNotFoundError):
        engine.validate_input_workbook(str(tmp_path / "does_not_exist.xlsx"))


def test_validate_input_workbook_missing_required_sheet(tmp_path):
    sheets = _valid_sheets()
    del sheets["Rooms"]
    path = tmp_path / "missing_sheet.xlsx"
    _write_workbook(path, sheets)
    with pytest.raises(engine.InputValidationError, match="Rooms"):
        engine.validate_input_workbook(str(path))


def test_validate_input_workbook_missing_column(tmp_path):
    sheets = _valid_sheets()
    sheets["Rooms"] = sheets["Rooms"].drop(columns=["Capacity"])
    path = tmp_path / "missing_col.xlsx"
    _write_workbook(path, sheets)
    with pytest.raises(engine.InputValidationError, match="Capacity"):
        engine.validate_input_workbook(str(path))


def test_validate_input_workbook_ok(tmp_path):
    path = tmp_path / "ok.xlsx"
    _write_workbook(path, _valid_sheets())
    result = engine.validate_input_workbook(str(path))
    assert set(result.keys()) == set(engine.ALL_SHEETS)


def test_conflicts_sheet_may_be_absent(tmp_path):
    sheets = _valid_sheets()
    del sheets["Conflicts"]
    path = tmp_path / "no_conflicts.xlsx"
    _write_workbook(path, sheets)
    result = engine.validate_input_workbook(str(path))
    assert set(result.keys()) == set(engine.ALL_SHEETS)
    assert result["Conflicts"].empty


def test_present_but_malformed_optional_sheet_is_reported(tmp_path):
    """A misspelt header is a mistake, not a decision to disable the sheet."""
    sheets = _valid_sheets()
    sheets["Conflicts"] = pd.DataFrame({"Topic1": ["A"], "Topic2": ["B"], "Level": [1]})
    path = tmp_path / "bad_optional.xlsx"
    _write_workbook(path, sheets)
    with pytest.raises(engine.InputValidationError, match="Conflict_Level"):
        engine.validate_input_workbook(str(path))


def test_extra_sheets_are_ignored(tmp_path):
    """A README sheet shipped with the template must not upset anything."""
    sheets = {"README_EN": pd.DataFrame({"Description": ["What each column means..."]}),
              "README_JP": pd.DataFrame({"説明": ["列の意味..."]})}
    sheets.update(_valid_sheets())
    path = tmp_path / "with_readme.xlsx"
    _write_workbook(path, sheets)
    result = engine.validate_input_workbook(str(path))
    assert set(result.keys()) == set(engine.ALL_SHEETS)


def test_runs_with_only_required_sheets_and_columns(tmp_path):
    """Three sheets, seven columns -- the documented minimum input."""
    sheets = {
        "Rooms": pd.DataFrame({"Room_ID": ["R1", "R2"], "Capacity": [50, 50]}),
        "TimeSlots": pd.DataFrame({"TimeSlot_ID": ["T1", "T2"], "Day": ["D1", "D1"]}),
        "Presentations": pd.DataFrame(
            {"Presentation_ID": ["S1", "S2"], "Topic": ["A", "B"], "Expected_Pax": [10, 10]}
        ),
    }
    result = _run(tmp_path, sheets, name="minimal")
    assert result["status"] == "Optimal"
    assert result["assigned_count"] == 2


# ---------------------------------------------------------------------
# Data-level validation
# ---------------------------------------------------------------------

@pytest.mark.parametrize("language, expected", [("ja", "重複"), ("en", "duplicates")])
def test_duplicate_presentation_id_is_reported(tmp_path, language, expected):
    """Used to produce a wrong schedule with no error at all."""
    msg.set_language(language)
    sheets = _valid_sheets()
    sheets["Presentations"] = pd.DataFrame(
        {"Presentation_ID": ["S1", "S1"], "Topic": ["A", "B"], "Expected_Pax": [5, 5]}
    )
    sheets["Rooms"] = _rooms(2)
    sheets["TimeSlots"] = _timeslots(2)
    path = tmp_path / "dup.xlsx"
    _write_workbook(path, sheets)
    with pytest.raises(engine.InputValidationError, match=expected):
        engine.validate_input_workbook(str(path))


def test_duplicate_room_id_is_reported(tmp_path):
    sheets = _valid_sheets()
    sheets["Rooms"] = pd.DataFrame({"Room_ID": ["R1", "R1"], "Capacity": [10, 10]})
    path = tmp_path / "duproom.xlsx"
    _write_workbook(path, sheets)
    with pytest.raises(engine.InputValidationError, match="Room_ID"):
        engine.validate_input_workbook(str(path))


@pytest.mark.parametrize("language, expected", [("ja", "3行目"), ("en", "row 3")])
def test_non_numeric_capacity_is_reported_with_the_row(tmp_path, language, expected):
    """Used to surface as a Python type error from inside the model."""
    msg.set_language(language)
    sheets = _valid_sheets()
    sheets["Rooms"] = pd.DataFrame(
        {"Room_ID": ["R1", "R2"], "Capacity": [10, "約120"]}
    )
    path = tmp_path / "badcap.xlsx"
    _write_workbook(path, sheets)
    with pytest.raises(engine.InputValidationError, match=expected):
        engine.validate_input_workbook(str(path))


def test_blank_expected_pax_is_reported(tmp_path):
    sheets = _valid_sheets()
    sheets["Presentations"] = pd.DataFrame(
        {"Presentation_ID": ["S1", "S2"], "Topic": ["A", "B"], "Expected_Pax": [5, None]}
    )
    sheets["Rooms"] = _rooms(2)
    sheets["TimeSlots"] = _timeslots(2)
    path = tmp_path / "blankpax.xlsx"
    _write_workbook(path, sheets)
    with pytest.raises(engine.InputValidationError, match="Expected_Pax"):
        engine.validate_input_workbook(str(path))


@pytest.mark.parametrize("language, expected", [("ja", "組み合わせ"), ("en", "room-slots")])
def test_more_presentations_than_room_slots_is_reported(tmp_path, language, expected):
    msg.set_language(language)
    sheets = {
        "Rooms": _rooms(1),
        "TimeSlots": _timeslots(2),
        "Presentations": _presentations(["A", "B", "C"]),
    }
    path = tmp_path / "toomany.xlsx"
    _write_workbook(path, sheets)
    with pytest.raises(engine.InputValidationError, match=expected):
        engine.validate_input_workbook(str(path))


def test_unknown_ng_timeslot_is_reported(tmp_path):
    presentations = _presentations(["A", "B"])
    presentations["NG_TimeSlot_ID"] = ["T9", ""]
    sheets = {"Rooms": _rooms(2), "TimeSlots": _timeslots(2), "Presentations": presentations}
    path = tmp_path / "badng.xlsx"
    _write_workbook(path, sheets)
    with pytest.raises(engine.InputValidationError, match="T9"):
        engine.validate_input_workbook(str(path))


def test_all_problems_are_reported_at_once(tmp_path):
    """Fixing a spreadsheet one error per run is miserable."""
    sheets = {
        "Rooms": pd.DataFrame({"Room_ID": ["R1", "R1"], "Capacity": [10, "x"]}),
        "TimeSlots": _timeslots(2),
        "Presentations": pd.DataFrame(
            {"Presentation_ID": ["S1", None], "Topic": ["A", "B"], "Expected_Pax": [5, 5]}
        ),
    }
    path = tmp_path / "many.xlsx"
    _write_workbook(path, sheets)
    with pytest.raises(engine.InputValidationError) as excinfo:
        engine.validate_input_workbook(str(path))
    message = str(excinfo.value)
    assert "Room_ID" in message and "Capacity" in message and "Presentation_ID" in message


# ---------------------------------------------------------------------
# summarize_input
# ---------------------------------------------------------------------

def test_summarize_input_reports_the_shape(tmp_path):
    sheets = {
        "Rooms": _rooms(3),
        "TimeSlots": _timeslots(4),
        "Presentations": _presentations(["A", "A", "B"], speakers=["x", "y", "x"]),
        "Conflicts": pd.DataFrame({"Topic1": ["A"], "Topic2": ["B"], "Conflict_Level": [1]}),
    }
    path = tmp_path / "shape.xlsx"
    _write_workbook(path, sheets)
    info = engine.summarize_input(str(path))
    assert info == {
        "presentations": 3, "rooms": 3, "timeslots": 4, "topics": 2,
        "speakers": 2, "conflict_pairs": 1, "slots_available": 12,
    }


def test_summarize_input_raises_on_a_bad_file(tmp_path):
    sheets = _valid_sheets()
    del sheets["Presentations"]
    path = tmp_path / "bad.xlsx"
    _write_workbook(path, sheets)
    with pytest.raises(engine.InputValidationError):
        engine.summarize_input(str(path))


# ---------------------------------------------------------------------
# save_excel_safe
# ---------------------------------------------------------------------

def test_save_excel_safe_normal(tmp_path):
    out = tmp_path / "out.xlsx"
    df = pd.DataFrame({"a": [1]})
    result_path = engine.save_excel_safe(str(out), df, df, df)
    assert result_path == str(out)
    assert os.path.exists(result_path)


def test_save_excel_safe_falls_back_when_locked(tmp_path, monkeypatch):
    out = tmp_path / "locked.xlsx"
    df = pd.DataFrame({"a": [1]})

    call_count = {"n": 0}
    real_writer = pd.ExcelWriter

    class FlakyWriter:
        def __init__(self, path, *args, **kwargs):
            call_count["n"] += 1
            if call_count["n"] == 1:
                raise PermissionError("simulated lock")
            self._inner = real_writer(path, *args, **kwargs)

        def __enter__(self):
            return self._inner.__enter__()

        def __exit__(self, *exc):
            return self._inner.__exit__(*exc)

    monkeypatch.setattr(pd, "ExcelWriter", FlakyWriter)
    result_path = engine.save_excel_safe(str(out), df, df, df)
    assert result_path != str(out)
    assert os.path.exists(result_path)


# ---------------------------------------------------------------------
# Full optimization regression tests
# ---------------------------------------------------------------------

def _write_small_instance(path):
    sheets = {
        "Rooms": pd.DataFrame(
            {"Room_ID": ["R1", "R2"], "Room_Name": ["Room A", "Room B"],
             "Capacity": [50, 20], "Screen": [1, 0]}
        ),
        "TimeSlots": pd.DataFrame({"TimeSlot_ID": ["T1", "T2", "T3"], "Day": ["D1", "D1", "D1"]}),
        "Presentations": pd.DataFrame(
            {
                "Presentation_ID": ["S1", "S2", "S3", "S4"],
                "Topic": ["A", "A", "B", "B"],
                "Expected_Pax": [30, 40, 10, 15],
                "Needs_Screen": [1, 0, 0, 0],
                "NG_TimeSlot_ID": ["", "", "T1", ""],
            }
        ),
        "Conflicts": pd.DataFrame({"Topic1": ["A"], "Topic2": ["B"], "Conflict_Level": [1.0]}),
    }
    _write_workbook(path, sheets)


def test_run_optimization_small_instance_is_feasible_and_optimal(tmp_path):
    input_path = tmp_path / "small.xlsx"
    output_path = tmp_path / "small_out.xlsx"
    _write_small_instance(input_path)

    result = engine.run_optimization(str(input_path), str(output_path))

    assert result["status"] == "Optimal"
    assert result["solution_quality"] == "Optimal (proven)"
    assert result["assigned_count"] == 4
    # Known-good objective value for this fixed instance, checkable by
    # hand. The normalisers are 2 / 9 / 60 (see the test below), and the
    # schedule achieves 2 continuity bonuses, one level-1 cross-topic
    # clash, no same-topic overlap and 45 empty seats:
    #     100*2/2  -  100*1/9  -  0  -  100*45/60
    #   =   100.00 -    11.11  -  0  -    75.00   =  13.89
    assert result["objective_score"] == pytest.approx(13.89, abs=0.01)
    assert os.path.exists(result["output_file"])

    out_sheets = pd.read_excel(result["output_file"], sheet_name=None)
    assert set(out_sheets.keys()) == {"Result_Timetable", "Result_List", "Simulation_Report"}
    assert len(out_sheets["Result_List"]) == 4
    assert sorted(out_sheets["Result_List"]["Presentation_ID"].astype(str)) == ["S1", "S2", "S3", "S4"]


def test_run_optimization_reports_infeasible_presentation_clearly(tmp_path):
    sheets = {
        "Rooms": pd.DataFrame({"Room_ID": ["R1"], "Capacity": [10], "Screen": [0]}),
        "TimeSlots": pd.DataFrame({"TimeSlot_ID": ["T1"], "Day": ["D1"]}),
        "Presentations": pd.DataFrame(
            {"Presentation_ID": ["S1"], "Topic": ["A"], "Expected_Pax": [999], "Needs_Screen": [0]}
        ),
    }
    with pytest.raises(engine.InfeasibleModelError, match="S1"):
        _run(tmp_path, sheets, name="infeasible")


def test_mip_gap_reaches_the_solver(monkeypatch, tmp_path):
    input_path = tmp_path / "small.xlsx"
    output_path = tmp_path / "small_out.xlsx"
    _write_small_instance(input_path)

    seen = {}
    real_get_cbc_solver = engine.get_cbc_solver

    def spy(time_limit=None, msg=1, gap_rel=None):
        seen["gap_rel"] = gap_rel
        seen["time_limit"] = time_limit
        return real_get_cbc_solver(time_limit=time_limit, msg=msg, gap_rel=gap_rel)

    monkeypatch.setattr(engine, "get_cbc_solver", spy)
    engine.run_optimization(str(input_path), str(output_path), mip_gap=0.05, time_limit=120)
    assert seen["gap_rel"] == pytest.approx(0.05)
    assert seen["time_limit"] == 120


# ---------------------------------------------------------------------
# Regression tests for the 2026-09 conflict-penalty defects
# ---------------------------------------------------------------------

def test_all_weights_zero_still_solves(tmp_path):
    """An empty objective is a legal request, not a failure.

    With every weight at zero the objective expression is empty; PuLP
    then reports its value as ``None`` and the engine used to mistake
    that for "no feasible solution", which made an all-terms-off
    ablation baseline impossible to run.
    """
    sheets = {"Rooms": _rooms(2), "TimeSlots": _timeslots(2), "Presentations": _presentations(["A", "B"])}
    result = _run(tmp_path, sheets, name="zero",
                  weight_continuity=0, weight_conflict=0, weight_empty_seats=0)
    assert result["status"] == "Optimal"
    assert result["objective_score"] == pytest.approx(0.0)
    assert result["assigned_count"] == 2


def test_three_same_topic_presentations_in_one_timeslot_is_feasible(tmp_path):
    """Regression: the old formulation compared raw presentation counts.

    ``z >= (count of g1) + (count of g2) - 1`` with a binary ``z`` becomes
    unsatisfiable as soon as three presentations share a timeslot, so a single
    topic with three parallel talks made the whole model infeasible --
    even though no cross-topic conflict had occurred at all.
    """
    sheets = {
        "Rooms": _rooms(3),
        "TimeSlots": _timeslots(1),
        "Presentations": _presentations(["A", "A", "A"]),
        "Conflicts": pd.DataFrame({"Topic1": ["A"], "Topic2": ["B"], "Conflict_Level": [1.0]}),
    }
    result = _run(tmp_path, sheets, name="triple",
                  weight_continuity=0, weight_conflict=1, weight_empty_seats=0)
    assert result["status"] == "Optimal"
    assert result["assigned_count"] == 3
    assert result["candidates"][0]["same_topic_overlaps"] == 2


def test_same_topic_presentations_are_not_charged_as_cross_topic_conflicts(tmp_path):
    """Regression: two presentations of one topic used to trigger the B1 penalty.

    Topic A is pinned to T1 and topic B is free to take T2, so no
    cross-topic conflict occurs; only the same-topic overlap is charged.
    """
    presentations = pd.DataFrame(
        {
            "Presentation_ID": ["S1", "S2", "S3"],
            "Topic": ["A", "A", "B"],
            "Expected_Pax": [10, 10, 10],
            "NG_TimeSlot_ID": ["T2", "T2", ""],
        }
    )
    sheets = {
        "Rooms": _rooms(2),
        "TimeSlots": _timeslots(2),
        "Presentations": presentations,
        "Conflicts": pd.DataFrame({"Topic1": ["A"], "Topic2": ["B"], "Conflict_Level": [1.0]}),
    }
    result = _run(tmp_path, sheets, name="no_false_b1",
                  weight_continuity=0, weight_conflict=1, weight_empty_seats=0)
    best = result["candidates"][0]
    assert best["cross_topic_conflicts"] == 0
    assert best["same_topic_overlaps"] == 1
    # span_conflict = timeslots*levels + 3*(presentations-topics) = 2*1 + 3*1 = 5,
    # so one same-topic overlap costs 100 * 3/5 = 60.
    assert result["objective_score"] == pytest.approx(-60.0, abs=0.01)


# ---------------------------------------------------------------------
# Property tests: the soft constraints must actually do something
# ---------------------------------------------------------------------

def _forced_cross_conflict_sheets():
    """Two conflicting topics with nowhere to go but the same timeslot."""
    return {
        "Rooms": _rooms(2),
        "TimeSlots": _timeslots(1),
        "Presentations": _presentations(["A", "B"]),
        "Conflicts": pd.DataFrame({"Topic1": ["A"], "Topic2": ["B"], "Conflict_Level": [1.0]}),
    }


def test_cross_topic_conflict_penalty_is_actually_charged(tmp_path):
    """The defect this suite failed to catch: B1 was identically zero.

    An unavoidable clash must cost something, and raising the weight must
    cost proportionally more. If the term is inert, all three runs return 0.
    """
    common = dict(weight_continuity=0, weight_empty_seats=0)
    off = _run(tmp_path, _forced_cross_conflict_sheets(), name="b1_off",
               weight_conflict=0, **common)
    on = _run(tmp_path, _forced_cross_conflict_sheets(), name="b1_on",
              weight_conflict=1, **common)
    strong = _run(tmp_path, _forced_cross_conflict_sheets(), name="b1_strong",
                  weight_conflict=2, **common)

    assert off["candidates"][0]["cross_topic_conflicts"] == 1
    assert off["objective_score"] == pytest.approx(0.0)
    assert on["objective_score"] < off["objective_score"]
    assert strong["objective_score"] == pytest.approx(2 * on["objective_score"])


def test_same_topic_penalty_is_actually_charged(tmp_path):
    """B2 must fire without any Conflicts sheet at all."""
    sheets = {"Rooms": _rooms(2), "TimeSlots": _timeslots(1), "Presentations": _presentations(["A", "A"])}
    common = dict(weight_continuity=0, weight_empty_seats=0)
    off = _run(tmp_path, sheets, name="b2_off", weight_conflict=0, **common)
    on = _run(tmp_path, sheets, name="b2_on", weight_conflict=1, **common)

    assert off["candidates"][0]["same_topic_overlaps"] == 1
    assert off["objective_score"] == pytest.approx(0.0)
    # No Conflicts sheet, so span_conflict = 3*(2-1) = 3 and the single
    # overlap costs the whole scale: 100 * 3/3.
    assert on["objective_score"] == pytest.approx(-engine.SCORE_SCALE, abs=0.01)


def test_conflict_penalty_steers_the_schedule_away_from_clashes(tmp_path):
    """With the penalty on and an escape available, the solver must take it."""
    sheets = {
        "Rooms": _rooms(2),
        "TimeSlots": _timeslots(3),
        "Presentations": _presentations(["A", "A", "B"]),
        "Conflicts": pd.DataFrame({"Topic1": ["A"], "Topic2": ["B"], "Conflict_Level": [3.0]}),
    }
    result = _run(tmp_path, sheets, name="steer",
                  weight_continuity=0, weight_conflict=1, weight_empty_seats=0)
    best = result["candidates"][0]
    assert best["cross_topic_conflicts"] == 0
    assert best["same_topic_overlaps"] == 0
    assert result["objective_score"] == pytest.approx(0.0)


def test_continuity_bonus_is_actually_awarded(tmp_path):
    """Raising the continuity weight must change the objective."""
    sheets = {"Rooms": _rooms(1), "TimeSlots": _timeslots(2), "Presentations": _presentations(["A", "A"])}
    common = dict(weight_conflict=0, weight_empty_seats=0)
    off = _run(tmp_path, sheets, name="cont_off", weight_continuity=0, **common)
    on = _run(tmp_path, sheets, name="cont_on", weight_continuity=1, **common)
    assert off["objective_score"] == pytest.approx(0.0)
    assert on["candidates"][0]["continuity_bonuses"] == 1
    # span_continuity = presentations - topics = 1, so the single bonus is the
    # whole scale.
    assert on["objective_score"] == pytest.approx(engine.SCORE_SCALE, abs=0.01)


def test_score_breakdown_sums_to_the_objective(tmp_path):
    """The four components are recomputed from the schedule itself.

    If they disagree with the solver's own objective, either the model or
    the reporting is wrong -- so this is an independent check on both.
    """
    _write_small_instance(tmp_path / "bd.xlsx")
    result = engine.run_optimization(str(tmp_path / "bd.xlsx"), str(tmp_path / "bd_out.xlsx"))
    best = result["candidates"][0]
    total = (best["score_continuity"] + best["score_conflict"]
             + best["score_same_topic"] + best["score_empty_seats"])
    assert total == pytest.approx(result["objective_score"], abs=0.01)


# ---------------------------------------------------------------------
# Speaker constraint
# ---------------------------------------------------------------------

def test_speaker_is_never_double_booked(tmp_path):
    sheets = {
        "Rooms": _rooms(2),
        "TimeSlots": _timeslots(2),
        "Presentations": _presentations(["A", "B"], speakers=["Dr. Tsuruta", "Dr. Tsuruta"]),
    }
    result = _run(tmp_path, sheets, name="speaker_ok")
    rows = pd.read_excel(result["output_file"], sheet_name="Result_List")
    assert rows["TimeSlot_ID"].nunique() == 2


def test_speaker_with_more_talks_than_timeslots_is_named(tmp_path):
    sheets = {
        "Rooms": _rooms(3),
        "TimeSlots": _timeslots(1),
        "Presentations": _presentations(["A", "B"], speakers=["Dr. Umemura", "Dr. Umemura"]),
    }
    with pytest.raises(engine.InfeasibleModelError, match="Dr. Umemura"):
        _run(tmp_path, sheets, name="speaker_ng")


def test_speaker_column_is_optional(tmp_path):
    sheets = {"Rooms": _rooms(2), "TimeSlots": _timeslots(1), "Presentations": _presentations(["A", "B"])}
    result = _run(tmp_path, sheets, name="no_speaker")
    assert result["status"] == "Optimal"
    report = pd.read_excel(result["output_file"], sheet_name="Simulation_Report")
    value = report.set_index("Metric")["Value"].to_dict()
    assert value["Speaker_Constraint"] == "Disabled"


# ---------------------------------------------------------------------
# Multiple candidate schedules
# ---------------------------------------------------------------------

def test_single_solution_uses_the_plain_sheet_names(tmp_path):
    sheets = {"Rooms": _rooms(2), "TimeSlots": _timeslots(2), "Presentations": _presentations(["A", "B"])}
    result = _run(tmp_path, sheets, name="one")
    out = pd.read_excel(result["output_file"], sheet_name=None)
    assert set(out.keys()) == {"Result_Timetable", "Result_List", "Simulation_Report"}
    assert result["solutions_returned"] == 1


def test_multiple_solutions_are_distinct(tmp_path):
    sheets = {
        "Rooms": _rooms(2),
        "TimeSlots": _timeslots(3),
        "Presentations": _presentations(["A", "B", "C", "D"]),
    }
    result = _run(tmp_path, sheets, name="three", solutions=3)
    assert result["solutions_returned"] >= 2

    out = pd.read_excel(result["output_file"], sheet_name=None)
    assert "Result_Timetable_1" in out
    rows = pd.read_excel(result["output_file"], sheet_name="Result_List")
    assert "Rank" in rows.columns

    placements = {
        rank: set(zip(group["Presentation_ID"], group["TimeSlot_ID"], group["Room"]))
        for rank, group in rows.groupby("Rank")
    }
    ranks = sorted(placements)
    for a, b in zip(ranks, ranks[1:]):
        assert len(placements[a] - placements[b]) >= engine.MIN_SOLUTION_DIFFERENCE


def test_solutions_is_clamped(tmp_path):
    sheets = {"Rooms": _rooms(1), "TimeSlots": _timeslots(1), "Presentations": _presentations(["A"])}
    result = _run(tmp_path, sheets, name="clamp", solutions=99)
    assert result["solutions_requested"] == engine.MAX_SOLUTIONS
    # Only one schedule exists, so the search stops early rather than failing.
    assert result["solutions_returned"] == 1


# ---------------------------------------------------------------------
# Solution quality reporting
# ---------------------------------------------------------------------

class _FakeProblem:
    def __init__(self, sol_status):
        self.sol_status = sol_status


def test_time_limited_solution_is_not_labelled_optimal():
    """CBC says "Optimal" even when the time limit stopped the search.

    The earlier version passed that label straight through, so a user
    could not tell a proven optimum from a truncated one. PuLP's
    ``sol_status`` does distinguish them.
    """
    import pulp

    truncated = engine._describe_solution_quality(
        _FakeProblem(pulp.LpSolutionIntegerFeasible), mip_gap=0.0
    )
    assert "time limit" in truncated.lower()
    assert not truncated.startswith("Optimal (proven)")


def test_exact_and_gap_limited_solutions_are_labelled_differently():
    import pulp

    exact = engine._describe_solution_quality(_FakeProblem(pulp.LpSolutionOptimal), mip_gap=0.0)
    within = engine._describe_solution_quality(_FakeProblem(pulp.LpSolutionOptimal), mip_gap=0.02)
    assert exact == "Optimal (proven)"
    assert "2%" in within


def test_report_records_what_the_schedule_achieved_and_the_settings_used(tmp_path):
    """The old report only ever printed "No conflicts found"."""
    sheets = {
        "Rooms": _rooms(2),
        "TimeSlots": _timeslots(1),
        "Presentations": _presentations(["A", "B"]),
        "Conflicts": pd.DataFrame({"Topic1": ["A"], "Topic2": ["B"], "Conflict_Level": [1.0]}),
    }
    result = _run(tmp_path, sheets, name="report",
                  weight_continuity=0, weight_conflict=1, weight_empty_seats=0)
    report = pd.read_excel(result["output_file"], sheet_name="Simulation_Report")
    value = report.set_index("Metric")["Value"].to_dict()
    assert value["Cross_Topic_Conflicts_B1"] == 1
    assert value["Same_Topic_Overlaps_B2"] == 0
    assert value["Solution_Quality"] == "Optimal (proven)"
    assert value["Time_Limit_Setting"] == "Unlimited"
    # The settings actually used are recorded, because the workbook no
    # longer carries them.
    assert value["Weight_Conflict"] == 1
    assert value["Weight_Continuity"] == 0


# ---------------------------------------------------------------------
# Normalisation
# ---------------------------------------------------------------------

def test_normalizers_are_derived_from_the_data(tmp_path):
    """The three terms are measured in different units, so each is divided
    by the range it can span for the data at hand.

    For the small instance:
      continuity : presentations - topics = 4 - 2 = 2 (one bonus per extra
                   presentation of a topic, at best)
      conflict   : timeslots * sum(levels) + 3 * (presentations - topics)
                   = 3*1 + 3*2 = 9
      empty seats: per presentation, the spread between the largest and the
                   smallest room it could use. S1 and S2 fit only R1 (0
                   each); S3 and S4 fit R1 and R2 (50-20 = 30 each) = 60
    """
    _write_small_instance(tmp_path / "norm.xlsx")
    result = engine.run_optimization(
        str(tmp_path / "norm.xlsx"), str(tmp_path / "norm_out.xlsx"))
    report = pd.read_excel(result["output_file"], sheet_name="Simulation_Report")
    value = report.set_index("Metric")["Value"].to_dict()
    assert value["Normalizer_Continuity"] == pytest.approx(2)
    assert value["Normalizer_Conflict"] == pytest.approx(9)
    assert value["Normalizer_EmptySeats"] == pytest.approx(60)
    assert value["Score_Scale"] == pytest.approx(engine.SCORE_SCALE)


def test_weights_scale_the_terms_they_belong_to(tmp_path):
    """Doubling one weight doubles that term's contribution and nothing else."""
    sheets = {
        "Rooms": _rooms(2),
        "TimeSlots": _timeslots(1),
        "Presentations": _presentations(["A", "A"]),
    }
    single = _run(tmp_path, sheets, name="w1",
                  weight_continuity=0, weight_conflict=1, weight_empty_seats=0)
    double = _run(tmp_path, sheets, name="w2",
                  weight_continuity=0, weight_conflict=2, weight_empty_seats=0)
    assert double["objective_score"] == pytest.approx(2 * single["objective_score"], abs=0.01)


# ---------------------------------------------------------------------
# Languages (scheduler_messages)
# ---------------------------------------------------------------------

import string  # noqa: E402


def _field_names(text):
    return {name for _, name, _, _ in string.Formatter().parse(text) if name}


def test_every_message_has_both_languages():
    for key, entry in msg.MESSAGES.items():
        for language in msg.LANGUAGES:
            assert entry.get(language, "").strip(), f"{key} has no {language} text"


def test_both_languages_fill_in_the_same_fields():
    """A field present in one language only would crash, or vanish, in the other."""
    for key, entry in msg.MESSAGES.items():
        assert _field_names(entry["ja"]) == _field_names(entry["en"]), key


def test_english_messages_contain_no_japanese():
    import re
    for key, entry in msg.MESSAGES.items():
        assert not re.search(r"[぀-ヿ一-鿿]", entry["en"]), key


@pytest.mark.parametrize("code, expected", [
    ("ja-JP", "ja"), ("ja", "ja"), ("ja_JP.UTF-8", "ja"),
    ("en-JP", "en"), ("en", "en"), ("fr-FR", "en"), ("", "en"), (None, "en"),
])
def test_language_from_code(code, expected):
    assert msg.language_from_code(code) == expected


def test_parse_apple_languages_reads_the_first_preferred_language():
    output = '(\n    "ja-JP",\n    "en-JP",\n    en\n)\n'
    assert msg.parse_apple_languages(output) == ["ja-JP", "en-JP", "en"]
    assert msg.parse_apple_languages('(\n    "en-US",\n    "ja-JP"\n)') [0] == "en-US"


def test_language_detection_falls_back_to_english(monkeypatch):
    def broken():
        raise OSError("setting unavailable")
    monkeypatch.setattr(msg, "_system_language_code", broken)
    assert msg.detect_language() == "en"


def test_language_detection_uses_the_system_setting(monkeypatch):
    monkeypatch.setattr(msg, "_system_language_code", lambda: "ja-JP")
    assert msg.detect_language() == "ja"
    monkeypatch.setattr(msg, "_system_language_code", lambda: "de-DE")
    assert msg.detect_language() == "en"


def test_unknown_language_is_rejected():
    with pytest.raises(ValueError):
        msg.set_language("fr")


@pytest.mark.parametrize("language, header, hint", [
    ("ja", "入力ファイルの形式が合っていません", "README_JP"),
    ("en", "does not match the required format", "README_EN"),
])
def test_input_errors_follow_the_chosen_language(tmp_path, language, header, hint):
    msg.set_language(language)
    sheets = _valid_sheets()
    del sheets["Rooms"]
    path = tmp_path / "missing.xlsx"
    _write_workbook(path, sheets)
    with pytest.raises(engine.InputValidationError) as excinfo:
        engine.validate_input_workbook(str(path))
    assert header in str(excinfo.value) and hint in str(excinfo.value)


# ---------------------------------------------------------------------
# Choosing a CBC that actually runs
# ---------------------------------------------------------------------

def _fake_cbc(tmp_path, name, exit_code):
    """A stand-in 'cbc' that exits with ``exit_code`` (a broken copy when non-zero)."""
    if sys.platform.startswith("win"):
        pytest.skip("shell-script stand-in is POSIX only")
    path = tmp_path / name
    path.write_text(f"#!/bin/sh\necho 'Version: 9.9.9'\nexit {exit_code}\n")
    path.chmod(0o755)
    return str(path)


def _real_pulp_cbc():
    import pulp
    return pulp.PULP_CBC_CMD(msg=0).path


def test_a_bundled_cbc_that_will_not_start_is_skipped(tmp_path, monkeypatch):
    """Existence is not enough: a Homebrew copy exists everywhere but runs only where
    its libraries are installed. The next candidate must be used instead."""
    broken = _fake_cbc(tmp_path, "cbc_broken", 1)
    monkeypatch.setattr(engine, "_cbc_candidates",
                        lambda: [(broken, "Bundled"), (_real_pulp_cbc(), "PuLP")])
    engine._probe_cbc.cache_clear()
    info = engine.resolve_cbc()
    assert info.source == "PuLP"
    assert info.version != "unknown"


def test_a_working_bundled_cbc_is_preferred(tmp_path, monkeypatch):
    working = _fake_cbc(tmp_path, "cbc_ok", 0)
    monkeypatch.setattr(engine, "_cbc_candidates",
                        lambda: [(working, "Bundled"), (_real_pulp_cbc(), "PuLP")])
    engine._probe_cbc.cache_clear()
    info = engine.resolve_cbc()
    assert (info.source, info.version) == ("Bundled", "9.9.9")


def test_no_working_cbc_gives_a_clear_error(tmp_path, monkeypatch):
    broken = _fake_cbc(tmp_path, "cbc_broken2", 1)
    monkeypatch.setattr(engine, "_cbc_candidates", lambda: [(broken, "Bundled")])
    engine._probe_cbc.cache_clear()
    with pytest.raises(engine.SolverUnavailableError, match="CBC"):
        engine.resolve_cbc()


def test_report_records_which_cbc_was_used(tmp_path):
    engine._probe_cbc.cache_clear()
    sheets = {"Rooms": _rooms(2), "TimeSlots": _timeslots(1), "Presentations": _presentations(["A", "B"])}
    result = _run(tmp_path, sheets, name="cbc_report")
    report = pd.read_excel(result["output_file"], sheet_name="Simulation_Report")
    value = report.set_index("Metric")["Value"].to_dict()
    assert value["Solver_CBC_Source"] in ("Bundled", "System", "PuLP")
    assert str(value["Solver_CBC_Version"]).strip()


# ---------------------------------------------------------------------
# Terms: presentations and topics
# ---------------------------------------------------------------------

def _conflict_case():
    return {
        "Rooms": _rooms(2),
        "TimeSlots": _timeslots(2),
        "Presentations": _presentations(["A", "A", "B"], speakers=["x", "y", "x"]),
        "Conflicts": pd.DataFrame({"Topic1": ["A"], "Topic2": ["B"], "Conflict_Level": [2.0]}),
    }


#: The words the input format used before 2026-09-23, in English and
#: Japanese. Spelled in pieces so that searching the code for them finds
#: nothing -- which is exactly what the last test below checks.
RETIRED_WORDS = ("sess" + "ion", "tra" + "ck", "セッ" + "ション", "トラ" + "ック")


def _retired_words_in(text):
    lowered = str(text).lower()
    return [w for w in RETIRED_WORDS if w in lowered]


def test_output_uses_the_current_names(tmp_path):
    result = _run(tmp_path, _conflict_case(), name="out_names")
    sheets = pd.read_excel(result["output_file"], sheet_name=None)
    assert {"Presentation_ID", "Topic"} <= set(sheets["Result_List"].columns)
    metrics = set(sheets["Simulation_Report"]["Metric"])
    assert {"Cross_Topic_Conflicts_B1", "Same_Topic_Overlaps_B2", "Score_SameTopic_B2"} <= metrics
    for name, frame in sheets.items():
        assert not _retired_words_in(name)
        for column in frame.columns:
            assert not _retired_words_in(column), (name, column)
    assert not any(_retired_words_in(m) for m in metrics)


@pytest.mark.parametrize("language", ["ja", "en"])
def test_messages_name_the_current_sheet_and_columns(tmp_path, language):
    msg.set_language(language)
    sheets = _conflict_case()
    sheets["Presentations"] = sheets["Presentations"].drop(columns=["Topic"])
    path = tmp_path / "no_topic.xlsx"
    _write_workbook(path, sheets)
    with pytest.raises(engine.InputValidationError) as excinfo:
        engine.validate_input_workbook(str(path))
    assert "Presentations" in str(excinfo.value) and "Topic" in str(excinfo.value)
    assert not _retired_words_in(excinfo.value)


def test_no_message_uses_a_retired_word():
    for key, entry in msg.MESSAGES.items():
        assert not _retired_words_in(key), key
        for language in msg.LANGUAGES:
            assert not _retired_words_in(entry[language]), (key, language)


def test_no_file_in_the_distribution_uses_a_retired_word():
    """Code, launchers, sample workbooks and this test file itself."""
    import openpyxl
    root = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
    found = []
    for folder, dirs, files in os.walk(root):
        dirs[:] = [d for d in dirs if d not in (".venv", "__pycache__", "build", "dist", "logs")]
        for name in files:
            path = os.path.join(folder, name)
            rel = os.path.relpath(path, root)
            if _retired_words_in(rel):
                found.append(rel)
            if name.endswith((".py", ".command", ".txt", ".spec", ".md", ".bat")):
                with open(path, encoding="utf-8") as handle:
                    for number, line in enumerate(handle, start=1):
                        if _retired_words_in(line):
                            found.append(f"{rel}:{number}")
            elif name.endswith(".xlsx") and not name.startswith(("output", "~$")):
                workbook = openpyxl.load_workbook(path, read_only=True)
                for sheet in workbook.worksheets:
                    if _retired_words_in(sheet.title):
                        found.append(f"{rel}[{sheet.title}]")
                    for row in sheet.iter_rows(values_only=True):
                        for value in row:
                            if isinstance(value, str) and _retired_words_in(value):
                                found.append(f"{rel}[{sheet.title}]: {value[:40]}")
    assert not found, found
