#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
scheduler_messages.py
======================
Every sentence the user can read, in Japanese and English, plus the
choice of which language to show.

Why a separate file
-------------------
Wording changes for different reasons than the model or the window
layout do, so it lives on its own. Nothing here imports the other
modules; the engine, the GUI and the entry point all read from it.
Keeping both languages side by side in one table also makes it easy to
check a translation, and lets a test confirm that no entry is missing
one of the two.

What is *not* translated
------------------------
Progress lines in the run log, CBC's own output and the column/sheet
names of the output workbook stay in English in both languages. They are
technical records rather than instructions, and one fixed form keeps
logs and outputs comparable across machines.

Usage
-----
    import scheduler_messages as msg
    msg.set_language(msg.detect_language())
    msg.t("input.blank", sheet="Rooms", column="Capacity", row=3)
"""

from __future__ import annotations

import locale
import os
import platform
import re
import subprocess
from typing import Dict, List

LANGUAGES = ("ja", "en")

#: Used whenever the system language is not Japanese, or cannot be read.
DEFAULT_LANGUAGE = "en"

#: Reading the macOS setting normally takes a few milliseconds; give up
#: well before a user would notice a slow start.
DETECT_TIMEOUT_SEC = 1.0

_current = DEFAULT_LANGUAGE


# =====================================================================
# Language selection
# =====================================================================

def set_language(language: str) -> None:
    global _current
    if language not in LANGUAGES:
        raise ValueError(f"Unsupported language: {language!r} (choose from {LANGUAGES})")
    _current = language


def get_language() -> str:
    return _current


def language_from_code(code) -> str:
    """Map a locale or language code ("ja-JP", "ja_JP.UTF-8", "en") to ja/en."""
    return "ja" if str(code or "").strip().lower().startswith("ja") else "en"


def parse_apple_languages(output: str) -> List[str]:
    """Extract the codes from ``defaults read -g AppleLanguages`` output.

    The command prints a property list such as::

        (
            "ja-JP",
            en
        )
    """
    return re.findall(r'^\s*"?([A-Za-z]{2,3}(?:[-_][A-Za-z0-9]+)*)"?\s*,?\s*$', output, re.M)


def _system_language_code() -> str:
    system = platform.system()
    if system == "Darwin":
        # The "Preferred languages" list in System Settings. Readable by
        # any process: no permission prompt is involved. The locale
        # environment variables are not used because an app started from
        # the Finder often has none set.
        result = subprocess.run(
            ["defaults", "read", "-g", "AppleLanguages"],
            capture_output=True, text=True, timeout=DETECT_TIMEOUT_SEC, check=True,
        )
        codes = parse_apple_languages(result.stdout)
        return codes[0] if codes else ""
    if system == "Windows":
        import ctypes
        lang_id = ctypes.windll.kernel32.GetUserDefaultUILanguage()  # type: ignore[attr-defined]
        return "ja" if (lang_id & 0x3FF) == 0x11 else "en"   # 0x11 = LANG_JAPANESE
    code = locale.getlocale()[0] or os.environ.get("LC_ALL") or os.environ.get("LANG") or ""
    return code


def detect_language() -> str:
    """Japanese only when the system's first preferred language is Japanese.

    Any other language, and any failure to read the setting, gives English.
    """
    try:
        return language_from_code(_system_language_code())
    except Exception:
        return DEFAULT_LANGUAGE


def t(key: str, **fields) -> str:
    """The message ``key`` in the current language, with ``fields`` filled in."""
    entry = MESSAGES[key]
    text = entry.get(_current) or entry[DEFAULT_LANGUAGE]
    return text.format(**fields) if fields else text


# =====================================================================
# Messages
# =====================================================================
# Keys are grouped by where the text appears:
#   input.*   problems found in the input workbook
#   run.*     problems found while building or solving the model
#   solver.*  locating CBC
#   gui.*     the window

MESSAGES: Dict[str, Dict[str, str]] = {
    # --- input workbook -----------------------------------------------
    "input.blank": {
        "ja": "{sheet} の {column} が空欄です（{row}行目）",
        "en": "{sheet}: {column} is blank (row {row})",
    },
    "input.duplicate": {
        "ja": "{sheet} の {column} が重複しています: {values}（同じ識別子は使えません）",
        "en": "{sheet}: {column} has duplicates: {values} (each ID must be unique)",
    },
    "input.not_number": {
        "ja": "{sheet} の {column} が数値ではありません（{row}行目: \"{value}\"）",
        "en": "{sheet}: {column} is not a number (row {row}: \"{value}\")",
    },
    "input.negative": {
        "ja": "{sheet} の {column} が負の値です（{row}行目: {value}）",
        "en": "{sheet}: {column} is negative (row {row}: {value})",
    },
    "input.empty_sheet": {
        "ja": "{sheet} にデータが1行もありません",
        "en": "{sheet} has no data rows",
    },
    "input.too_many_presentations": {
        "ja": "発表が {presentations} 件ありますが、会場×時間枠の組み合わせは "
              "{rooms}×{slots}＝{capacity} 通りしかありません（会場か時間枠を増やしてください）",
        "en": "There are {presentations} presentations but only {rooms} rooms × {slots} timeslots "
              "= {capacity} room-slots (add rooms or timeslots)",
    },
    "input.unknown_ng_slot": {
        "ja": "Presentations の NG_TimeSlot_ID に、TimeSlots に無い時間枠が指定されています: {values}",
        "en": "Presentations: NG_TimeSlot_ID refers to timeslots that are not in TimeSlots: {values}",
    },
    "input.file_not_found": {
        "ja": "入力ファイルが見つかりません: '{path}'",
        "en": "Input file not found: '{path}'",
    },
    "input.unreadable": {
        "ja": "Excelファイルを読み込めませんでした: '{name}'\n詳細: {detail}",
        "en": "Could not read the Excel file: '{name}'\nDetails: {detail}",
    },
    "input.missing_sheet": {
        "ja": "シート '{sheet}' がありません",
        "en": "Sheet '{sheet}' is missing",
    },
    "input.missing_columns": {
        "ja": "シート '{sheet}' に必要な列がありません: {columns}",
        "en": "Sheet '{sheet}' is missing required columns: {columns}",
    },
    "input.missing_columns_optional": {
        "ja": "シート '{sheet}' に必要な列がありません: {columns}"
              "（このシート自体は任意です。使わない場合はシートごと削除してください）",
        "en": "Sheet '{sheet}' is missing required columns: {columns} "
              "(the sheet itself is optional; delete the whole sheet if you do not use it)",
    },
    "input.header": {
        "ja": "入力ファイルの形式が合っていません:",
        "en": "The input file does not match the required format:",
    },
    "input.more": {
        "ja": "  ...ほか {count} 件",
        "en": "  ...and {count} more",
    },
    "input.readme_hint": {
        "ja": "入力ファイルの README_JP シートに各列の説明があります。",
        "en": "The README_EN sheet of the input file explains each column.",
    },

    # --- building / solving the model ---------------------------------
    "run.settings_sheet_ignored": {
        "ja": " -> [Info] シート {sheets} は使用しません"
              "（重み・時間制限・許容ギャップ・候補案の数はアプリ側で指定します）。",
        "en": " -> [Info] Sheet {sheets} is not used "
              "(weights, time limit, MIP gap and number of candidates are set in the app).",
    },
    "run.no_room_line": {
        "ja": "  - 発表 {presentation}（想定参加者数 {pax}{screen}）",
        "en": "  - Presentation {presentation} (expected attendance {pax}{screen})",
    },
    "run.no_room_screen": {
        "ja": "、スクリーン必要 {value}",
        "en": ", needs screen: {value}",
    },
    "run.no_room": {
        "ja": "次の発表は、収容人数・設備・登壇不可時間枠の条件を満たす"
              "会場と時間枠の組み合わせが1つもありません:\n{details}\n\n"
              "会場の定員や設備、NG_TimeSlot_ID の設定を見直してください。",
        "en": "No combination of room and timeslot meets the capacity, equipment, and "
              "unavailable-timeslot conditions for these presentations:\n{details}\n\n"
              "Check the room capacities, the equipment and NG_TimeSlot_ID.",
    },
    "run.speaker_line": {
        "ja": "  - 登壇者 '{name}' は {count} 件の発表がありますが、"
              "割り当て可能な時間枠が {usable} 個しかありません（対象: {presentations}）",
        "en": "  - Speaker '{name}' has {count} presentations but only {usable} available "
              "timeslots (presentations: {presentations})",
    },
    "run.speaker_overlap": {
        "ja": "同一人物が同じ時間枠に重複して登壇しない条件を満たせません:\n{details}\n\n"
              "時間枠を増やすか、NG_TimeSlot_ID の設定を見直してください。\n"
              "※ 別人が同姓同名の場合は、Speaker 名の末尾に A, B, C… を付けて区別してください。",
        "en": "The constraint that no speaker presents twice in the same timeslot cannot be satisfied:\n"
              "{details}\n\n"
              "Add timeslots or check NG_TimeSlot_ID.\n"
              "Note: if two different people share a name, add A, B, C… to the end of "
              "the Speaker name to tell them apart.",
    },
    "run.infeasible": {
        "ja": "実行可能なスケジュールが見つかりませんでした (ステータス: {status})。\n\n"
              "【考えられる原因】\n"
              "・会場の定員に対して想定参加者数が大きすぎる発表がある\n"
              "・設備要件（Needs_Screen）を満たす会場が足りない\n"
              "・登壇不可時間枠（NG_TimeSlot_ID）の指定が厳しすぎる\n"
              "・同一人物が同じ時間枠に重ならないように配置できない\n"
              "・会場数 × 時間枠数が発表数に対して足りない",
        "en": "No feasible schedule was found (status: {status}).\n\n"
              "Possible causes:\n"
              "- a presentation expects more attendees than any room can hold\n"
              "- too few rooms meet the equipment requirement (Needs_Screen)\n"
              "- the unavailable timeslots (NG_TimeSlot_ID) are too restrictive\n"
              "- a speaker's presentations cannot all be placed in different timeslots\n"
              "- the total number of room-slots (rooms × timeslots) is less than the number of presentations",
    },
    "run.breakdown_mismatch": {
        "ja": "[Warning] 目的関数の内訳の合計 ({recomputed}) が "
              "ソルバーの報告値 ({reported}) と一致しません。",
        "en": "[Warning] The score breakdown adds up to {recomputed}, "
              "but the solver reported {reported}.",
    },

    # --- locating CBC ---------------------------------------------------
    "solver.unavailable": {
        "ja": "CBC ソルバーを起動できませんでした。\n"
              "同梱の CBC と PuLP 付属の CBC のどちらも起動できません。\n"
              "Apple Silicon の Mac では、PuLP 付属の CBC に Rosetta 2 が必要です。",
        "en": "Could not start the CBC solver.\n"
              "Neither the bundled CBC nor the one that ships with PuLP would run.\n"
              "On Apple Silicon Macs, the CBC that ships with PuLP needs Rosetta 2.",
    },

    # --- window: layout -----------------------------------------------
    "gui.files": {"ja": "ファイル", "en": "Files"},
    "gui.input": {"ja": "入力", "en": "Input"},
    "gui.output": {"ja": "出力", "en": "Output"},
    "gui.browse": {"ja": "参照...", "en": "Browse..."},
    "gui.change": {"ja": "変更...", "en": "Change..."},
    "gui.weights": {"ja": "重み（何を優先するか）", "en": "Weights (what to prioritize)"},
    "gui.w_continuity": {"ja": "分野の連続性", "en": "Topic continuity"},
    "gui.w_conflict": {"ja": "分野の衝突", "en": "Topic conflicts"},
    "gui.w_empty": {"ja": "空席", "en": "Empty seats"},
    "gui.weights_hint": {
        "ja": "3つを同じ値にすると対等に扱います（0〜10、0.1刻み）。0 にするとその観点を考慮しません。",
        "en": "Setting equal values gives all three criteria equal weight (0–10, in steps of 0.1). "
              "A value of 0 ignores that criterion.",
    },
    "gui.run_settings": {"ja": "計算の設定", "en": "Run settings"},
    "gui.mip_gap": {"ja": "許容ギャップ", "en": "MIP gap"},
    "gui.time_limit": {"ja": "時間制限", "en": "Time limit"},
    "gui.seconds_unit": {"ja": "秒", "en": "s"},
    "gui.solutions": {"ja": "候補案の数", "en": "Candidates"},
    "gui.settings_hint": {
        "ja": "許容ギャップ 空欄＝厳密解　／　時間制限 空欄＝無制限　／　"
              "候補案 {max} を選ぶと別案も出します（案ごとに解き直すため時間も増えます）",
        "en": "MIP gap blank = exact solution  /  Time limit blank = no limit  /  "
              "Choose {max} candidates to get alternative schedules "
              "(each one is solved again, so it takes longer)",
    },
    "gui.save_log": {"ja": "実行ログをファイルに保存", "en": "Save the run log to a file"},
    "gui.run": {"ja": "▶ 最適化を実行", "en": "▶ Run optimization"},
    "gui.log": {"ja": "実行ログ", "en": "Run log"},
    "gui.preview": {"ja": "時間割を表示", "en": "Show timetable"},
    "gui.open_excel": {"ja": "Excelを開く", "en": "Open Excel"},
    "gui.open_folder": {"ja": "フォルダを開く", "en": "Open folder"},
    "gui.quit": {"ja": "終了", "en": "Quit"},

    # --- window: input summary ----------------------------------------
    "gui.summary_log_header": {"ja": "\n[入力ファイルの確認]\n", "en": "\n[Input file check]\n"},
    "gui.summary_invalid": {
        "ja": "⚠ {first}（詳細はログを参照）",
        "en": "⚠ {first} (see the run log for details)",
    },
    "gui.summary_unreadable": {"ja": "⚠ 読み込めませんでした: {error}", "en": "⚠ Could not read the file: {error}"},
    "gui.n_presentations": {"ja": "{n} 発表", "en": "{n} presentations"},
    "gui.n_rooms": {"ja": "{n} 会場", "en": "{n} rooms"},
    "gui.n_timeslots": {"ja": "{n} コマ", "en": "{n} timeslots"},
    "gui.n_topics": {"ja": "{n} 分野", "en": "{n} topics"},
    "gui.n_speakers": {"ja": "{n} 名の登壇者", "en": "{n} speakers"},
    "gui.n_conflicts": {"ja": "競合 {n} 組", "en": "{n} conflict pairs"},
    "gui.slots_available": {"ja": "　（配置できる枠は {n} 通り）", "en": "  ({n} room-slots available)"},

    # --- window: dialogs ----------------------------------------------
    "gui.choose_input": {"ja": "入力Excelファイルを選択", "en": "Choose the input Excel file"},
    "gui.choose_output": {
        "ja": "出力Excelファイルの保存先を指定",
        "en": "Choose where to save the output Excel file",
    },
    "gui.input_error_title": {"ja": "入力エラー", "en": "Input error"},
    "gui.not_number": {"ja": "{label}は数値で入力してください。{note}", "en": "{label} must be a number. {note}"},
    "gui.negative": {"ja": "{label}に負の値は指定できません。{note}", "en": "{label} cannot be negative. {note}"},
    "gui.gap_note": {"ja": "空欄なら厳密解を求めます。", "en": "Leave it blank for an exact solution."},
    "gui.limit_note": {"ja": "空欄なら時間制限なしで計算します。", "en": "Leave it blank for no time limit."},
    "gui.input_empty": {"ja": "入力ファイルを選択してください。", "en": "Choose an input file."},
    "gui.file_missing_title": {"ja": "ファイルが見つかりません", "en": "File not found"},
    "gui.input_missing": {"ja": "入力ファイルが見つかりません:\n{path}", "en": "Input file not found:\n{path}"},
    "gui.output_error_title": {"ja": "出力エラー", "en": "Output error"},
    "gui.output_missing": {
        "ja": "出力ファイルの保存先を指定してください。",
        "en": "Choose where to save the output file.",
    },
    "gui.overwrite_title": {"ja": "出力ファイルの上書き", "en": "Overwrite output file"},
    "gui.overwrite_body": {
        "ja": "既に同じ名前のファイルがあります:\n{name}\n\n"
              "上書きしますか？\n「いいえ」を選ぶと、日時を付けた別名で保存します。",
        "en": "A file with this name already exists:\n{name}\n\n"
              "Overwrite it?\nChoose \"No\" to save it under a new name with the date and time added.",
    },
    "gui.warning_title": {"ja": "警告", "en": "Warning"},
    "gui.error_title": {"ja": "エラー", "en": "Error"},

    # --- window: running and results ----------------------------------
    "gui.log_saving": {"ja": "[Info] 実行ログを保存します: {path}\n", "en": "[Info] Saving the run log to: {path}\n"},
    "gui.log_open_failed": {
        "ja": "[Warning] ログファイルを開けませんでした: {error}\n",
        "en": "[Warning] Could not open the log file: {error}\n",
    },
    "gui.elapsed": {"ja": "経過 {n} 秒", "en": "Elapsed {n} s"},
    "gui.running": {"ja": "計算中...", "en": "Solving..."},
    "gui.run_started": {"ja": "[{time}] 最適化を開始します\n", "en": "[{time}] Starting optimization\n"},
    "gui.preview_load_failed": {
        "ja": "[Warning] プレビュー用の読み込みに失敗しました: {error}\n",
        "en": "[Warning] Could not load the result for the preview: {error}\n",
    },
    "gui.headline": {
        "ja": "スコア {score} ／ {time} 秒 ／ {count} 件を配置",
        "en": "Score {score}  /  {time} s  /  {count} presentations placed",
    },
    "gui.headline_candidates": {"ja": " ／ 候補 {n} 案", "en": "  /  {n} candidates"},
    "gui.quality_truncated": {
        "ja": "時間制限により打ち切られた解（最適であることは確認できていません）",
        "en": "Stopped by the time limit (not proven optimal)",
    },
    "gui.quality_gap": {"ja": "最適値から {percent} 以内の解", "en": "Within {percent} of the optimum"},
    "gui.quality_exact": {
        "ja": "厳密最適解（最適であることを証明済み）",
        "en": "Exact optimum (proven optimal)",
    },
    "gui.breakdown": {
        "ja": "内訳　連続性 {a_count}件 {a:+g}　／　異分野の衝突 {b1_count}組 {b1:+g}"
              "　／　同一分野の重複 {b2_count}件 {b2:+g}　／　空席 {c_count:g}席 {c:+g}",
        "en": "Breakdown  continuity {a_count} {a:+g}  /  cross-topic conflicts {b1_count} {b1:+g}"
              "  /  same-topic overlaps {b2_count} {b2:+g}  /  empty seats {c_count:g} {c:+g}",
    },
    "gui.aborted": {"ja": "\n[Error] 最適化を中止しました:\n", "en": "\n[Error] Optimization stopped:\n"},
    "gui.unknown_error": {"ja": "不明なエラー", "en": "Unknown error"},
    "gui.see_log": {
        "ja": "詳細は下の実行ログを確認してください。",
        "en": "See the run log below for details.",
    },
    "gui.no_result": {"ja": "表示できる結果がありません。", "en": "No results to display."},
    "gui.timetable_title": {"ja": "時間割", "en": "Timetable"},
    "gui.candidate_tab": {"ja": "案 {n}", "en": "Candidate {n}"},
    "gui.open_file_failed": {"ja": "ファイルを開けませんでした:\n{error}", "en": "Could not open the file:\n{error}"},
    "gui.output_not_found": {"ja": "出力ファイルが見つかりません。", "en": "Output file not found."},
    "gui.open_folder_failed": {"ja": "フォルダを開けませんでした:\n{error}", "en": "Could not open the folder:\n{error}"},
    "gui.folder_not_found": {"ja": "フォルダが見つかりません:\n{folder}", "en": "Folder not found:\n{folder}"},
}
