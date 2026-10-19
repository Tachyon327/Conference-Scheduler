#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Conference Presentation Scheduling & Venue Optimization System
File: new_conference_scheduler_V1.py

Thin entry point. The optimization model lives in ``scheduler_engine.py``,
the desktop GUI in ``scheduler_gui.py`` and every user-facing sentence, in
Japanese and English, in ``scheduler_messages.py``. Splitting them means the
MILP engine can be unit-tested and reused headlessly, ``--cli`` mode does
not require Tkinter at all, and PyInstaller still bundles both local
modules because it follows imports statically.

The input workbook holds conference data only. Weights and run settings
are given here (or in the GUI), so a result is never ambiguous about
which values produced it.

Examples:
    python3 new_conference_scheduler_V1.py
    python3 new_conference_scheduler_V1.py --cli -i input.xlsx -o output.xlsx
    python3 new_conference_scheduler_V1.py --cli -i input.xlsx   # output.xlsx beside the input
    python3 new_conference_scheduler_V1.py --cli -i input.xlsx --w-conflict 10 --w-continuity 0 --w-empty 0
    python3 new_conference_scheduler_V1.py --cli -i input.xlsx --mip-gap 2 --time-limit 600
    python3 new_conference_scheduler_V1.py --cli -i input.xlsx --solutions 3
    python3 new_conference_scheduler_V1.py --lang ja          # force Japanese

Language: Japanese when the system's first preferred language is
Japanese, English otherwise. ``--lang`` overrides it in both modes.
"""

import argparse
import os
import sys

import scheduler_engine as engine
import scheduler_messages as msg


def main():
    parser = argparse.ArgumentParser(
        description="Conference Presentation Scheduling & Venue Optimization System"
    )
    parser.add_argument("--cli", action="store_true", help="Run in command-line mode without GUI")
    parser.add_argument("--input", "-i", default=None,
                        help="Path to input Excel file (required with --cli)")
    parser.add_argument("--output", "-o", default=None,
                        help="Path to output Excel file "
                             "(default: output.xlsx in the input file's folder)")
    parser.add_argument("--lang", choices=msg.LANGUAGES, default=None,
                        help="Language of messages (default: follow the system setting)")

    weights = parser.add_argument_group("weights (what to prioritise)")
    weights.add_argument("--w-continuity", type=float, default=engine.DEFAULT_WEIGHT,
                         metavar="N", help="Topic-continuity weight (default: 1.0)")
    weights.add_argument("--w-conflict", type=float, default=engine.DEFAULT_WEIGHT,
                         metavar="N", help="Parallel-conflict weight (default: 1.0)")
    weights.add_argument("--w-empty", type=float, default=engine.DEFAULT_WEIGHT,
                         metavar="N", help="Empty-seat weight (default: 1.0)")

    run = parser.add_argument_group("run settings")
    run.add_argument("--mip-gap", type=float, default=None, metavar="PERCENT",
                     help="Relative MIP gap in percent (e.g. 2 for 2%%). "
                          "Omit for an exact solve.")
    run.add_argument("--time-limit", type=int, default=None, metavar="SECONDS",
                     help="Stop each solve after this many seconds. "
                          "Omit to solve without a limit.")
    run.add_argument("--solutions", type=int, default=engine.DEFAULT_SOLUTIONS,
                     choices=range(1, engine.MAX_SOLUTIONS + 1),
                     metavar=f"1-{engine.MAX_SOLUTIONS}",
                     help="Number of candidate schedules (default: 1). Each extra "
                          "candidate is solved from scratch, so the runtime scales "
                          "with this number.")

    args = parser.parse_args()

    if not args.cli:
        # Imported lazily so `--cli` works on systems without Tkinter.
        import scheduler_gui
        scheduler_gui.launch_gui(language=args.lang)
        return

    msg.set_language(args.lang or msg.detect_language())

    # As in the window: nothing is assumed about the working directory.
    if not args.input:
        parser.error("--cli needs an input file: -i / --input PATH")
    output = args.output or os.path.join(
        os.path.dirname(os.path.abspath(args.input)), "output.xlsx")

    try:
        engine.run_optimization(
            args.input,
            output,
            weight_continuity=args.w_continuity,
            weight_conflict=args.w_conflict,
            weight_empty_seats=args.w_empty,
            mip_gap=(args.mip_gap / 100.0) if args.mip_gap is not None else engine.DEFAULT_MIP_GAP,
            time_limit=args.time_limit,
            solutions=args.solutions,
        )
    except Exception as exc:
        print(f"\n[Fatal Error] {exc}", file=sys.stderr)
        sys.exit(1)


if __name__ == "__main__":
    main()
