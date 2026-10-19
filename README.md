# Conference Scheduler

A tool that builds the program of parallel oral sessions at a conference — which presentation goes in which room and timeslot — from an Excel input, using mathematical optimization (mixed-integer linear programming).
The output timetable is meant as a first draft; final adjustments are expected to be made by people.

## Features
- Hard constraints (always satisfied)
  - Each presentation is scheduled exactly once
  - At most one presentation per room and timeslot
  - Room capacity ≥ expected attendance
  - (Optional) Screen requirements, timeslots a speaker cannot attend, no speaker in two presentations at the same time
- Objectives (priorities adjustable with weights)
  - A: Keep presentations of the same topic consecutive in the same room
  - B: Reduce conflicts between presentations held at the same time (related topics / the same topic)
  - C: Reduce empty seats
- Up to 3 candidate schedules; time limit and MIP gap can be set
- Japanese / English interface

## Requirements
- App: macOS 26 or later, Apple Silicon
- Python version: Python 3.9 or later

## Usage
### App
1. Download `Conference_Scheduler_v1.0.0_macOS.zip` from [Releases](https://github.com/Tachyon327/Conference-Scheduler/releases/latest) and double-click it to extract.
2. Open `Conference_Scheduler.app`. If macOS blocks it on first launch, go to System Settings → Privacy & Security → "Open Anyway".
3. Select an input Excel file, set the weights and other options, and run. The result is saved as an Excel file.

### Python
```
pip install -r requirements.txt
python3 new_conference_scheduler_V1.py                              # GUI
python3 new_conference_scheduler_V1.py --cli -i input_sample.xlsx   # command line
```
- Main command-line options: `--w-continuity` `--w-conflict` `--w-empty` (weights, default 1), `--time-limit` (seconds), `--mip-gap` (%), `--solutions` (1–3), `-o` (output file), `--lang ja|en`. See `--help` for all options.
- The CBC solver bundled with PuLP is used automatically; no separate installation is needed.
- On a Mac, you can also start the program by double-clicking `Launch_App.command` (it installs the required libraries on the first run). Before the first run, execute `chmod +x Launch_App.command` in Terminal.

## Input Excel
Use `input_sample.xlsx` as a template. Each column is described in the README_EN / README_JP sheets of the same file.

| Sheet | Required columns | Optional columns |
|---|---|---|
| Rooms | Room_ID, Capacity | Room_Name, Screen |
| TimeSlots | TimeSlot_ID, Day | Start_Time, End_Time |
| Presentations | Presentation_ID, Topic, Expected_Pax | Needs_Screen, NG_TimeSlot_ID, Speaker |
| Conflicts (optional sheet) | Topic1, Topic2, Conflict_Level (1–3) | |

Weights, time limit and other run settings are set in the app window (or on the command line), not in the Excel file.

## Output Excel
- Result_Timetable: timetable (rooms × timeslots)
- Result_List: assignment of each presentation
- Simulation_Report: score and each objective, computation time, settings used, and CBC version

## Files
```
new_conference_scheduler_V1.py   Entry point (GUI / command line)
scheduler_engine.py              Input validation, optimization model, output
scheduler_gui.py                 GUI
scheduler_messages.py            Japanese / English messages
input_sample.xlsx                Input template
Launch_App.command               Starts the Python version on a Mac
tests/                           Tests
data/                            Validation data (inputs and outputs)
```

## Validation data (data/)
Inputs and outputs computed with v1.0.0. The computing environment (CBC version, etc.) is recorded in the Simulation_Report sheet of each output.

- `1_abstract_reverification/`: a small dataset solved repeatedly with four weight settings
  - `input_Abstract.xlsx`: 40 presentations, 10 rooms, 5 timeslots, 7 topics
  - 18 output files, named `output<weights>_<m><n>.xlsx`
    - Weights (A, B, C): ① (1,1,1), ② (1,0,0), ③ (0,1,0), ④ (0,0,1)
    - m: number of candidate schedules; n: run number
- `2_practical_test/`: mock data based on the public program of a real conference
  - `input_Mock.xlsx`: 89 presentations, 4 rooms, 24 timeslots, 8 topics. Room capacities, expected attendance and conflicting topic pairs are assumed
  - Output: weights 1:1:1, time limit 90 s

## Tests
```
pip install -r requirements-dev.txt
pytest -q
```

## License
MIT License.
The app (zip) includes open-source software such as CBC (Eclipse Public License 2.0), Python, PuLP and pandas, each under its own license.

## Related presentations
- "A Practical System For Optimizing the Large-Scale Conference Programs", 28th SPVM National Physics Conference, October 2026

## Roadmap (updates planned until around March 2027)
- Switch the solver from CBC to HiGHS
- Improve the definition of objective B (conflicts between simultaneous presentations)
- Build a Windows app and support more macOS versions
