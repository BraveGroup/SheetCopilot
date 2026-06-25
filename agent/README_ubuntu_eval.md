# Ubuntu (Excel-free) Evaluation

The original SheetCopilot evaluator (`evaluation.py` + `utils/compare_sheets.py`)
drives **Excel 2019 on Windows** through `win32com`/`pywin32`.  This directory
adds a **parallel, Excel-free re-implementation that runs on Ubuntu**, producing
the *same* outcome-based metrics (Exec@1, Pass@1, A_mean/A50/A90) against the
*same* result folders and `*_check.yaml` ground-truth check-boards.

The original Windows scripts are **left completely untouched**. Everything here
is additive.

## What was added

| File | Purpose |
|------|---------|
| `evaluation_ubuntu.py` | Entry point: multiprocessing driver, logging, checkpoint/resume, metrics (mirrors `evaluation.py`). |
| `ubuntu_eval/dispatcher.py` | `compare_workbooks(gt, res, check_boards)` — drop-in replacement for the Windows function; routes each category to the right engine. |
| `ubuntu_eval/oxl_engine.py` | **openpyxl** engine: cells (values + formatting + data type + hyperlinks), conditional formatting, filters, frozen panes. |
| `ubuntu_eval/uno_engine.py` | **LibreOffice UNO** engine: charts and pivot tables (needs recomputed formula/series/pivot data). |
| `ubuntu_eval/soffice.py` | Headless-LibreOffice instance manager + UNO bootstrap (one private instance per worker process). |
| `ubuntu_eval/logging_utils.py` | Multiprocess-safe `QueueListener` logging. |
| `ubuntu_eval/selftest.py` | Smoke test: re-evaluates the bundled example logs and checks against the Windows results. |
| `requirements_ubuntu.txt`, `setup_ubuntu_eval.sh` | Dependencies + one-shot installer. |

## How the outcome-based evaluation works (recap)

For every task, the evaluator loads each reference solution `*_gt*.xlsx` and its
`*_check.yaml`.  The check-board maps each **sheet index** (`'1'`, `'2'`, …) to a
set of property flags across six categories:

- `cells` — `values`, `formatting`, `hyperlink`
- `charts` — `chart_type`, `title`, `legend`, `axes`, `series`, `trendlines`, …
- `pivot_tables` — `source`, `rows`, `columns`, `values`, `filters`, …
- `filters` — post-AutoFilter visible range
- `format_conditions` — conditional formatting (`formula1`, colours, font, …)
- `view` — `freeze_pane`

Only flagged properties are compared. A task **passes** (`Pass@1`) if **any** of
its reference solutions matches; `Exec@1` counts whether the run executed without
error (`Success Count > 0` and a result file exists). These semantics, the
header-matched column comparison, the numeric tolerance (`1e-8`), the
"data type is always enforced once values match" behaviour, the chart/pivot
exhaustive matching and the freeze-pane gating are all reproduced faithfully.

## The hybrid backend

| Category | Engine | Why |
|----------|--------|-----|
| cells, format_conditions, filters, view | **openpyxl** | Excel caches formula results in the saved `.xlsx`, so cached values + style XML are exact and fast — no office process needed. |
| charts, pivot_tables | **headless LibreOffice (UNO)** | These need recomputed series data and materialised pivot (DataPilot) fields, which openpyxl cannot provide. |

Key faithfulness principle: both the ground-truth and result workbooks are
Excel-saved `.xlsx` files, so *any consistent* extraction (openpyxl or the same
LibreOffice build on both files) yields the same relative verdict as the Windows
Excel-COM code. LibreOffice is only started when a task actually checks charts or
pivot tables, and each worker reuses one private instance.

## Setup

From this `agent/` directory (use `sudo` if not root):

```bash
bash setup_ubuntu_eval.sh
```

This installs LibreOffice Calc + `python3-uno` (apt) and the Python deps (pip),
then runs the smoke test. To install without the test: `bash setup_ubuntu_eval.sh --no-test`.

## Running

Use the **same** `config/config.yaml` as the Windows evaluator (only the `path.*`
and `repeat` fields are read; `worker` sets the default process count):

```bash
# default: workers from config, [Order]_[Sheet Name] folder naming
python3 evaluation_ubuntu.py -c config/config.yaml

# override worker count
python3 evaluation_ubuntu.py -c config/config.yaml --workers 8

# [No.]_[Sheet Name] folder naming (equivalent to USE_NO_AND_SHEETNAME=True)
python3 evaluation_ubuntu.py -c config/config.yaml --use-no-and-sheetname

# openpyxl only — skip charts/pivot tables, no LibreOffice (fast debug)
python3 evaluation_ubuntu.py -c config/config.yaml --no-uno
```

### Output

- Metrics + per-task verdicts are written to `<save_path>/eval_result_ubuntu.yaml`
  (kept separate from the Windows `eval_result.yaml`).
- Interleaved, per-process logs go to `<save_path>/eval_ubuntu.log` and stderr.
- The run is **checkpointed after every task**: re-running resumes and skips
  already-evaluated tasks. Delete `eval_result_ubuntu.yaml` to re-evaluate from
  scratch.

## Validation

Two independent checks back the port's reliability.

### 1. Windows-oracle match (`selftest`)

`python3 -m ubuntu_eval.selftest` re-evaluates the bundled
`SheetCopilot_example_logs/` and confirms the Ubuntu backend reproduces the
exact per-task Exec@1 / Pass@1 verdicts from the example `eval_result.yaml`
(which was generated by the original **Windows/Excel** evaluator). On the 12
bundled tasks — spanning cells, formulas, formatting, pivot tables and charts —
both produce **Exec@1 = 0.833, Pass@1 = 0.5** with identical pass/fail per task.

### 2. Op / non-op reliability test (`reliability_test`)

`python3 -m ubuntu_eval.reliability_test [--gt <dir>] [--workers N]` checks both
failure modes a comparator can have:

- **op (positive)** — every ground truth vs. *itself* with its own check-board
  must pass (it must not reject correct results);
- **non-op (negative)** — every ground truth vs. the *raw source* spreadsheet
  (`dataset/task_sheets/<SheetName>.xlsx`, i.e. before any task operation) must
  fail (it must not trivially accept everything).

Results on both ground-truth sets:

| GT set | OP pass | NON-OP pass | Errors |
|--------|---------|-------------|--------|
| `task_sheet_answers_v2` | 390/390 (100%) | 4/390 (1.0%) | 0 |
| `task_sheet_answers` (v1) | 253/253 (100%) | 4/253 (1.6%) | 0 |

The ~1% of non-op "passes" are **faithful to the original** (not port bugs): each
is a property of the benchmark's own check-boards that the Windows evaluator
reproduces identically — e.g. a reference with zero conditional-formatting rules
matching a source that also has none (`3_DemographicProfile`), a check-board that
only covers input columns already present in the source (`4_FutureValue`), the
`A1.End(xlDown)` contiguous-range edge case (`6_SmallBalanceSheet`), and a
check-board that references a non-existent sheet so the check is skipped
(`8_WeeklySales`).

## Notes & limitations

- Result/ground-truth `.xlsx` files must contain Excel's cached formula values
  (true for the dataset GTs and for any file an Excel-based agent saves). If a
  future Linux-based agent emits files *without* cached values, open and re-save
  them once through LibreOffice (`soffice --headless --convert-to xlsx`) so the
  caches are populated before evaluation.
- The same workbook-layout rules as the original apply: new sheets go to the
  left of the first sheet, and each sheet's table must start at `A1` and be
  contiguous.

## Alternative agent & inspection tools

Three additional tools support model testing and failure analysis on Linux:

### `run_planning_probe.py` — open-loop model probe
Tests any OpenAI-compatible model's SheetCopilot *planning* on N tasks **without Excel** (the full agent is closed-loop and needs Excel). It builds the real planning prompt + a sheet-state read from the source workbook, asks for a plan, and saves a trajectory per task. Parses both the `@API(...)@` form and the Qwen3 `<tool_call>` form.
```
python run_planning_probe.py --base-url <url/v1> --api-key <key> --model <name> --num 10
```

### `claude_code_agent.py` — Claude Code as an alternative agent
Hands each task to the **Claude Code CLI**, which edits the workbook with Python/openpyxl/LibreOffice, and writes results in the exact layout `evaluation_ubuntu.py` expects (so both agents are scored identically) plus a trajectory JSON.
```
ANTHROPIC_BASE_URL=<gateway> ANTHROPIC_API_KEY=<token> \
python claude_code_agent.py --claude /path/to/claude --model <name> \
    --num 5 --permission-mode default --allowed-tools "Bash Read Write Edit Glob Grep"
```
Notes: Claude Code refuses `bypassPermissions` as **root**, so use `--permission-mode default` with an `--allowed-tools` allow-list (pre-approved tools run without prompts and need no bypass). Point it at an Anthropic-Messages-compatible gateway. Then evaluate the `save_path` with `evaluation_ubuntu.py`.

### `visualize_trajectories.py` — trajectory visualizer
Renders all `*_trajectory.json` under a directory into one self-contained HTML page: per trajectory it shows the **state-machine path**, and per call the **query / response / reasoning / token usage / latency / parsed+executed actions / errors**. If an `eval_result_ubuntu.yaml` is present it overlays the true **outcome** pass/fail so you can jump to failures and see why.
```
python visualize_trajectories.py -d <results_dir>     # writes <results_dir>/trajectories.html
```
