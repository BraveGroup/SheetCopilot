"""
Self-test / smoke test for the Ubuntu evaluator.

It re-evaluates the bundled ``SheetCopilot_example_logs`` (real Excel-generated
results shipped with the repo) and checks that the Ubuntu backend reproduces the
per-task Exec@1 / Pass@1 verdicts stored in the example ``eval_result.yaml``
(which was produced by the original Windows/Excel evaluator).  A clean run prints
``SELFTEST PASSED``.

Run from the ``agent`` directory::

    python -m ubuntu_eval.selftest
"""

import os
import sys

import yaml

HERE = os.path.dirname(os.path.abspath(__file__))
AGENT_DIR = os.path.dirname(HERE)
sys.path.insert(0, AGENT_DIR)

from ubuntu_eval import compare_workbooks  # noqa: E402
from ubuntu_eval.soffice import shutdown_process_instance  # noqa: E402


def _to_set(s):
    if isinstance(s, list):
        return set(s)
    return set(x.strip() for x in str(s).split(",") if x.strip())


def main():
    logs_dir = os.path.join(AGENT_DIR, "SheetCopilot_example_logs")
    gt_path = os.path.join(AGENT_DIR, "..", "dataset", "task_sheet_answers_v2")
    expected_file = os.path.join(logs_dir, "eval_result.yaml")

    if not os.path.exists(expected_file):
        print("SELFTEST SKIPPED: example eval_result.yaml not found")
        return 0

    expected = yaml.safe_load(open(expected_file))["check_result_each_repeat"][1]
    expected_pass = _to_set(expected["success_list"])

    import pandas as pd

    df = pd.read_excel(os.path.join(AGENT_DIR, "..", "dataset", "dataset.xlsx"))

    got_pass = set()
    n_checked = 0
    try:
        for index, row in df.iterrows():
            name = f"{index + 1}_{row['Sheet Name']}"
            task_dir = os.path.join(logs_dir, name)
            if not os.path.isdir(task_dir):
                continue
            res = os.path.join(task_dir, f"{name}_1.xlsx")
            log_file = os.path.join(task_dir, f"{name}_log.yaml")
            if not (os.path.exists(res) and os.path.exists(log_file)):
                continue
            task_log = yaml.load(open(log_file, encoding="utf-8"), Loader=yaml.Loader)
            n_checked += 1
            gt_folder = os.path.join(gt_path, row["Sheet Name"], f"{row['No.']}_{row['Sheet Name']}")
            for gt_file in sorted(os.listdir(gt_folder)):
                if not gt_file.endswith(".xlsx") or "$" in gt_file:
                    continue
                gt = os.path.join(gt_folder, gt_file)
                cb = yaml.safe_load(open(gt.replace(".xlsx", "_check.yaml")))["check_board"]
                _r, ok = compare_workbooks(gt, res, cb)
                if ok and len(task_log.get("Success Response", [])) > 0:
                    got_pass.add(name)
                    break
    finally:
        shutdown_process_instance()

    print(f"checked={n_checked}  expected_pass={len(expected_pass)}  got_pass={len(got_pass)}")
    if got_pass == expected_pass:
        print("SELFTEST PASSED")
        return 0
    print("SELFTEST FAILED")
    print("  expected-only:", sorted(expected_pass - got_pass))
    print("  got-only:", sorted(got_pass - expected_pass))
    return 1


if __name__ == "__main__":
    raise SystemExit(main())
