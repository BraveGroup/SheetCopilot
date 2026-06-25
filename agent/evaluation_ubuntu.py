"""
Ubuntu/Linux outcome-based evaluation for SheetCopilot.

This is a parallel, Excel-free counterpart to ``agent/evaluation.py``.  It reads
the same config file, the same ``dataset.xlsx`` task list, the same result
folders and the same ``*_check.yaml`` ground-truth check-boards, and produces the
same metrics (Exec@1, Pass@1, A_mean / A50_norm / A90_norm, plus per-category
breakdowns).  The comparison itself is delegated to :mod:`ubuntu_eval`, a hybrid
openpyxl + headless-LibreOffice backend, so no Windows or Excel is required.

The original Windows scripts (``evaluation.py``, ``utils/compare_sheets.py``) are
left completely untouched; this file lives alongside them.

Highlights
----------
* **Multiprocessing**: one worker process per ``worker`` (config) / ``--workers``;
  each worker that needs charts or pivot tables lazily spins up and reuses its
  own private headless LibreOffice instance.
* **Logging**: a shared ``QueueListener`` writes interleaved, per-process logs to
  both stderr and ``<save_path>/eval_ubuntu.log``.
* **Checkpoint / resume**: per-task outcomes are persisted to
  ``<save_path>/eval_result_ubuntu.yaml`` after every task; re-running skips
  already-evaluated tasks.  Delete that file to re-evaluate from scratch.

Usage
-----
    python evaluation_ubuntu.py -c config/config.yaml
    python evaluation_ubuntu.py -c config/config.yaml --workers 8
    python evaluation_ubuntu.py -c config/config.yaml --use-no-and-sheetname
"""

import argparse
import logging
import os
import time
from collections import defaultdict
from datetime import datetime
from functools import partial

import numpy as np
import pandas as pd
import tqdm
import yaml

# Allow "python evaluation_ubuntu.py" from the agent/ folder and "-m" alike.
import sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from ubuntu_eval import compare_workbooks  # noqa: E402
from ubuntu_eval.logging_utils import configure_process, start_listener  # noqa: E402
from ubuntu_eval.soffice import shutdown_process_instance  # noqa: E402

RESULT_FILENAME = "eval_result_ubuntu.yaml"
LOG_FILENAME = "eval_ubuntu.log"

# Globals populated in each worker by ``_worker_init``.
_G = {}


# --------------------------------------------------------------------------- #
# Worker
# --------------------------------------------------------------------------- #
def _worker_init(queue, gt_path, save_path, use_no_and_sheetname, no_uno):
    configure_process(queue)
    _G["gt_path"] = gt_path
    _G["save_path"] = save_path
    _G["use_no_and_sheetname"] = use_no_and_sheetname
    _G["no_uno"] = no_uno
    # Make sure each worker tears down its LibreOffice instance on exit.
    import atexit

    atexit.register(shutdown_process_instance)


def _task_name(index, row, use_no_and_sheetname):
    if use_no_and_sheetname:
        return f"{row['No.']}_{row['Sheet Name']}"
    return f"{index + 1}_{row['Sheet Name']}"


def _count_plan_actions(log, repeat_id):
    """Number of actions in the (refined) plan, mirroring ``evaluation.py``."""
    try:
        plan = log["Success Response"][repeat_id - 1]["refined response"]
        return sum(len(steps) for steps in plan)
    except Exception:
        return 0


def evaluate_one_task(job):
    """Evaluate a single (repeat, task) pair. Returns a record dict.

    ``status`` is one of: ``missing`` (no folder/log/result), ``error``
    (exception while comparing) or ``ok``.
    """
    repeat_id, index, row = job
    gt_path = _G["gt_path"]
    save_path = _G["save_path"]
    log = logging.getLogger("eval")

    task_name = _task_name(index, row, _G["use_no_and_sheetname"])
    record = {
        "task_name": task_name,
        "repeat_id": repeat_id,
        "cates": [c for c in str(row["Categories"]).split(", ")],
        "exec_success": False,
        "success": False,
        "matched_gt": None,
        "num_acts": None,
        "gt_min_acts": None,
        "status": "ok",
        "error": "",
    }

    task_dir = os.path.join(save_path, task_name)
    if not os.path.isdir(task_dir):
        record["status"] = "missing"
        return record

    res_path = os.path.join(task_dir, f"{task_name}_{repeat_id}.xlsx")
    log_file = os.path.join(task_dir, f"{os.path.basename(task_dir)}_log.yaml")
    if not os.path.exists(log_file):
        record["status"] = "missing"
        return record

    try:
        with open(log_file, "r", encoding="utf-8") as f:
            task_log = yaml.load(f, Loader=yaml.Loader)
    except Exception as e:
        record["status"] = "error"
        record["error"] = f"log load failed: {e}"
        return record

    res_exists = os.path.exists(res_path)
    if task_log.get("Success Count", 0) > 0 and res_exists:
        record["exec_success"] = True

    if not res_exists:
        log.info("[%s] no result xlsx -> exec/pass = False", task_name)
        return record

    gt_folder = os.path.join(gt_path, row["Sheet Name"], f"{row['No.']}_{row['Sheet Name']}")
    if not os.path.isdir(gt_folder):
        record["status"] = "error"
        record["error"] = f"gt folder missing: {gt_folder}"
        return record

    gt_files = [
        x for x in os.listdir(gt_folder)
        if x.endswith(".xlsx") and "$" not in x
    ]

    matched = False
    for gt_file in sorted(gt_files):
        gt = os.path.join(gt_folder, gt_file)
        check_yaml = os.path.join(gt_folder, gt_file.replace(".xlsx", "_check.yaml"))
        if not os.path.exists(check_yaml):
            continue
        try:
            with open(check_yaml, "r") as f:
                check_boards = yaml.load(f, Loader=yaml.Loader)["check_board"]
        except Exception as e:
            record["error"] += f" [check load {gt_file}: {e}]"
            continue

        try:
            _report, ok = compare_workbooks(
                gt, res_path, check_boards, disable_uno=_G["no_uno"]
            )
        except Exception as e:
            record["status"] = "error"
            record["error"] += f" [compare {gt_file}: {e}]"
            log.warning("[%s] compare error vs %s: %s", task_name, gt_file, e)
            continue

        if ok and len(task_log.get("Success Response", [])) > 0:
            matched = True
            record["success"] = True
            record["matched_gt"] = gt_file
            record["num_acts"] = _count_plan_actions(task_log, repeat_id)
            gt_actions = [x for x in str(row["Atomic actions"]).split(",") if "function" not in x]
            record["gt_min_acts"] = len(gt_actions)
            log.info("[%s] PASS (matched %s)", task_name, gt_file)
            break

    if not matched:
        log.info("[%s] exec=%s pass=False", task_name, record["exec_success"])
    return record


# --------------------------------------------------------------------------- #
# Aggregation
# --------------------------------------------------------------------------- #
def _aggregate(records, repeat_id):
    """Turn per-task records into the metric block mirroring ``evaluation.py``."""
    checked, exec_success, success = [], [], []
    checked_by_cate = defaultdict(list)
    exec_by_cate = defaultdict(list)
    success_by_cate = defaultdict(list)
    action_cnt, gt_min_cnt, matched_gt_lst, errors = [], [], [], []

    for rec in records:
        if rec["status"] == "missing":
            continue
        name = rec["task_name"]
        checked.append(name)
        for c in rec["cates"]:
            checked_by_cate[c].append(name)
        if rec["error"]:
            errors.append(f"{name}: {rec['error'].strip()}")
        if rec["exec_success"]:
            exec_success.append(name)
            for c in rec["cates"]:
                exec_by_cate[c].append(name)
        if rec["success"]:
            success.append(name)
            for c in rec["cates"]:
                success_by_cate[c].append(name)
            if rec["num_acts"] is not None:
                action_cnt.append(rec["num_acts"])
                gt_min_cnt.append(rec["gt_min_acts"])
            if rec["matched_gt"]:
                matched_gt_lst.append(rec["matched_gt"])

    total = len(checked)
    eval_results = {"Total": total}
    if total:
        eval_results["Exec@1"] = len(exec_success) / total
        eval_results["Pass@1"] = len(success) / total
        for k, v in exec_by_cate.items():
            cate_total = len(checked_by_cate[k])
            eval_results[f"{k} Exec & Pass"] = "{:d}/{:d} & {:d}/{:d}".format(
                len(v), cate_total, len(success_by_cate[k]), cate_total
            )
        if action_cnt:
            ac = np.array(action_cnt, dtype=float)
            gc = np.array(gt_min_cnt, dtype=float)
            eval_results["A_mean"] = float(np.mean(ac))
            eval_results["A50_norm"] = float(np.median(ac / gc))
            eval_results["A90_norm"] = float(np.percentile(ac / gc, 90))

    return {
        "eval_results": eval_results,
        "matched_gt_lst": matched_gt_lst,
        "checked_list": ", ".join(checked),
        "exec_success_list": ", ".join(exec_success),
        "success_list": ", ".join(success),
        "checked_list_by_cate": {k: ", ".join(v) for k, v in checked_by_cate.items()},
        "exec_success_list_by_cate": {k: ", ".join(v) for k, v in exec_by_cate.items()},
        "success_list_by_cate": {k: ", ".join(v) for k, v in success_by_cate.items()},
        "action_cnt_list": ", ".join(str(x) for x in action_cnt),
        "gt_min_action_cnt_list": ", ".join(str(x) for x in gt_min_cnt),
        "error_log": errors,
    }


def _save(eval_result, path):
    tmp = path + ".tmp"
    with open(tmp, "w") as f:
        yaml.dump(eval_result, f, allow_unicode=True)
    os.replace(tmp, path)


# --------------------------------------------------------------------------- #
# Driver
# --------------------------------------------------------------------------- #
def evaluate(config, workers, use_no_and_sheetname, no_uno):
    task_path = config["path"]["task_path"]
    gt_path = config["path"]["gt_path"]
    save_path = config["path"]["save_path"]
    repeat = config.get("repeat", 1)

    result_path = os.path.join(save_path, RESULT_FILENAME)
    log_path = os.path.join(save_path, LOG_FILENAME)
    os.makedirs(save_path, exist_ok=True)

    queue, listener = start_listener(log_path)
    log = logging.getLogger("eval")
    configure_process(queue)  # main process logs through the same sink

    log.info("Evaluating results at %s", save_path)
    log.info("Ground truths: %s | workers: %d | uno: %s", gt_path, workers, not no_uno)

    if os.path.exists(result_path):
        with open(result_path, "r") as f:
            eval_result = yaml.load(f, Loader=yaml.Loader) or {}
    else:
        eval_result = {}
    eval_result.setdefault("check_result_each_repeat", {})
    eval_result.setdefault("_records", {})  # {repeat_id: {task_name: record}}

    task_df = pd.read_excel(task_path, header=0)

    try:
        for repeat_id in range(1, repeat + 1):
            t0 = time.time()
            done = eval_result["_records"].setdefault(repeat_id, {})

            # Build the job list: tasks whose result folder exists and that have
            # not yet been evaluated for this repeat (checkpoint/resume).
            jobs = []
            for index, row in task_df.iterrows():
                name = _task_name(index, row, use_no_and_sheetname)
                if name in done:
                    continue
                if not os.path.isdir(os.path.join(save_path, name)):
                    continue
                jobs.append((repeat_id, index, dict(row)))

            log.info("Repeat %d: %d task(s) to evaluate (%d cached)",
                     repeat_id, len(jobs), len(done))

            if jobs:
                init = partial(
                    _worker_init,
                    gt_path=gt_path,
                    save_path=save_path,
                    use_no_and_sheetname=use_no_and_sheetname,
                    no_uno=no_uno,
                )
                import multiprocessing as mp

                ctx = mp.get_context("spawn")
                with ctx.Pool(processes=max(1, workers), initializer=init,
                              initargs=(queue,)) as pool:
                    it = pool.imap_unordered(evaluate_one_task, jobs)
                    for rec in tqdm.tqdm(it, total=len(jobs),
                                         desc=f"Repeat {repeat_id}"):
                        done[rec["task_name"]] = rec
                        # Persist after each task so a crash loses nothing.
                        eval_result["check_result_each_repeat"][repeat_id] = _aggregate(
                            list(done.values()), repeat_id
                        )
                        _save(eval_result, result_path)

            # Final aggregation for this repeat.
            agg = _aggregate(list(done.values()), repeat_id)
            eval_result["check_result_each_repeat"][repeat_id] = agg
            _save(eval_result, result_path)

            log.info("Repeat %d finished in %.1fs", repeat_id, time.time() - t0)
            for k, v in agg["eval_results"].items():
                log.info("  %s: %s", k, v)
            if agg["error_log"]:
                log.warning("Errors (%d): \n%s", len(agg["error_log"]),
                            "\n".join(agg["error_log"]))
    finally:
        listener.stop()

    print("\n{} evaluated against {} at {}".format(
        save_path, gt_path, datetime.now().strftime("%H:%M:%S")))


def main():
    parser = argparse.ArgumentParser(description="Ubuntu (Excel-free) SheetCopilot evaluation.")
    parser.add_argument("--config", "-c", default="./config/config.yaml", type=str,
                        help="path to the config yaml")
    parser.add_argument("--workers", "-w", type=int, default=None,
                        help="number of worker processes (default: config 'worker')")
    parser.add_argument("--use-no-and-sheetname", action="store_true",
                        help="use the [No.]_[Sheet Name] folder naming convention")
    parser.add_argument("--no-uno", action="store_true",
                        help="skip charts/pivot tables (openpyxl-only; no LibreOffice)")
    args = parser.parse_args()

    with open(args.config, "r") as f:
        config = yaml.load(f, Loader=yaml.Loader)

    workers = args.workers if args.workers is not None else config.get("worker", 1)
    evaluate(config, workers, args.use_no_and_sheetname, args.no_uno)


if __name__ == "__main__":
    main()
