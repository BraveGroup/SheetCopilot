"""
Reliability test for the Ubuntu evaluator: op (positive) + non-op (negative).

A faithful evaluator must do two things well:

* **op / positive** -- a correct result must be accepted.  We use the strongest
  positive case: every ground truth compared *against itself* with its own
  check-board must pass.  (Pass rate should be ~100%.)
* **non-op / negative** -- an *unmodified* workbook must be rejected.  We compare
  every ground truth against the *raw source* spreadsheet
  (``dataset/task_sheets/<SheetName>.xlsx``), i.e. the workbook before any task
  operation was applied, using the GT's own check-board.  Because the task's
  operations are missing, this must (almost always) fail.  (Pass rate should be
  ~0%; the gap between op and non-op pass rates is the evaluator's
  discrimination power.)

This guards against both failure modes: an evaluator that wrongly rejects correct
results, and -- more dangerously -- one that trivially accepts everything.

Run from the ``agent`` directory::

    python -m ubuntu_eval.reliability_test
    python -m ubuntu_eval.reliability_test --gt ../dataset/task_sheet_answers --workers 8
"""

import argparse
import glob
import os
import sys
from collections import Counter

import yaml

HERE = os.path.dirname(os.path.abspath(__file__))
AGENT_DIR = os.path.dirname(HERE)
sys.path.insert(0, AGENT_DIR)


def _build_jobs(gt_root, source_root):
    jobs = []
    for chk in glob.glob(os.path.join(gt_root, "**", "*_check.yaml"), recursive=True):
        gt = chk.replace("_check.yaml", ".xlsx")
        if not os.path.exists(gt) or "$" in os.path.basename(gt):
            continue
        # Sheet name is the first path component under gt_root.
        rel = os.path.relpath(gt, gt_root)
        sheet_name = rel.split(os.sep)[0]
        source = os.path.join(source_root, sheet_name + ".xlsx")
        if not os.path.exists(source):
            continue
        jobs.append((gt, chk, source))
    return sorted(jobs)


def _worker_init(queue=None):
    # Lazily import so a missing LibreOffice doesn't break import of this module.
    import atexit

    from ubuntu_eval.soffice import shutdown_process_instance

    atexit.register(shutdown_process_instance)


def _run_one(job):
    """Return (name, op_pass, nonop_pass, nonop_report, error)."""
    from ubuntu_eval import compare_workbooks

    gt, chk, source = job
    name = os.path.relpath(gt, AGENT_DIR)
    try:
        cb = yaml.safe_load(open(chk))["check_board"]
    except Exception as e:
        return (name, None, None, None, f"check load: {e}")

    error = ""
    try:
        _r, op_pass = compare_workbooks(gt, gt, cb)
    except Exception as e:
        op_pass, error = None, error + f" [op: {e}]"

    try:
        nonop_report, nonop_pass = compare_workbooks(gt, source, cb)
    except Exception as e:
        nonop_pass, nonop_report = None, None
        error += f" [nonop: {e}]"

    return (name, op_pass, nonop_pass, _passed_categories(cb), error)


def _passed_categories(cb):
    """Which categories the check-board exercises (for breakdown stats)."""
    cats = set()
    for conf in cb.values():
        if not isinstance(conf, dict):
            continue
        for cat, props in conf.items():
            if isinstance(props, dict) and any(props.values()):
                cats.add(cat)
    return tuple(sorted(cats))


def main():
    parser = argparse.ArgumentParser(description="Op/non-op reliability test.")
    parser.add_argument("--gt", default=os.path.join(AGENT_DIR, "..", "dataset", "task_sheet_answers_v2"))
    parser.add_argument("--source", default=os.path.join(AGENT_DIR, "..", "dataset", "task_sheets"))
    parser.add_argument("--workers", "-w", type=int, default=8)
    args = parser.parse_args()

    jobs = _build_jobs(args.gt, args.source)
    print(f"Reliability test over {len(jobs)} ground truths")
    print(f"  GT root : {os.path.abspath(args.gt)}")
    print(f"  source  : {os.path.abspath(args.source)}\n")

    import multiprocessing as mp

    results = []
    ctx = mp.get_context("spawn")
    with ctx.Pool(processes=max(1, args.workers), initializer=_worker_init) as pool:
        import tqdm

        for r in tqdm.tqdm(pool.imap_unordered(_run_one, jobs), total=len(jobs)):
            results.append(r)

    op_pass = sum(1 for _, o, _, _, _ in results if o is True)
    op_fail = sum(1 for _, o, _, _, _ in results if o is False)
    nonop_pass = sum(1 for _, _, n, _, _ in results if n is True)
    nonop_fail = sum(1 for _, _, n, _, _ in results if n is False)
    errors = [(name, e) for name, _, _, _, e in results if e]
    n = len(results)

    print("\n================ Reliability summary ================")
    print(f"  OP (GT vs GT)         : {op_pass}/{n} pass  ({100*op_pass/n:.1f}%)   "
          f"[want ~100%]   fail={op_fail}")
    print(f"  NON-OP (GT vs source) : {nonop_pass}/{n} pass  ({100*nonop_pass/n:.1f}%)   "
          f"[want ~0%]   fail={nonop_fail}")
    print(f"  Discrimination gap    : {100*(op_pass-nonop_pass)/n:.1f} points")
    print(f"  Errors                : {len(errors)}")

    # Op failures are bugs (a correct file rejected).
    op_failures = [name for name, o, _, _, _ in results if o is False]
    if op_failures:
        print("\n  OP FAILURES (correct file wrongly rejected -- investigate):")
        for name in op_failures[:30]:
            print("    ", name)

    # Non-op passes are false positives (raw source wrongly accepted).
    fps = [(name, cats) for name, _, n_, cats, _ in results if n_ is True]
    if fps:
        print(f"\n  NON-OP FALSE POSITIVES ({len(fps)}) -- raw source accepted as correct:")
        cat_counter = Counter()
        for name, cats in fps:
            cat_counter[cats] += 1
        for cats, c in cat_counter.most_common():
            print(f"    {c:3d}  categories={cats}")
        for name, cats in fps[:20]:
            print("      e.g.", name, cats)

    if errors:
        print("\n  ERRORS:")
        for name, e in errors[:20]:
            print("    ", name, "->", e.strip())

    # Exit non-zero if discrimination is poor.
    ok = (op_pass == n) and (nonop_pass / n < 0.05) and (len(errors) == 0)
    print("\n  RESULT:", "RELIABLE" if ok else "REVIEW NEEDED")
    return 0 if ok else 1


if __name__ == "__main__":
    raise SystemExit(main())
