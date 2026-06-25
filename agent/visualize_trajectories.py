"""
Trajectory visualizer for SheetCopilot.

Scans a directory for ``*_trajectory.json`` files (produced by
``utils/trajectory.py`` -- both the live agent and the planning probe write this
format) and renders a single, self-contained, dependency-free HTML page that lets
you inspect, for every trajectory:

* the **state-machine path** the agent took (one badge per LLM call, coloured by
  stage, ending in the final ``end`` / ``fail`` state);
* each call's **query** (all messages sent), **response**, hidden **reasoning**,
  **token usage** and **latency**;
* the **parsed** and **executed** actions, with execution success/failure and any
  errors highlighted -- so you can see *why* a task failed at a glance.

The page is a single ``.html`` file with the data embedded (no server, no network,
no build step). Open it directly in a browser.

Usage::

    python visualize_trajectories.py --results-dir ./planning_probe_IQuest-Coder
    python visualize_trajectories.py --results-dir <save_path> --out report.html
"""

import argparse
import glob
import html
import json
import os


def load_eval_verdicts(results_dir):
    """If an ``eval_result_ubuntu.yaml`` exists, return
    ``{task_name: 'pass'|'fail'}`` for the tasks the evaluator checked, so the
    visualizer can show *outcome* pass/fail (not just "did the agent run")."""
    verdicts = {}
    path = os.path.join(results_dir, "eval_result_ubuntu.yaml")
    if not os.path.exists(path):
        return verdicts
    try:
        import yaml
        data = yaml.safe_load(open(path))
        rep = data.get("check_result_each_repeat", {}).get(1, {})
        def to_set(s):
            if isinstance(s, list):
                return set(s)
            return set(x.strip() for x in str(s or "").split(",") if x.strip())
        checked = to_set(rep.get("checked_list"))
        passed = to_set(rep.get("success_list"))
        for t in checked:
            verdicts[t] = "pass" if t in passed else "fail"
    except Exception as e:
        print(f"  (could not read eval verdicts: {e})")
    return verdicts


def load_trajectories(results_dir):
    verdicts = load_eval_verdicts(results_dir)
    items = []
    for path in sorted(glob.glob(os.path.join(results_dir, "**", "*_trajectory.json"), recursive=True)):
        try:
            with open(path, encoding="utf-8") as f:
                data = json.load(f)
        except Exception as e:
            print(f"  skip {path}: {e}")
            continue
        meta = data.get("meta", {}) or {}
        calls = data.get("calls", []) or []
        name = meta.get("task") or os.path.basename(os.path.dirname(path))
        items.append({
            "file": os.path.relpath(path, results_dir),
            "name": name,
            "meta": meta,
            "calls": calls,
            # Outcome verdict from the evaluator if available, else None.
            "eval": verdicts.get(name),
        })
    return items


HTML_TEMPLATE = r"""<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>SheetCopilot Trajectories</title>
<style>
  :root {
    --bg:#0f1115; --panel:#171a21; --panel2:#1e222b; --border:#2a2f3a;
    --text:#e6e9ef; --muted:#9aa3b2; --accent:#5b9dff; --green:#3fb950; --red:#f85149;
    --amber:#d29922; --purple:#a371f7; --cyan:#39c5cf;
  }
  * { box-sizing:border-box; }
  body { margin:0; font:14px/1.5 -apple-system,Segoe UI,Roboto,Helvetica,Arial,sans-serif;
         background:var(--bg); color:var(--text); }
  header { padding:14px 20px; border-bottom:1px solid var(--border); background:var(--panel);
           display:flex; align-items:center; gap:18px; flex-wrap:wrap; position:sticky; top:0; z-index:5; }
  header h1 { font-size:16px; margin:0; font-weight:650; }
  header .stat { color:var(--muted); font-size:13px; }
  header .stat b { color:var(--text); }
  .layout { display:flex; height:calc(100vh - 52px); }
  .sidebar { width:320px; min-width:260px; border-right:1px solid var(--border); overflow-y:auto;
             background:var(--panel); }
  .controls { padding:10px; border-bottom:1px solid var(--border); position:sticky; top:0; background:var(--panel); }
  .controls input, .controls select { width:100%; padding:7px 9px; margin-bottom:6px; background:var(--panel2);
             color:var(--text); border:1px solid var(--border); border-radius:6px; }
  .task { padding:9px 12px; border-bottom:1px solid var(--border); cursor:pointer; display:flex;
          align-items:center; gap:8px; }
  .task:hover { background:var(--panel2); }
  .task.active { background:#243049; border-left:3px solid var(--accent); }
  .task .nm { flex:1; font-weight:550; font-size:13px; overflow:hidden; text-overflow:ellipsis; white-space:nowrap; }
  .task .sub { color:var(--muted); font-size:11.5px; }
  .pill { font-size:11px; padding:1px 7px; border-radius:10px; font-weight:600; white-space:nowrap; }
  .pill.pass { background:rgba(63,185,80,.16); color:var(--green); }
  .pill.fail { background:rgba(248,81,73,.16); color:var(--red); }
  .main { flex:1; overflow-y:auto; padding:18px 24px; }
  .meta-card { background:var(--panel); border:1px solid var(--border); border-radius:10px; padding:14px 16px; margin-bottom:16px; }
  .meta-card h2 { margin:0 0 8px; font-size:15px; }
  .meta-grid { display:grid; grid-template-columns:repeat(auto-fit,minmax(150px,1fr)); gap:6px 18px; color:var(--muted); font-size:12.5px; }
  .meta-grid b { color:var(--text); font-weight:600; }
  .instr { margin-top:10px; padding:10px 12px; background:var(--panel2); border-radius:8px; white-space:pre-wrap; }
  .path { display:flex; flex-wrap:wrap; align-items:center; gap:6px; margin:14px 0 22px; }
  .node { padding:4px 10px; border-radius:7px; font-size:12px; font-weight:600; border:1px solid var(--border); }
  .arrow { color:var(--muted); }
  .call { background:var(--panel); border:1px solid var(--border); border-radius:10px; margin-bottom:14px; overflow:hidden; }
  .call.err { border-color:var(--red); }
  .call-head { display:flex; align-items:center; gap:10px; padding:10px 14px; background:var(--panel2);
               cursor:pointer; flex-wrap:wrap; }
  .call-head .idx { color:var(--muted); font-variant-numeric:tabular-nums; }
  .stage { font-size:11.5px; padding:2px 9px; border-radius:7px; font-weight:700; letter-spacing:.2px; }
  .call-head .spacer { flex:1; }
  .chip { font-size:11px; color:var(--muted); padding:1px 7px; border:1px solid var(--border); border-radius:10px; }
  .chip.ok { color:var(--green); border-color:rgba(63,185,80,.4); }
  .chip.bad { color:var(--red); border-color:rgba(248,81,73,.4); }
  .call-body { padding:6px 14px 14px; display:none; }
  .call-body.open { display:block; }
  .section { margin-top:12px; }
  .section .lbl { font-size:11.5px; text-transform:uppercase; letter-spacing:.5px; color:var(--muted); margin-bottom:5px; }
  pre { margin:0; background:#0b0d11; border:1px solid var(--border); border-radius:8px; padding:10px 12px;
        white-space:pre-wrap; word-break:break-word; max-height:340px; overflow:auto; font:12.5px/1.5 ui-monospace,SFMono-Regular,Menlo,monospace; }
  .msg { border-left:3px solid var(--border); padding-left:10px; margin-bottom:8px; }
  .msg .role { font-size:11px; font-weight:700; color:var(--accent); text-transform:uppercase; }
  .msg.system .role { color:var(--purple); } .msg.assistant .role { color:var(--green); } .msg.user .role { color:var(--cyan); }
  .acts { display:flex; flex-wrap:wrap; gap:6px; }
  .act { font:12px ui-monospace,monospace; background:#0b0d11; border:1px solid var(--border); border-radius:6px; padding:3px 8px; }
  .act.bad { border-color:var(--red); color:var(--red); }
  .errbox { background:rgba(248,81,73,.1); border:1px solid var(--red); color:#ffb4ae; border-radius:8px; padding:8px 12px; margin-top:10px; white-space:pre-wrap; }
  details { margin-top:6px; } summary { cursor:pointer; color:var(--muted); font-size:12px; }
  .empty { color:var(--muted); padding:40px; text-align:center; }
  .toggle { color:var(--accent); cursor:pointer; font-size:12px; user-select:none; }
</style>
</head>
<body>
<header>
  <h1>🛰️ SheetCopilot Trajectories</h1>
  <span class="stat">Source: <b id="src"></b></span>
  <span class="stat"><b id="ntotal">0</b> trajectories</span>
  <span class="stat"><b id="npass" style="color:var(--green)">0</b> pass / <b id="nfail" style="color:var(--red)">0</b> fail</span>
  <span class="stat"><b id="ntok">0</b> tokens</span>
</header>
<div class="layout">
  <div class="sidebar">
    <div class="controls">
      <input id="search" placeholder="Search task / instruction...">
      <select id="filter">
        <option value="all">All trajectories</option>
        <option value="fail">Failed only</option>
        <option value="pass">Passed only</option>
      </select>
    </div>
    <div id="tasklist"></div>
  </div>
  <div class="main" id="detail"><div class="empty">Select a trajectory on the left.</div></div>
</div>
<script>
const DATA = __DATA__;
const SRC = __SRC__;

const STAGE_COLORS = {
  coarse_planning:'#5b9dff', fine_planning:'#a371f7', planning:'#5b9dff',
  extract_actions:'#39c5cf', execute:'#3fb950', chat:'#9aa3b2'
};
function stageColor(s){ return STAGE_COLORS[s] || '#d29922'; }
function esc(s){ return (s==null?'':String(s)); }

function callSuccess(c){
  if(c.error) return false;
  if(c.execution_success===false) return false;
  if(c.invalid_apis && c.invalid_apis.length) return false;
  return true;
}
// Prefer the evaluator's outcome verdict; fall back to whether the agent ran.
function trajPassed(t){ if(t.eval) return t.eval==='pass'; return !!(t.meta && t.meta.success); }
function trajLabel(t){ if(t.eval) return t.eval==='pass'?'PASS':'FAIL'; return (t.meta&&t.meta.success)?'RAN':'ERR'; }
function trajTokens(t){ return (t.meta && t.meta.usage_summary && t.meta.usage_summary.total_tokens) || 0; }

// ---- header stats ----
document.getElementById('src').textContent = SRC;
document.getElementById('ntotal').textContent = DATA.length;
document.getElementById('npass').textContent = DATA.filter(trajPassed).length;
document.getElementById('nfail').textContent = DATA.filter(t=>!trajPassed(t)).length;
document.getElementById('ntok').textContent = DATA.reduce((a,t)=>a+trajTokens(t),0).toLocaleString();

let activeIdx = -1;

function renderList(){
  const q = document.getElementById('search').value.toLowerCase();
  const f = document.getElementById('filter').value;
  const list = document.getElementById('tasklist');
  list.innerHTML='';
  DATA.forEach((t,i)=>{
    const pass = trajPassed(t);
    if(f==='pass'&&!pass) return;
    if(f==='fail'&&pass) return;
    const instr = esc(t.meta.instruction||'');
    if(q && !(t.name.toLowerCase().includes(q) || instr.toLowerCase().includes(q))) return;
    const div=document.createElement('div');
    div.className='task'+(i===activeIdx?' active':'');
    div.innerHTML = `<div style="flex:1;min-width:0">
        <div class="nm">${esc(t.name)}</div>
        <div class="sub">${t.calls.length} calls · ${trajTokens(t).toLocaleString()} tok</div>
      </div><span class="pill ${pass?'pass':'fail'}">${trajLabel(t)}</span>`;
    div.onclick=()=>{ activeIdx=i; renderList(); renderDetail(t); };
    list.appendChild(div);
  });
}

function messagesHTML(msgs){
  if(!msgs) return '';
  return msgs.map(m=>`<div class="msg ${esc(m.role)}"><div class="role">${esc(m.role)}</div><pre>${esc(m.content)}</pre></div>`).join('');
}

function callHTML(c){
  const ok = callSuccess(c);
  const u = c.usage||{};
  const acts = (c.executed_actions || c.parsed_actions || []);
  const actsHTML = acts.length ? `<div class="section"><div class="lbl">${c.executed_actions?'Executed actions':'Parsed actions'}</div>
      <div class="acts">${acts.map(a=>`<span class="act">${esc(a)}</span>`).join('')}</div></div>` : '';
  const invHTML = (c.invalid_apis&&c.invalid_apis.length)?`<div class="section"><div class="lbl">Invalid APIs</div>
      <div class="acts">${c.invalid_apis.map(a=>`<span class="act bad">${esc(a)}</span>`).join('')}</div></div>`:'';
  const reason = c.response && c.response.reasoning_content;
  return `<div class="call ${ok?'':'err'}">
    <div class="call-head" onclick="this.nextElementSibling.classList.toggle('open')">
      <span class="idx">#${c.id}</span>
      <span class="stage" style="background:${stageColor(c.stage)}22;color:${stageColor(c.stage)}">${esc(c.stage)}</span>
      <span class="spacer"></span>
      ${c.latency_s!=null?`<span class="chip">${c.latency_s}s</span>`:''}
      ${u.total_tokens!=null?`<span class="chip">${u.prompt_tokens||0}+${u.completion_tokens||0}=${u.total_tokens} tok</span>`:''}
      ${c.execution_success!=null?`<span class="chip ${c.execution_success?'ok':'bad'}">exec ${c.execution_success?'ok':'fail'}</span>`:''}
      ${c.error?`<span class="chip bad">error</span>`:''}
    </div>
    <div class="call-body">
      ${c.error?`<div class="errbox">${esc(c.error)}</div>`:''}
      <div class="section"><div class="lbl">Response</div><pre>${esc(c.response&&c.response.content)}</pre></div>
      ${reason?`<details><summary>Show reasoning (${reason.length} chars)</summary><pre>${esc(reason)}</pre></details>`:''}
      ${actsHTML}${invHTML}
      <details><summary>Show query (${(c.request_messages||[]).length} messages)</summary>${messagesHTML(c.request_messages)}</details>
    </div></div>`;
}

function renderDetail(t){
  const m=t.meta, us=m.usage_summary||{};
  const path = t.calls.map(c=>`<span class="node" style="background:${stageColor(c.stage)}22;color:${stageColor(c.stage)};border-color:${stageColor(c.stage)}66" title="call #${c.id}">${esc(c.stage)}${c.error?' ✕':''}</span>`).join('<span class="arrow">→</span>');
  const finalState = trajPassed(t) ? `<span class="node" style="background:rgba(63,185,80,.16);color:var(--green)">end ✓</span>` : `<span class="node" style="background:rgba(248,81,73,.16);color:var(--red)">fail ✕</span>`;
  document.getElementById('detail').innerHTML = `
    <div class="meta-card">
      <h2>${esc(t.name)} <span class="pill ${trajPassed(t)?'pass':'fail'}">${trajLabel(t)}</span></h2>
      <div class="meta-grid">
        <div>Model: <b>${esc(m.model)}</b></div>
        <div>Calls: <b>${t.calls.length}</b></div>
        <div>Tokens: <b>${(us.total_tokens||0).toLocaleString()}</b></div>
        <div>Wall time: <b>${us.wall_time_s!=null?us.wall_time_s+'s':'-'}</b></div>
        <div>Mode: <b>${esc(m.mode||'agent')}</b></div>
      </div>
      ${m.instruction?`<div class="instr">${esc(m.instruction)}</div>`:''}
      <div class="section"><div class="lbl">State-machine path</div><div class="path">${path}<span class="arrow">→</span>${finalState}</div></div>
      <div class="toggle" onclick="document.querySelectorAll('.call-body').forEach(e=>e.classList.add('open'))">⊕ expand all calls</div>
    </div>
    ${t.calls.map(callHTML).join('')}`;
}

document.getElementById('search').oninput=renderList;
document.getElementById('filter').onchange=renderList;
renderList();
if(DATA.length){ activeIdx=0; renderList(); renderDetail(DATA[0]); }
</script>
</body>
</html>
"""


def build_html(items, source_label):
    data_json = json.dumps(items, ensure_ascii=False)
    out = HTML_TEMPLATE.replace("__DATA__", data_json)
    out = out.replace("__SRC__", json.dumps(source_label, ensure_ascii=False))
    return out


def main():
    p = argparse.ArgumentParser(description="Visualize SheetCopilot trajectories as a self-contained HTML page.")
    p.add_argument("--results-dir", "-d", required=True,
                   help="directory containing *_trajectory.json files (searched recursively)")
    p.add_argument("--out", "-o", default=None, help="output HTML path (default: <results-dir>/trajectories.html)")
    args = p.parse_args()

    items = load_trajectories(args.results_dir)
    if not items:
        print(f"No *_trajectory.json found under {args.results_dir}")
        return 1

    out_path = args.out or os.path.join(args.results_dir, "trajectories.html")
    with open(out_path, "w", encoding="utf-8") as f:
        f.write(build_html(items, os.path.abspath(args.results_dir)))

    n_pass = sum(1 for t in items if (t["eval"] == "pass" if t["eval"] else t["meta"].get("success")))
    print(f"Wrote {out_path}")
    print(f"  {len(items)} trajectories  ({n_pass} pass / {len(items) - n_pass} fail)")
    print(f"  Open it in a browser: file://{os.path.abspath(out_path)}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
