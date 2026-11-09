"""Excel「必勝カウントダウン」シートから閲覧用 index.html を生成する。
使い方: python3 build.py [xlsxパス]
公開するのはカウントダウン用の数値・予定・提出状況のみ（他シートは出力しない）。
"""
import sys, json, datetime as dt, pathlib, warnings
import openpyxl

warnings.filterwarnings("ignore")
DEFAULT = pathlib.Path.home() / "Downloads" / "ITストラテジスト_必勝合格スケジュール_11月締め_独立移行版_自動集計修正版 (1).xlsx"
src = pathlib.Path(sys.argv[1]) if len(sys.argv) > 1 else DEFAULT
here = pathlib.Path(__file__).parent

ws = openpyxl.load_workbook(src)["必勝カウントダウン"]
COLS = range(11, 26)  # K..Y

def num(v):
    return v if isinstance(v, (int, float)) and not isinstance(v, bool) else None

def iso(v):
    return v.strftime("%Y-%m-%d") if isinstance(v, (dt.datetime, dt.date)) else None

params = {k: ws[f"C{r}"].value for k, r in
          [("cap", 64), ("floor", 65), ("th1", 66), ("th2", 67), ("th3", 68), ("ex", 69), ("set", 70), ("lap", 71)]}
exam = iso(ws["C61"].value)

KEYS = ["jk", "gt", "mk", "hb", "aw", "kk", "fl", "sb", "ot"]
tasks = []
for i, k in enumerate(KEYS):
    pr, ar = 31 + 2 * i, 32 + 2 * i
    plan = [num(ws.cell(pr, c).value) or 0 for c in COLS]
    act = [num(ws.cell(ar, c).value) for c in COLS] if k != "sb" else None
    mn, tg = ws[f"E{ar}"].value, ws[f"F{ar}"].value
    if k == "gt" and not num(mn): mn = params["ex"] * params["lap"]
    if k == "ot": mn = tg = sum(plan)
    tasks.append({"key": k, "label": ws[f"B{pr}"].value, "unit": ws[f"C{pr}"].value, "std": num(ws[f"D{pr}"].value) or 0,
                  "min": num(mn), "target": num(tg), "plan": plan, "act": act})

subs = []
for r in range(54, 59):
    subs.append({"name": ws[f"B{r}"].value, "due": iso(ws[f"C{r}"].value), "status": ws[f"E{r}"].value or "",
                 "submitted": iso(ws[f"G{r}"].value), "returned": iso(ws[f"I{r}"].value), "memo": ws[f"L{r}"].value or ""})

data = {
    "generated": dt.datetime.now().strftime("%Y-%m-%d %H:%M"),
    "exam": exam, "params": params,
    "dates": [iso(ws.cell(13, c).value) for c in COLS],
    "flow": [ws.cell(16, c).value or "" for c in COLS],
    "actMin": [num(ws.cell(19, c).value) for c in COLS],
    "memo": [ws.cell(50, c).value or "" for c in COLS],
    "tasks": tasks, "subs": subs,
}
tpl = (here / "template.html").read_text(encoding="utf-8")
(here / "index.html").write_text(tpl.replace("/*__DATA__*/null", json.dumps(data, ensure_ascii=False)), encoding="utf-8")
print("index.html generated from", src.name, "at", data["generated"])
