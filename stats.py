"""
ChangeGear 派單準確率報告
==========================

從 changegear_history.db 計算 bot 派單的準確率，分多個維度呈現：

  1. 整體準確率（自啟用以來 + 過去 N 天）
  2. 分項準確率（Owner / Assigned / 三層 Inc Type / Req Item）
     僅對「擴充 schema 之後新派的單」有效，
     因為舊資料的 original_* 是空的。
  3. 決策來源分布（Claude / CMDB / DB / Excel / Default）
     + 各來源的成功率
  4. 信心分數 vs 實際準確率（驗證 Claude confidence 是否可信）
  5. 最常被修正的「人 / Inc Type / 寄件者」
  6. 學習資料統計（feedback + correction + 原則檔）

執行方式：
    python stats.py              # 預設過去 7 天
    python stats.py --days 30    # 過去 30 天
    python stats.py --all        # 自啟用以來全部
    python stats.py --json       # JSON 格式輸出（便於串其他工具）
"""

import sqlite3
import sys
import os
import json
from datetime import datetime, timedelta
from collections import Counter

# 強制 stdout 用 UTF-8（cp950 cmd 視窗無法顯示 unicode block 字元）
try:
    sys.stdout.reconfigure(encoding="utf-8")
except (AttributeError, Exception):
    pass

DB_PATH = "changegear_history.db"
PRINCIPLES_FILE = "learned_principles.md"


# ── DB query ─────────────────────────────────────────────────
def fetch_dispatches(conn: sqlite3.Connection, cutoff: str | None) -> list[dict]:
    """抓取 bot 派單記錄。cutoff=None 時抓全部。"""
    where = "WHERE bot_assigned = 1"
    params = ()
    if cutoff:
        where += " AND created_at >= ?"
        params = (cutoff,)
    rows = conn.execute(f"""
        SELECT ticket_id, requester,
               owner, assigned_to,
               inc_parent, inc_child, inc_item, req_item,
               original_owner, original_assigned_to,
               original_inc_parent, original_inc_child,
               original_inc_item, original_req_item,
               decision_source, confidence,
               corrected, created_at
        FROM assignments
        {where}
        ORDER BY created_at DESC
    """, params).fetchall()
    cols = ["ticket_id", "requester",
            "owner", "assigned_to",
            "inc_parent", "inc_child", "inc_item", "req_item",
            "original_owner", "original_assigned_to",
            "original_inc_parent", "original_inc_child",
            "original_inc_item", "original_req_item",
            "decision_source", "confidence",
            "corrected", "created_at"]
    return [dict(zip(cols, r)) for r in rows]


# ── Compute ──────────────────────────────────────────────────
def pct(n: int, d: int) -> str:
    return f"{(n/d*100):.1f}%" if d > 0 else "n/a"


def overall_accuracy(rows: list[dict]) -> dict:
    total = len(rows)
    corrected = sum(1 for r in rows if r["corrected"] == 1)
    correct = total - corrected
    return {
        "total": total,
        "correct": correct,
        "corrected": corrected,
        "accuracy_pct": (correct / total * 100) if total > 0 else 0,
    }


def field_accuracy(rows: list[dict]) -> dict:
    """分項準確率 — 比 original_* 與 current（被修正後）的差異。
    只算 original_* 非空且 corrected=1 的記錄，加上所有 corrected=0 的記錄。
    """
    fields = [
        ("owner",         "Owner"),
        ("assigned_to",   "Assigned To"),
        ("inc_parent",    "Inc Parent"),
        ("inc_child",     "Inc Child"),
        ("inc_item",      "Inc Item"),
        ("req_item",      "Req Item"),
    ]
    result = {}
    for col, label in fields:
        total = 0
        same  = 0
        for r in rows:
            orig = (r.get(f"original_{col}") or "").strip()
            cur  = (r.get(col) or "").strip()
            if r["corrected"] == 0:
                # 沒被修正 → bot 正確
                total += 1
                same  += 1
            elif orig:
                # 被修正 + 有快照 → 比對
                total += 1
                if orig.lower() == cur.lower():
                    same += 1
        result[label] = {
            "total":   total,
            "correct": same,
            "pct":     (same / total * 100) if total > 0 else None,
        }
    return result


def source_breakdown(rows: list[dict]) -> dict:
    """各決策來源的派單數與準確率"""
    by_source: dict[str, dict] = {}
    for r in rows:
        src = r.get("decision_source") or "unknown"
        d = by_source.setdefault(src, {"total": 0, "corrected": 0, "conf_sum": 0.0})
        d["total"] += 1
        d["conf_sum"] += float(r.get("confidence") or 0)
        if r["corrected"] == 1:
            d["corrected"] += 1
    for src, d in by_source.items():
        d["correct"] = d["total"] - d["corrected"]
        d["acc_pct"] = (d["correct"] / d["total"] * 100) if d["total"] else 0
        d["avg_conf"] = (d["conf_sum"] / d["total"]) if d["total"] else 0
        del d["conf_sum"]
    return dict(sorted(by_source.items(), key=lambda x: -x[1]["total"]))


def confidence_vs_accuracy(rows: list[dict]) -> list[dict]:
    """信心分數分箱統計"""
    buckets = [
        ("0.90 – 1.00", 0.90, 1.01),
        ("0.85 – 0.90", 0.85, 0.90),
        ("0.75 – 0.85", 0.75, 0.85),
        ("0.65 – 0.75", 0.65, 0.75),
        ("< 0.65",      0.00, 0.65),
    ]
    out = []
    for label, lo, hi in buckets:
        bucket_rows = [r for r in rows
                       if lo <= float(r.get("confidence") or 0) < hi]
        total = len(bucket_rows)
        corrected = sum(1 for r in bucket_rows if r["corrected"] == 1)
        out.append({
            "bucket":  label,
            "total":   total,
            "correct": total - corrected,
            "acc_pct": ((total - corrected) / total * 100) if total else None,
        })
    return out


def top_corrected(rows: list[dict], by: str, top_n: int = 5) -> list[dict]:
    """找出被修正最多的「人 / 寄件者 / req_item / inc_type 組合」"""
    counter = Counter()
    for r in rows:
        if r["corrected"] != 1:
            continue
        key = r.get(by) or "(無)"
        if not key.strip():
            key = "(無)"
        counter[key] += 1
    return [{"key": k, "corrections": v} for k, v in counter.most_common(top_n)]


def weekly_trend(rows: list[dict]) -> list[dict]:
    """過去 4 週滑動準確率"""
    now = datetime.now()
    out = []
    for offset in range(3, -1, -1):
        end = now - timedelta(days=7 * offset)
        start = end - timedelta(days=7)
        bucket = [
            r for r in rows
            if start.strftime("%Y-%m-%d") <= r["created_at"][:10] < end.strftime("%Y-%m-%d")
        ]
        total = len(bucket)
        correct = total - sum(1 for r in bucket if r["corrected"] == 1)
        out.append({
            "week":      f"{start.strftime('%m/%d')} – {end.strftime('%m/%d')}",
            "total":     total,
            "correct":   correct,
            "acc_pct":   (correct / total * 100) if total else None,
        })
    return out


def learning_resources_summary(conn: sqlite3.Connection) -> dict:
    has_feedback = conn.execute(
        "SELECT name FROM sqlite_master WHERE name='feedback'"
    ).fetchone()
    feedback_n = 0
    if has_feedback:
        feedback_n = conn.execute("SELECT COUNT(*) FROM feedback").fetchone()[0]

    corrected_n = conn.execute(
        "SELECT COUNT(*) FROM assignments WHERE corrected = 1"
    ).fetchone()[0]

    tracked_n = conn.execute(
        "SELECT COUNT(*) FROM assignments WHERE bot_assigned = 2"
    ).fetchone()[0]

    has_principles = os.path.exists(PRINCIPLES_FILE)
    principle_size = 0
    if has_principles:
        with open(PRINCIPLES_FILE, encoding="utf-8") as f:
            principle_size = sum(1 for line in f if line.startswith("## 原則"))

    return {
        "feedback_count":           feedback_n,
        "implicit_corrected_count": corrected_n,
        "tracked_only_count":       tracked_n,
        "principles_file_exists":   has_principles,
        "principles_count":         principle_size,
    }


# ── Render text report ──────────────────────────────────────
def bar(value: float, width: int = 30) -> str:
    """把百分比畫成 ASCII bar"""
    if value is None:
        return "n/a"
    filled = int(round(value / 100 * width))
    return "█" * filled + "░" * (width - filled)


def render_text(report: dict) -> str:
    out = []
    out.append("=" * 70)
    out.append(f"  ChangeGear 派單準確率報告 — {report['period']}")
    out.append("=" * 70)
    out.append("")

    # 整體
    o = report["overall"]
    out.append("── 整體 ──")
    out.append(f"  派單總數:        {o['total']}")
    out.append(f"  未被修正（正確）: {o['correct']}")
    out.append(f"  被修正:          {o['corrected']}")
    out.append(f"  整體準確率:      {o['accuracy_pct']:.1f}%   {bar(o['accuracy_pct'])}")
    out.append("")

    # 分項
    out.append("── 分項準確率（僅算有 original_* 快照的記錄）──")
    fa = report["field_accuracy"]
    for label, d in fa.items():
        pct_str = f"{d['pct']:.1f}%" if d["pct"] is not None else "n/a"
        out.append(f"  {label:14s}: {d['correct']:3d} / {d['total']:3d}   {pct_str:>6s}   {bar(d['pct'] or 0)}")
    out.append("")

    # 來源
    out.append("── 決策來源分布 ──")
    sb = report["source_breakdown"]
    if not sb:
        out.append("  （尚無資料）")
    else:
        for src, d in sb.items():
            out.append(
                f"  {src:10s}: {d['total']:3d} 筆 ({pct(d['total'], o['total']):>5s})   "
                f"準確 {d['acc_pct']:5.1f}%   avg信心 {d['avg_conf']:.2f}"
            )
    out.append("")

    # 信心 vs 準確
    out.append("── 信心分數 vs 實際準確率 ──")
    for b in report["confidence_buckets"]:
        if b["total"] == 0:
            continue
        pct_str = f"{b['acc_pct']:.1f}%" if b["acc_pct"] is not None else "n/a"
        out.append(f"  信心 {b['bucket']:14s}: {b['correct']:3d} / {b['total']:3d}   {pct_str:>6s}")
    out.append("")

    # 趨勢
    out.append("── 過去 4 週滑動準確率 ──")
    for w in report["weekly_trend"]:
        pct_str = f"{w['acc_pct']:.1f}%" if w["acc_pct"] is not None else "n/a"
        if w["total"] > 0:
            out.append(f"  {w['week']}: {w['correct']:3d} / {w['total']:3d}   {pct_str:>6s}   {bar(w['acc_pct'] or 0)}")
        else:
            out.append(f"  {w['week']}: (無資料)")
    out.append("")

    # Top 被修正
    out.append("── 最常被修正的 Assigned To（bot 派錯給誰）──")
    for item in report["top_corrected_assigned"]:
        out.append(f"  {item['key']:30s}  ×{item['corrections']}")
    out.append("")

    out.append("── 最常被修正的 Inc Child ──")
    for item in report["top_corrected_inc_child"]:
        out.append(f"  {item['key']:30s}  ×{item['corrections']}")
    out.append("")

    out.append("── 最常被修正的 Requester（誰開的單最常派錯）──")
    for item in report["top_corrected_requester"]:
        out.append(f"  {item['key']:30s}  ×{item['corrections']}")
    out.append("")

    # 學習資料
    l = report["learning"]
    out.append("── 學習資料概況 ──")
    out.append(f"  feedback 教學紀錄:        {l['feedback_count']}")
    out.append(f"  隱式修正（corrected=1）:  {l['implicit_corrected_count']}")
    out.append(f"  追蹤中（未派）工單:       {l['tracked_only_count']}")
    if l["principles_file_exists"]:
        out.append(f"  learned_principles.md:    存在（{l['principles_count']} 條原則）")
    else:
        out.append(f"  learned_principles.md:    (尚未產生)")
    out.append("")

    # 提示
    out.append("─" * 70)
    if any(d["pct"] is None for d in fa.values()):
        out.append("⚠ 分項準確率有 n/a 的欄位 → 表示舊資料缺 original_* 快照。")
        out.append("  以後新派的單會有完整資料。")
    out.append("=" * 70)
    return "\n".join(out)


# ── Main ────────────────────────────────────────────────────
def main():
    args = sys.argv[1:]
    days = 7
    show_all = False
    as_json = False
    while args:
        a = args.pop(0)
        if a == "--days" and args:
            days = int(args.pop(0))
        elif a == "--all":
            show_all = True
        elif a == "--json":
            as_json = True
        elif a in ("--help", "-h"):
            print(__doc__)
            return

    if not os.path.exists(DB_PATH):
        print(f"⛔ 找不到 {DB_PATH}")
        sys.exit(1)

    cutoff = None
    if not show_all:
        cutoff = (datetime.now() - timedelta(days=days)).strftime("%Y-%m-%d %H:%M:%S")

    conn = sqlite3.connect(DB_PATH)
    try:
        rows = fetch_dispatches(conn, cutoff)
        period = "自啟用以來" if show_all else f"過去 {days} 天"

        report = {
            "period":                  period,
            "overall":                 overall_accuracy(rows),
            "field_accuracy":          field_accuracy(rows),
            "source_breakdown":        source_breakdown(rows),
            "confidence_buckets":      confidence_vs_accuracy(rows),
            "weekly_trend":            weekly_trend(rows),
            "top_corrected_assigned":  top_corrected(rows, "assigned_to"),
            "top_corrected_inc_child": top_corrected(rows, "inc_child"),
            "top_corrected_requester": top_corrected(rows, "requester"),
            "learning":                learning_resources_summary(conn),
        }
    finally:
        conn.close()

    if as_json:
        print(json.dumps(report, indent=2, ensure_ascii=False))
    else:
        print(render_text(report))


if __name__ == "__main__":
    try:
        main()
    except KeyboardInterrupt:
        print("\n已中止")
        sys.exit(0)
