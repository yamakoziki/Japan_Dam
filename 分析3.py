"""
全国ダム地質DB 地質推定可能性分析 v3
=====================================
北海道開発局ダム（管理者=国交省かつ北海道）のsymbolをリファレンスとして、
他グループのダムの各symbolが推定可能かどうかを評価し、調査優先ダムを提案する。

C1_北海道開発局リファレンス  : symbolリファレンスセット（管理者=国交省かつ北海道のダム一覧）
C2_グループ1推定可能性       : 北海道内・開発局外ダムの各symbolがリファレンスに存在するか評価
C3_グループ2推定可能性       : 北海道以外ダムの各symbolがリファレンスに存在するか評価
D1_調査優先ダムリスト        : 未情報symbolを持つダムの優先調査提案（展開価値ベース）
D2_未掲載Symbol調査効果      : 未掲載symbol別・解決がもたらす推定可能化ダム数の分析

北海道開発局ダム定義  : 管理者コード=1（国交省）かつ所在地が北海道
ダムグループ１定義    : 北海道内にある北海道開発局以外のダム
ダムグループ２定義    : 北海道以外のダム
推定可能判定          : symbolが北海道開発局リファレンスセットに含まれるか（層位置・セット無関係）
展開価値              : ある未知symbolが持つ他ダムへの波及効果（同symbolを持つ他ターゲットダム数の合計）
"""

import argparse
import datetime
import sys
from collections import Counter, defaultdict
from pathlib import Path
from statistics import mean

from openpyxl import load_workbook, Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

# ─── スタイル ─────────────────────────────────────────────────
HDR_FILL  = PatternFill("solid", fgColor="1F4E79")
HDR_FONT  = Font(name="Arial", bold=True, color="FFFFFF", size=10)
SUB_FILL  = PatternFill("solid", fgColor="2E75B6")
SUB_FONT  = Font(name="Arial", bold=True, color="FFFFFF", size=10)
SEC_FILL  = PatternFill("solid", fgColor="D6E4F0")
BODY_FONT = Font(name="Arial", size=10)
BOLD_FONT = Font(name="Arial", bold=True, size=10)
ALT_FILL  = PatternFill("solid", fgColor="F2F7FC")
RED_FILL  = PatternFill("solid", fgColor="FFE0E0")
YEL_FILL  = PatternFill("solid", fgColor="FFF8DC")
GRN_FILL  = PatternFill("solid", fgColor="E8F5E9")
CTR  = Alignment(horizontal="center", vertical="center", wrap_text=True)
LEFT = Alignment(horizontal="left",   vertical="center", wrap_text=True)
RIGHT = Alignment(horizontal="right",  vertical="center")
THIN = Side(style="thin", color="BFBFBF")
BDR  = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)

# ─── 定数 ─────────────────────────────────────────────────────
LAYER_ID_COL = {1: 23, 2: 35, 3: 47, 4: 59, 5: 71}
LAYER_NAMES = {
    1: "Layer1: Pre-N（新第三紀以前）",
    2: "Layer2: N-古（新第三紀・古い方）",
    3: "Layer3: N-新（新第三紀・新しい方）",
    4: "Layer4: Q-old（中期更新世）",
    5: "Layer5: Q-H（後期更新世〜完新世）",
}

BEARING_SCORE = {
    "低": 1, "低〜中": 2, "中": 3,
    "中〜高": 4, "中〜高（溶結度依存）": 4,
    "高（続成固結）": 5, "高": 5,
}
PERM_SCORE = {
    "低〜中": 1, "中（節理閉鎖）": 2, "中": 2, "中〜高": 3,
    "高": 4, "高（柱状節理）": 4, "高（柱状節理・カルデラ）": 4,
}

def b_score(v): return BEARING_SCORE.get(v, 0)
def p_score(v): return PERM_SCORE.get(v, 0)


# ─── スタイルヘルパー ─────────────────────────────────────────
def hdr(ws, r, c, v, w=None):
    cell = ws.cell(r, c, v)
    cell.font = HDR_FONT; cell.fill = HDR_FILL
    cell.alignment = CTR; cell.border = BDR
    if w: ws.column_dimensions[get_column_letter(c)].width = w
    return cell

def sub(ws, r, c, v, w=None):
    cell = ws.cell(r, c, v)
    cell.font = SUB_FONT; cell.fill = SUB_FILL
    cell.alignment = CTR; cell.border = BDR
    if w: ws.column_dimensions[get_column_letter(c)].width = w
    return cell

def body(ws, r, c, v, align=LEFT, bold=False, fill=None):
    cell = ws.cell(r, c, v)
    cell.font = BOLD_FONT if bold else BODY_FONT
    cell.alignment = align; cell.border = BDR
    if fill: cell.fill = fill
    return cell

def sec_title(ws, r, c1, c2, text):
    ws.merge_cells(f"{get_column_letter(c1)}{r}:{get_column_letter(c2)}{r}")
    cell = ws.cell(r, c1, text)
    cell.font = BOLD_FONT; cell.fill = SEC_FILL; cell.alignment = LEFT
    return cell

def sheet_title(ws, r, c1, c2, text):
    ws.merge_cells(f"{get_column_letter(c1)}{r}:{get_column_letter(c2)}{r}")
    cell = ws.cell(r, c1, text)
    cell.font = Font(name="Arial", bold=True, size=12, color="1F4E79")
    cell.alignment = LEFT
    return cell

def pct(n, total): return round(n / total * 100, 1) if total else 0
def avg(lst): return round(mean(lst), 2) if lst else ""
def scol(c): return get_column_letter(c)


# ─── データ読み込み ───────────────────────────────────────────
def load_data(wb):
    ws_db = wb["全国ダム地質DB"]
    ws_gl = wb["Glossary"]

    glossary = {}
    fields = ["id", "symbol", "geo_surface", "geo_era", "geo_rock",
              "formationAge_ja", "group_ja", "lithology_ja", "geo_rock_label",
              "bearing_cap", "permeability", "main_risk"]
    for row in ws_gl.iter_rows(min_row=4, values_only=True):
        if row[0] is None: continue
        try:
            rec = {f: row[i] for i, f in enumerate(fields)}
            glossary[int(rec["id"])] = rec
        except:
            continue

    dams = []
    for row in range(3, ws_db.max_row + 1):
        name = ws_db.cell(row, 3).value
        if not name: continue

        layers = {}
        for lnum, col in LAYER_ID_COL.items():
            gid = ws_db.cell(row, col).value
            if gid is not None and not isinstance(gid, str):
                try:
                    layers[lnum] = glossary.get(int(gid))
                except:
                    layers[lnum] = None
            else:
                layers[lnum] = None

        filled = [k for k, v in layers.items() if v is not None]
        recs = [v for v in layers.values() if v is not None]

        loc = ws_db.cell(row, 17).value or ""
        mgr = ws_db.cell(row, 14).value
        height = ws_db.cell(row, 10).value

        dams.append({
            "name": name,
            "loc": loc,
            "pref": loc[:3] if loc else "",
            "mgr_code": mgr,
            "height": height if isinstance(height, (int, float)) else None,
            "layers": layers,
            "recs": recs,
            "filled_layers": filled,
            "layer_count": len(filled),
        })
    return dams, glossary


# ─── C系: 北海道開発局リファレンス推定可能性分析 ────────────────

def build_ref_set(hokdev_dams):
    """北海道開発局ダムのsymbol全体セット（層位置関係なし）を構築。"""
    ref_set    = set()
    ref_detail = defaultdict(list)
    for d in hokdev_dams:
        for lnum in range(1, 6):
            rec = d["layers"][lnum]
            if rec and rec.get("symbol"):
                sym = rec["symbol"]
                ref_set.add(sym)
                ref_detail[sym].append((d["name"], lnum))
    return ref_set, ref_detail


def _ref_label(n_ok, n_total):
    if n_total == 0:    return "データなし"
    if n_ok == n_total: return "全層推定可能"
    if n_ok == 0:       return "全層推定不可"
    return "一部推定不可"


def _ref_fill_row(n_ok, n_total):
    if n_total == 0:    return None
    if n_ok == n_total: return GRN_FILL
    if n_ok == 0:       return RED_FILL
    return YEL_FILL


def _estimability_fill(layer_count):
    if layer_count == 0:      return RED_FILL
    if layer_count <= 2:      return YEL_FILL
    return GRN_FILL

def _estimability_label(layer_count):
    if layer_count == 0:  return "推定不可"
    if layer_count <= 2:  return "推定限定的"
    return "推定良好"


# ─── C1: 北海道開発局リファレンス ────────────────────────────
def write_c1(wb, hokdev_dams, glossary, ref_set, ref_detail):
    ws = wb.create_sheet("C1_北海道開発局リファレンス")
    ws.sheet_view.showGridLines = False
    sheet_title(ws, 1, 1, 12,
        "■ C1 北海道開発局リファレンス — symbolリファレンスセット（管理者=国交省かつ北海道）")
    row = 2

    total = len(hokdev_dams)

    # ── サマリー ──
    sec_title(ws, row, 1, 3, "▼ 北海道開発局ダム サマリー")
    row += 1
    ws.column_dimensions["A"].width = 34
    ws.column_dimensions["B"].width = 12
    for label, val in [
        ("北海道開発局ダム数（管理者=国交省かつ北海道）", total),
        ("リファレンスsymbol数（ユニーク）",              len(ref_set)),
        ("symbol延べ出現数",                             sum(len(v) for v in ref_detail.values())),
        ("データあり（1層以上）",                        sum(1 for d in hokdev_dams if d["layer_count"] > 0)),
        ("データなし（0層）",                            sum(1 for d in hokdev_dams if d["layer_count"] == 0)),
        ("平均層数",                                     avg([d["layer_count"] for d in hokdev_dams])),
    ]:
        body(ws, row, 1, label, bold=True)
        body(ws, row, 2, val, align=RIGHT)
        row += 1
    row += 1

    # ── ダム一覧 ──
    sec_title(ws, row, 1, 12, f"▼ 北海道開発局ダム一覧（全{total}件）")
    row += 1
    hdrs   = ["ダム名", "所在地", "堤高(m)", "層数",
              "L1-symbol", "L2-symbol", "L3-symbol", "L4-symbol", "L5-symbol",
              "データ充実度", "平均強度", "平均透水性"]
    widths = [16, 20, 8, 6, 16, 16, 16, 16, 16, 14, 10, 10]
    hdr_row = row
    for ci, (h, w) in enumerate(zip(hdrs, widths), 1):
        hdr(ws, row, ci, h, w)
    row += 1

    for d in sorted(hokdev_dams, key=lambda d: (-d["layer_count"], -(d["height"] or 0))):
        fill  = _estimability_fill(d["layer_count"])
        b_avg = avg([b_score(r["bearing_cap"])  for r in d["recs"] if b_score(r["bearing_cap"])])
        p_avg = avg([p_score(r["permeability"]) for r in d["recs"] if p_score(r["permeability"])])
        syms  = [d["layers"][n]["symbol"] if d["layers"][n] and d["layers"][n].get("symbol") else ""
                 for n in range(1, 6)]
        vals  = [d["name"], d["loc"][:20] if d["loc"] else "", d["height"], d["layer_count"]] + \
                syms + [_estimability_label(d["layer_count"]), b_avg, p_avg]
        for ci, v in enumerate(vals, 1):
            body(ws, row, ci, v, align=RIGHT if ci in (3, 4, 11, 12) else LEFT, fill=fill)
        row += 1
    row += 1

    # ── Symbolリファレンス一覧 ──
    sec_title(ws, row, 1, 5, f"▼ Symbolリファレンス一覧（{len(ref_set)}種、出現頻度降順）")
    row += 1
    for ci, (t, w) in enumerate(zip(
        ["symbol", "延べ出現数", "出現ダム数", "出現情報（ダム名:層番号）", "地質情報（era / rock / lithology）"],
        [20, 12, 12, 60, 50]
    ), 1):
        hdr(ws, row, ci, t, w)
    row += 1

    sym_to_glo = {rec["symbol"]: rec for rec in glossary.values() if rec.get("symbol")}
    for sym, occurrences in sorted(ref_detail.items(), key=lambda x: -len(x[1])):
        dam_cnt = len(set(name for name, _ in occurrences))
        occ_str = "  ".join(f"{name}:L{lnum}" for name, lnum in occurrences[:8])
        glo     = sym_to_glo.get(sym)
        glo_str = ""
        if glo:
            parts = [str(glo.get("geo_era") or ""), str(glo.get("geo_rock") or ""),
                     str(glo.get("lithology_ja") or "")]
            glo_str = "  /  ".join(p for p in parts if p)
        for ci, v in enumerate([sym, len(occurrences), dam_cnt, occ_str, glo_str], 1):
            body(ws, row, ci, v, align=RIGHT if ci in (2, 3) else LEFT)
        row += 1

    ws.freeze_panes = "A3"
    ws.auto_filter.ref = f"A{hdr_row}:{scol(len(hdrs))}{hdr_row}"


# ─── C2/C3 共通: リファレンスベース推定可能性シート ──────────
def _write_ref_group(wb, sheet_name, title, target_dams, ref_set, group_label):
    ws = wb.create_sheet(sheet_name)
    ws.sheet_view.showGridLines = False
    sheet_title(ws, 1, 1, 12, title)
    row = 2

    total = len(target_dams)

    # per-dam 推定可能性を事前計算
    dam_stats = []
    for d in target_dams:
        layer_info = {}
        n_ok = n_total = 0
        for lnum in range(1, 6):
            rec = d["layers"][lnum]
            if rec is None or not rec.get("symbol"):
                layer_info[lnum] = (None, None)
            else:
                sym = rec["symbol"]
                est = sym in ref_set
                layer_info[lnum] = (sym, est)
                n_total += 1
                if est: n_ok += 1
        dam_stats.append({**d, "_li": layer_info, "_n_ok": n_ok, "_n_total": n_total})

    n_all_ok  = sum(1 for d in dam_stats if d["_n_total"] > 0 and d["_n_ok"] == d["_n_total"])
    n_partial = sum(1 for d in dam_stats if 0 < d["_n_ok"] < d["_n_total"])
    n_all_ng  = sum(1 for d in dam_stats if d["_n_total"] > 0 and d["_n_ok"] == 0)
    n_nodata  = sum(1 for d in dam_stats if d["_n_total"] == 0)

    # ── サマリー ──
    sec_title(ws, row, 1, 4, "▼ 推定可能性サマリー（北海道開発局symbolをリファレンスとした評価）")
    row += 1
    ws.column_dimensions["A"].width = 32
    ws.column_dimensions["B"].width = 10
    ws.column_dimensions["C"].width = 10
    for label, val, pct_v, sfill in [
        ("総ダム数",                    total,     "",                           None),
        ("全層推定可能",                n_all_ok,  f"{pct(n_all_ok,  total)}%",  GRN_FILL),
        ("一部推定不可",                n_partial, f"{pct(n_partial, total)}%",  YEL_FILL),
        ("全層推定不可（データあり）",  n_all_ng,  f"{pct(n_all_ng,  total)}%",  RED_FILL),
        ("データなし（0層）",           n_nodata,  f"{pct(n_nodata,  total)}%",  None),
    ]:
        body(ws, row, 1, label, bold=True, fill=sfill)
        body(ws, row, 2, val,   align=RIGHT, fill=sfill)
        body(ws, row, 3, pct_v, align=RIGHT, fill=sfill)
        row += 1
    row += 1

    # ── 層別マッチ率 ──
    sec_title(ws, row, 1, 6, "▼ 層別リファレンスマッチ率（開発局symbolに対応する割合）")
    row += 1
    for ci, (t, w) in enumerate(zip(
        ["層番号", "層名", "データあり", "マッチ(推定可)", "マッチ率(%)", "非マッチ(推定不可)"],
        [8, 30, 10, 14, 12, 16]
    ), 1):
        hdr(ws, row, ci, t, w)
    row += 1

    for lnum in range(1, 6):
        with_sym  = [d for d in dam_stats if d["_li"][lnum][0] is not None]
        match_cnt = sum(1 for d in with_sym if d["_li"][lnum][1])
        no_match  = len(with_sym) - match_cnt
        mr = pct(match_cnt, len(with_sym))
        mfill = GRN_FILL if mr >= 70 else YEL_FILL if mr >= 40 else (RED_FILL if with_sym else None)
        for ci, v in enumerate(
            [lnum, LAYER_NAMES[lnum], len(with_sym), match_cnt, mr, no_match], 1
        ):
            body(ws, row, ci, v,
                 align=RIGHT if ci in (1, 3, 4, 5, 6) else LEFT,
                 fill=mfill if ci in (4, 5) else None)
        row += 1
    row += 1

    # ── ダム別詳細 ──
    sec_title(ws, row, 1, 12,
        f"▼ ダム別詳細（緑=推定可 赤=推定不可 無色=データなし、全{total}件）")
    row += 1
    hdrs_c   = ["ダム名", "所在地", "堤高(m)", "管理者", "層数",
                "L1-symbol", "L2-symbol", "L3-symbol", "L4-symbol", "L5-symbol",
                "推定可能層数", "全体評価"]
    widths_c = [16, 20, 8, 8, 6, 16, 16, 16, 16, 16, 12, 14]
    hdr_row  = row
    for ci, (h, w) in enumerate(zip(hdrs_c, widths_c), 1):
        hdr(ws, row, ci, h, w)
    row += 1

    for d in sorted(dam_stats, key=lambda x: (-x["_n_ok"], -(x["height"] or 0))):
        rfill = _ref_fill_row(d["_n_ok"], d["_n_total"])
        for ci, v in enumerate(
            [d["name"], d["loc"][:20] if d["loc"] else "",
             d["height"], d["mgr_code"], d["layer_count"]], 1
        ):
            body(ws, row, ci, v, align=RIGHT if ci in (3, 4, 5) else LEFT, fill=rfill)
        for li, lnum in enumerate(range(1, 6), 6):
            sym, est = d["_li"][lnum]
            cfill = GRN_FILL if est else (RED_FILL if sym else None)
            body(ws, row, li, sym or "", fill=cfill)
        body(ws, row, 11, d["_n_ok"],                             align=RIGHT, fill=rfill)
        body(ws, row, 12, _ref_label(d["_n_ok"], d["_n_total"]),               fill=rfill)
        row += 1

    ws.freeze_panes = "A3"
    ws.auto_filter.ref = f"A{hdr_row}:{scol(len(hdrs_c))}{hdr_row}"
    row += 1

    # ── リファレンス未掲載symbol一覧 ──
    non_match_cnt  = Counter()
    non_match_pref = defaultdict(set)
    non_match_name = defaultdict(list)
    for d in dam_stats:
        for lnum in range(1, 6):
            sym, est = d["_li"][lnum]
            if sym and not est:
                non_match_cnt[sym] += 1
                non_match_pref[sym].add(d["pref"])
                non_match_name[sym].append(d["name"])

    if non_match_cnt:
        sec_title(ws, row, 1, 5,
            f"▼ リファレンス未掲載symbol一覧（{group_label}固有・全{len(non_match_cnt)}種、出現頻度降順）")
        row += 1
        for ci, (t, w) in enumerate(zip(
            ["symbol", "出現ダム数", "出現延べ数", "所在都道府県（上位3）", "代表ダム名（上位3）"],
            [20, 12, 12, 30, 40]
        ), 1):
            hdr(ws, row, ci, t, w)
        row += 1
        for sym, cnt in non_match_cnt.most_common(60):
            dam_cnt = len(set(non_match_name[sym]))
            prefs   = "、".join(list(non_match_pref[sym])[:3])
            names   = "、".join(non_match_name[sym][:3])
            for ci, v in enumerate([sym, dam_cnt, cnt, prefs, names], 1):
                body(ws, row, ci, v, align=RIGHT if ci in (2, 3) else LEFT, fill=RED_FILL)
            row += 1


def write_c2(wb, dams, ref_set):
    group1 = [d for d in dams
              if d["pref"].startswith("北海道") and d["mgr_code"] != 1]
    _write_ref_group(
        wb,
        "C2_グループ1推定可能性",
        "■ C2 ダムグループ１ リファレンス推定可能性（北海道内・北海道開発局以外）",
        group1, ref_set,
        "北海道内・開発局外",
    )


def write_c3(wb, dams, ref_set):
    group2 = [d for d in dams if not d["pref"].startswith("北海道")]
    _write_ref_group(
        wb,
        "C3_グループ2推定可能性",
        "■ C3 ダムグループ２ リファレンス推定可能性（北海道以外）",
        group2, ref_set,
        "北海道以外",
    )


# ─── D1: 調査優先ダムリスト ──────────────────────────────────
def write_d1(wb, dams, ref_set):
    ws = wb.create_sheet("D1_調査優先ダムリスト")
    ws.sheet_view.showGridLines = False
    sheet_title(ws, 1, 1, 13,
        "■ D1 調査優先ダムリスト — 未情報symbolを持つダムの優先調査提案（展開価値ベース）")
    row = 2

    group1 = [d for d in dams if d["pref"].startswith("北海道") and d["mgr_code"] != 1]
    group2 = [d for d in dams if not d["pref"].startswith("北海道")]
    targets = [(d, 1) for d in group1] + [(d, 2) for d in group2]

    def unknown_syms_of(d):
        return [(lnum, d["layers"][lnum]["symbol"])
                for lnum in range(1, 6)
                if d["layers"][lnum] and d["layers"][lnum].get("symbol")
                and d["layers"][lnum]["symbol"] not in ref_set]

    # 全ターゲットダムにわたる未知symbol頻度
    sym_freq = Counter()
    for d, _ in targets:
        for _, sym in unknown_syms_of(d):
            sym_freq[sym] += 1

    # per-dam スコアリング
    dam_rows = []
    for d, grp in targets:
        unk = unknown_syms_of(d)
        if not unk:
            continue
        known_cnt = sum(
            1 for lnum in range(1, 6)
            if d["layers"][lnum] and d["layers"][lnum].get("symbol")
            and d["layers"][lnum]["symbol"] in ref_set
        )
        # 展開価値: 未知symbolを解決したとき恩恵を受ける他ダム数の合計
        expansion = sum(sym_freq[sym] for _, sym in unk)
        h_bonus   = min((d["height"] or 0) / 5, 10)
        mgr_bonus = 3 if d["mgr_code"] in (1, 10) else 1
        priority  = expansion * 2 + len(unk) * 10 + h_bonus + mgr_bonus
        dam_rows.append({
            **d,
            "_grp":       grp,
            "_unk":       unk,
            "_unk_cnt":   len(unk),
            "_known":     known_cnt,
            "_expansion": expansion,
            "_priority":  priority,
        })

    dam_rows.sort(key=lambda x: -x["_priority"])

    g1_rows = [d for d in dam_rows if d["_grp"] == 1]
    g2_rows = [d for d in dam_rows if d["_grp"] == 2]

    # ── サマリー ──
    sec_title(ws, row, 1, 4, "▼ 調査対象サマリー")
    row += 1
    for ci, (h, w) in enumerate(zip(["項目", "グループ１", "グループ２", "合計"], [36, 12, 12, 12]), 1):
        hdr(ws, row, ci, h, w)
    row += 1
    g1_total = len(group1)
    g2_total = len(group2)
    for label, v1, v2 in [
        ("グループ内ダム総数",             g1_total, g2_total),
        ("未知symbolを持つダム数",         len(g1_rows), len(g2_rows)),
        ("全層推定不可（データなし）",
            sum(1 for d in g1_rows if d["layer_count"] == 0),
            sum(1 for d in g2_rows if d["layer_count"] == 0)),
        ("ユニーク未知symbol数",
            len(set(sym for d in g1_rows for _, sym in d["_unk"])),
            len(set(sym for d in g2_rows for _, sym in d["_unk"]))),
        ("最大展開価値（上位ダム）",
            max((d["_expansion"] for d in g1_rows), default=0),
            max((d["_expansion"] for d in g2_rows), default=0)),
    ]:
        body(ws, row, 1, label, bold=True)
        body(ws, row, 2, v1, align=RIGHT)
        body(ws, row, 3, v2, align=RIGHT)
        body(ws, row, 4, v1 + v2, align=RIGHT)
        row += 1
    row += 1

    # ── 優先スコアの考え方 ──
    sec_title(ws, row, 1, 13, "▼ 優先スコア算出方法")
    row += 1
    ws.column_dimensions["A"].width = 36
    for line in [
        "展開価値 = Σ（未知symbolごとに：そのsymbolを持つ他ターゲットダム数）",
        "  → 調査によってリファレンスが拡充されたとき、恩恵を受ける他ダムの総数に相当する",
        "優先スコア = 展開価値×2 ＋ 未知symbol数×10 ＋ 堤高ボーナス（height÷5、上限10pt）＋ 管理者ボーナス（国交省/開発局=3、他=1）",
        "グループ１: 北海道内・開発局以外　　グループ２: 北海道以外",
        "推奨: 優先スコア上位のダムから調査することで、リファレンス拡充効果が最大化される",
    ]:
        cell = ws.cell(row, 1, line); cell.font = BODY_FONT; cell.alignment = LEFT
        row += 1
    row += 1

    # ── 優先調査リスト ──
    sec_title(ws, row, 1, 13,
        f"▼ 調査優先ダムリスト（優先スコア降順、全{len(dam_rows)}件）")
    row += 1
    hdrs_d   = ["順位", "ダム名", "グループ", "所在地", "堤高(m)", "管理者",
                "層数", "既知sym数", "未知sym数", "展開価値", "優先スコア",
                "未知symbol一覧（展開頻度）", "既知symbol一覧"]
    widths_d = [6, 16, 10, 20, 8, 8, 6, 10, 10, 10, 10, 50, 40]
    hdr_row  = row
    for ci, (h, w) in enumerate(zip(hdrs_d, widths_d), 1):
        hdr(ws, row, ci, h, w)
    row += 1

    for rank, d in enumerate(dam_rows, 1):
        unk_str   = "  ".join(f"{sym}(×{sym_freq[sym]})" for _, sym in d["_unk"])
        known_str = "  ".join(
            d["layers"][lnum]["symbol"]
            for lnum in range(1, 6)
            if d["layers"][lnum] and d["layers"][lnum].get("symbol")
            and d["layers"][lnum]["symbol"] in ref_set
        )
        fill = RED_FILL if d["_unk_cnt"] >= 3 else YEL_FILL if d["_unk_cnt"] >= 2 \
               else ALT_FILL if rank % 2 == 0 else None
        for ci, v in enumerate([
            rank, d["name"], f"グループ{d['_grp']}",
            d["loc"][:18] if d["loc"] else "",
            d["height"], d["mgr_code"], d["layer_count"],
            d["_known"], d["_unk_cnt"],
            d["_expansion"], round(d["_priority"], 1),
            unk_str, known_str
        ], 1):
            body(ws, row, ci, v,
                 align=RIGHT if ci in (1, 5, 6, 7, 8, 9, 10, 11) else LEFT,
                 fill=fill)
        row += 1

    ws.freeze_panes = "A3"
    ws.auto_filter.ref = f"A{hdr_row}:{scol(len(hdrs_d))}{hdr_row}"
    row += 1

    # ── 都道府県別集計 ──
    sec_title(ws, row, 1, 7, "▼ 都道府県別 調査対象サマリー（調査対象ダム数降順）")
    row += 1
    for ci, (t, w) in enumerate(zip(
        ["都道府県", "グループ", "調査対象ダム数", "ユニーク未知sym数",
         "合計展開価値", "平均優先スコア", "代表ダム（上位3）"],
        [14, 10, 14, 16, 14, 14, 40]
    ), 1):
        hdr(ws, row, ci, t, w)
    row += 1

    pref_grp = defaultdict(list)
    for d in dam_rows:
        pref_grp[(d["pref"], d["_grp"])].append(d)

    for (pref, grp), pdams in sorted(pref_grp.items(), key=lambda x: -len(x[1])):
        unique_unk = len(set(sym for d in pdams for _, sym in d["_unk"]))
        total_exp  = sum(d["_expansion"] for d in pdams)
        avg_pri    = avg([d["_priority"] for d in pdams])
        top3names  = "、".join(d["name"] for d in pdams[:3])
        for ci, v in enumerate(
            [pref, f"グループ{grp}", len(pdams), unique_unk, total_exp, avg_pri, top3names], 1
        ):
            body(ws, row, ci, v, align=RIGHT if ci in (3, 4, 5, 6) else LEFT)
        row += 1


# ─── D2: 未掲載Symbol 調査効果分析 ──────────────────────────
def write_d2(wb, dams, ref_set):
    ws = wb.create_sheet("D2_未掲載Symbol調査効果")
    ws.sheet_view.showGridLines = False
    sheet_title(ws, 1, 1, 8,
        "■ D2 未掲載Symbol 調査効果分析 — 1symbolの解決がもたらす推定可能化ダム数")
    row = 2

    group1 = [d for d in dams if d["pref"].startswith("北海道") and d["mgr_code"] != 1]
    group2 = [d for d in dams if not d["pref"].startswith("北海道")]

    def unk_set(d):
        return {
            d["layers"][lnum]["symbol"]
            for lnum in range(1, 6)
            if d["layers"][lnum] and d["layers"][lnum].get("symbol")
            and d["layers"][lnum]["symbol"] not in ref_set
        }

    g1_unk = {d["name"]: unk_set(d) for d in group1}
    g2_unk = {d["name"]: unk_set(d) for d in group2}

    all_unk = (set().union(*g1_unk.values()) if g1_unk else set()) | \
              (set().union(*g2_unk.values()) if g2_unk else set())

    sym_g1_cnt = Counter(sym for s in g1_unk.values() for sym in s)
    sym_g2_cnt = Counter(sym for s in g2_unk.values() for sym in s)

    # 解決効果: そのsymbolが「最後の未知symbol」だったダム数（= 解決で全層推定可能になるダム数）
    def resolve_benefit(sym, groups_unk):
        return sum(1 for s in groups_unk.values() if sym in s and len(s) == 1)

    g1_syms = set().union(*g1_unk.values()) if g1_unk else set()
    g2_syms = set().union(*g2_unk.values()) if g2_unk else set()

    # ── サマリー ──
    sec_title(ws, row, 1, 3, "▼ 未掲載symbol サマリー")
    row += 1
    ws.column_dimensions["A"].width = 32
    ws.column_dimensions["B"].width = 12
    g1_only = g1_syms - g2_syms
    g2_only = g2_syms - g1_syms
    common  = g1_syms & g2_syms
    for label, val in [
        ("ユニーク未掲載symbol総数（G1+G2）", len(all_unk)),
        ("G1のみに存在する未掲載symbol",       len(g1_only)),
        ("G2のみに存在する未掲載symbol",       len(g2_only)),
        ("G1・G2共通の未掲載symbol",           len(common)),
    ]:
        body(ws, row, 1, label, bold=True)
        body(ws, row, 2, val, align=RIGHT)
        row += 1
    row += 1

    # ── symbol別 調査効果一覧 ──
    sec_title(ws, row, 1, 8,
        "▼ Symbol別 調査効果一覧（G1+G2合計出現数降順）— この1symbolを解決したとき得られる効果")
    row += 1
    for ci, (t, w) in enumerate(zip(
        ["未掲載symbol", "G1出現ダム数", "G2出現ダム数", "合計出現",
         "G1解決効果（全層推定可になるダム数）", "G2解決効果", "合計解決効果", "調査優先度評価"],
        [22, 14, 14, 10, 28, 14, 14, 30]
    ), 1):
        hdr(ws, row, ci, t, w)
    row += 1

    for sym in sorted(all_unk, key=lambda s: -(sym_g1_cnt[s] + sym_g2_cnt[s])):
        g1c  = sym_g1_cnt[sym]
        g2c  = sym_g2_cnt[sym]
        tot  = g1c + g2c
        g1b  = resolve_benefit(sym, g1_unk)
        g2b  = resolve_benefit(sym, g2_unk)
        totb = g1b + g2b
        if tot >= 10:   reason = "高頻度symbol — 調査価値大"
        elif totb >= 3: reason = f"解決で{totb}件が全層推定可能に"
        elif tot >= 3:  reason = "中頻度symbol — 調査推奨"
        else:           reason = "低頻度・単独出現"
        fill = GRN_FILL if tot >= 10 else YEL_FILL if tot >= 3 else RED_FILL
        for ci, v in enumerate([sym, g1c, g2c, tot, g1b, g2b, totb, reason], 1):
            body(ws, row, ci, v, align=RIGHT if ci in (2, 3, 4, 5, 6, 7) else LEFT, fill=fill)
        row += 1

    ws.freeze_panes = "A3"


# ─── 入力ファイル自動検索 ─────────────────────────────────────
def find_input_file(data_dir="data"):
    data_path = Path(data_dir)
    if not data_path.exists():
        raise FileNotFoundError(f"dataフォルダが見つかりません: {data_path.resolve()}")
    candidates = sorted(data_path.glob("*.xlsx"), key=lambda p: p.stat().st_mtime, reverse=True)
    if not candidates:
        raise FileNotFoundError("dataフォルダ内にxlsxファイルが見つかりません")
    for path in candidates:
        try:
            wb = load_workbook(path, read_only=True)
            sheets = wb.sheetnames; wb.close()
            if "全国ダム地質DB" in sheets and "Glossary" in sheets:
                return path
        except Exception:
            continue
    raise FileNotFoundError("dataフォルダ内に「全国ダム地質DB」と「Glossary」シートを持つxlsxが見つかりません")


def parse_args():
    p = argparse.ArgumentParser(
        description="全国ダム地質DB 地質推定可能性分析 v3",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog=(
            "例:\n"
            "  python3 分析3.py                              # data/ を自動検索\n"
            "  python3 分析3.py --input data/myfile.xlsx     # ファイルを直接指定\n"
            "  python3 分析3.py --output results/分析3.xlsx  # 出力先を指定\n"
        )
    )
    p.add_argument("--input",  default=None, help="入力xlsx（省略時はdata/フォルダを自動検索）")
    p.add_argument("--output", default=None, help="出力xlsx（省略時はdata/分析3結果_YYYYMMDD_HHMM.xlsx）")
    p.add_argument("--data",   default="data", help="自動検索フォルダ（デフォルト: data）")
    return p.parse_args()


def main():
    args = parse_args()

    if args.input:
        input_path = Path(args.input)
        if not input_path.exists():
            print(f"エラー: 入力ファイルが見つかりません: {input_path}")
            sys.exit(1)
    else:
        print(f"{args.data}/ フォルダから入力ファイルを自動検索...")
        input_path = find_input_file(args.data)
        mtime = datetime.datetime.fromtimestamp(input_path.stat().st_mtime)
        print(f"  → 使用ファイル: {input_path.name}  (更新: {mtime.strftime('%Y-%m-%d %H:%M')})")

    if args.output:
        output_path = Path(args.output)
        output_path.parent.mkdir(parents=True, exist_ok=True)
    else:
        ts = datetime.datetime.now().strftime("%Y%m%d_%H%M")
        output_path = Path(args.data) / f"分析3結果_{ts}.xlsx"

    print(f"読み込み: {input_path}")
    print(f"出力先:   {output_path}")

    wb_in = load_workbook(input_path)
    dams, glossary = load_data(wb_in)

    hokdev = [d for d in dams if d["pref"].startswith("北海道") and d["mgr_code"] == 1]
    ref_set, ref_detail = build_ref_set(hokdev)

    g1_cnt = sum(1 for d in dams if d["pref"].startswith("北海道") and d["mgr_code"] != 1)
    g2_cnt = sum(1 for d in dams if not d["pref"].startswith("北海道"))
    print(f"総ダム数: {len(dams)}")
    print(f"  北海道開発局（リファレンス）: {len(hokdev)}件 / リファレンスsymbol: {len(ref_set)}種")
    print(f"  グループ１（北海道内・開発局外）: {g1_cnt}件")
    print(f"  グループ２（北海道以外）: {g2_cnt}件")

    wb_out = Workbook()
    wb_out.remove(wb_out.active)

    print("C1 北海道開発局リファレンス...")
    write_c1(wb_out, hokdev, glossary, ref_set, ref_detail)
    print("C2 グループ1推定可能性（北海道内・開発局外）...")
    write_c2(wb_out, dams, ref_set)
    print("C3 グループ2推定可能性（北海道以外）...")
    write_c3(wb_out, dams, ref_set)
    print("D1 調査優先ダムリスト...")
    write_d1(wb_out, dams, ref_set)
    print("D2 未掲載Symbol調査効果...")
    write_d2(wb_out, dams, ref_set)

    wb_out.save(output_path)
    print(f"\n保存完了: {output_path}")


if __name__ == "__main__":
    main()
