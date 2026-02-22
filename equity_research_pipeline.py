#!/usr/bin/env python3
from __future__ import annotations

import argparse
import json
import math
import urllib.error
import urllib.request
import xml.sax.saxutils as saxutils
from dataclasses import dataclass
from pathlib import Path
from typing import Dict, List, Optional, Sequence
from zipfile import ZIP_DEFLATED, ZipFile

SCENARIOS = [0.5, 1.0, 1.5, 2.0]
PE_SENSITIVITY_MARGIN = 0.32
FIXED_MARKET_CAPS = {
    "AXTI": 1.65e9,
    "Nittobo": 799.7e9,
    "Ibiden": 2.5797e12,
    "Coherent": 13.2e9,
    "Sumitomo Electric": 3.78e12,
}


@dataclass
class CompanyConfig:
    name: str
    ticker: str
    ai_exposure_ratio: float
    ai_incremental_margin: float
    base_tags: List[str]


COMPANIES: List[CompanyConfig] = [
    CompanyConfig("AXTI", "AXTI", 0.35, 0.45, ["光芯片", "InP材料"]),
    CompanyConfig("Coherent", "COHR", 0.30, 0.40, ["光芯片", "光模块"]),
    CompanyConfig("Sumitomo Electric", "5802.T", 0.18, 0.30, ["光模块", "AI PCB"]),
    CompanyConfig("Ibiden", "4062.T", 0.28, 0.35, ["AI PCB", "CCL / T-Glass"]),
    CompanyConfig("Shinko", "6967.T", 0.26, 0.35, ["AI PCB", "CCL / T-Glass"]),
    CompanyConfig("Nittobo", "3110.T", 0.24, 0.33, ["CCL / T-Glass", "光模块"]),
]

TAG_KEYWORDS = {
    "光芯片": ["laser", "photonic", "optical", "chip"],
    "InP材料": ["indium phosphide", "inp", "substrate", "epi"],
    "CCL / T-Glass": ["glass", "laminate", "ccl", "t-glass"],
    "AI PCB": ["package", "substrate", "pcb", "interconnect"],
    "光模块": ["transceiver", "datacom", "module", "optic"],
}


def load_fallback(path: Path) -> Dict:
    return json.loads(path.read_text(encoding="utf-8"))


def fetch_yahoo_summary(ticker: str) -> Optional[Dict]:
    url = (
        f"https://query2.finance.yahoo.com/v10/finance/quoteSummary/{ticker}"
        "?modules=incomeStatementHistory,assetProfile,defaultKeyStatistics"
    )
    req = urllib.request.Request(url, headers={"User-Agent": "Mozilla/5.0"})
    try:
        with urllib.request.urlopen(req, timeout=15) as resp:
            payload = json.loads(resp.read().decode("utf-8"))
        return payload["quoteSummary"]["result"][0]
    except Exception:
        return None


def parse_income_history(raw: Dict) -> List[Dict]:
    history = raw.get("incomeStatementHistory", {}).get("incomeStatementHistory", [])
    out = []
    for row in history:
        end = row.get("endDate", {}).get("fmt")
        if not end:
            continue
        year = int(end[:4])
        rev = row.get("totalRevenue", {}).get("raw")
        gp = row.get("grossProfit", {}).get("raw")
        ni = row.get("netIncome", {}).get("raw")
        if rev is None or ni is None:
            continue
        out.append({"year": year, "revenue": rev, "gross_profit": gp, "net_income": ni})
    out.sort(key=lambda x: x["year"])
    return out


def infer_tags(base_tags: Sequence[str], summary: str) -> List[str]:
    tags = set(base_tags)
    text = (summary or "").lower()
    for tag, kws in TAG_KEYWORDS.items():
        if any(k in text for k in kws):
            tags.add(tag)
    return sorted(tags)


def cagr(last: float, first: float, periods: int) -> float:
    if periods <= 0 or first <= 0:
        return 0.0
    return (last / first) ** (1 / periods) - 1


def build_rows(config: CompanyConfig, annual: List[Dict], shares: float, summary: str) -> Dict[str, List[List]]:
    latest = annual[-1]
    revenue = float(latest["revenue"])
    gross_profit = float(latest.get("gross_profit") or 0.0)
    net_income = float(latest["net_income"])
    gross_margin = gross_profit / revenue if revenue else 0.0
    net_margin = net_income / revenue if revenue else 0.0
    ai_revenue = revenue * config.ai_exposure_ratio
    tags = ", ".join(infer_tags(config.base_tags, summary))

    snap = [[
        config.name,
        config.ticker,
        latest["year"],
        revenue,
        gross_margin,
        net_income,
        config.ai_exposure_ratio,
        ai_revenue,
        tags,
        shares,
    ]]

    ni_3y = [x["net_income"] for x in annual[-3:]] if len(annual) >= 3 else [x["net_income"] for x in annual]
    base_growth = cagr(ni_3y[-1], ni_3y[0], len(ni_3y) - 1)
    years_to_2026 = max(0, 2026 - latest["year"])
    ni_2026_base = net_income * ((1 + base_growth) ** years_to_2026)

    scenario_rows = []
    ni2026_rows = []
    eps_rows = []
    for s in SCENARIOS:
        incr = ai_revenue * s * config.ai_incremental_margin
        ni_s = net_income + incr
        ni_2026 = ni_2026_base + incr
        eps = ni_s / shares if shares else 0.0
        scenario_rows.append([
            config.name,
            f"+{int(s*100)}%",
            s,
            ai_revenue,
            incr,
            net_income,
            ni_s,
            ni_2026,
            eps,
            net_margin,
            ni_s / revenue if revenue else 0.0,
        ])
        ni2026_rows.append([config.name, f"+{int(s*100)}%", ni_2026])
        eps_rows.append([config.name, int(s * 100), eps])

    pe_rows = []
    market_cap = FIXED_MARKET_CAPS.get(config.name)
    if market_cap is not None:
        for s in SCENARIOS:
            delta_ni = ai_revenue * s * PE_SENSITIVITY_MARGIN
            new_ni = net_income + delta_ni
            eps = new_ni / shares if shares else 0.0
            forward_pe = market_cap / new_ni if new_ni else 0.0
            pe_rows.append([
                config.name,
                f"+{int(s * 100)}%",
                net_income,
                new_ni,
                eps,
                forward_pe,
            ])

    hist_rows = []
    for row in annual:
        rev = float(row["revenue"])
        gp = float(row.get("gross_profit") or 0.0)
        ni = float(row["net_income"])
        hist_rows.append([
            config.name,
            config.ticker,
            row["year"],
            rev,
            gp,
            ni,
            gp / rev if rev else 0.0,
            ni / rev if rev else 0.0,
        ])

    return {
        "AI收入占比": snap,
        "利润弹性情景表": scenario_rows,
        "2026E净利润变化_long": ni2026_rows,
        "历史财报": hist_rows,
        "EPS": eps_rows,
        "PE Sensitivity": pe_rows,
    }


def write_svg_eps(eps_rows: List[List], out_path: Path) -> None:
    by_company: Dict[str, List[List]] = {}
    for c, x, y in eps_rows:
        by_company.setdefault(c, []).append([x, y])

    width, height = 900, 520
    margin = 60
    xs = [50, 100, 150, 200]
    all_y = [p[1] for pts in by_company.values() for p in pts] or [0.0]
    ymin, ymax = min(all_y), max(all_y)
    if math.isclose(ymax, ymin):
        ymax = ymin + 1.0

    def x_map(v):
        return margin + (v - min(xs)) * (width - 2 * margin) / (max(xs) - min(xs))

    def y_map(v):
        return height - margin - (v - ymin) * (height - 2 * margin) / (ymax - ymin)

    colors = ["#1f77b4", "#ff7f0e", "#2ca02c", "#d62728", "#9467bd", "#8c564b"]
    parts = [
        f'<svg xmlns="http://www.w3.org/2000/svg" width="{width}" height="{height}">',
        '<rect width="100%" height="100%" fill="white"/>',
        f'<text x="{width/2}" y="30" text-anchor="middle" font-size="18">EPS Sensitivity</text>',
        f'<line x1="{margin}" y1="{height-margin}" x2="{width-margin}" y2="{height-margin}" stroke="black"/>',
        f'<line x1="{margin}" y1="{margin}" x2="{margin}" y2="{height-margin}" stroke="black"/>',
    ]
    for xv in xs:
        x = x_map(xv)
        parts.append(f'<line x1="{x}" y1="{height-margin}" x2="{x}" y2="{height-margin+5}" stroke="black"/>')
        parts.append(f'<text x="{x}" y="{height-margin+22}" text-anchor="middle" font-size="12">{xv}%</text>')

    for i, (company, pts) in enumerate(sorted(by_company.items())):
        pts.sort(key=lambda x: x[0])
        poly = " ".join(f"{x_map(x)},{y_map(y)}" for x, y in pts)
        color = colors[i % len(colors)]
        parts.append(f'<polyline fill="none" stroke="{color}" stroke-width="2" points="{poly}"/>')
        lx = width - margin + 10
        ly = margin + i * 18
        parts.append(f'<line x1="{lx}" y1="{ly}" x2="{lx+15}" y2="{ly}" stroke="{color}" stroke-width="2"/>')
        parts.append(f'<text x="{lx+20}" y="{ly+4}" font-size="11">{saxutils.escape(company)}</text>')

    parts.append("</svg>")
    out_path.parent.mkdir(parents=True, exist_ok=True)
    out_path.write_text("\n".join(parts), encoding="utf-8")


def col_letter(n: int) -> str:
    s = ""
    while n:
        n, r = divmod(n - 1, 26)
        s = chr(65 + r) + s
    return s


def sheet_xml(rows: List[List]) -> str:
    row_xml = []
    for r_idx, row in enumerate(rows, start=1):
        cells = []
        for c_idx, val in enumerate(row, start=1):
            ref = f"{col_letter(c_idx)}{r_idx}"
            if isinstance(val, (int, float)):
                cells.append(f'<c r="{ref}"><v>{val}</v></c>')
            else:
                text = saxutils.escape(str(val))
                cells.append(f'<c r="{ref}" t="inlineStr"><is><t>{text}</t></is></c>')
        row_xml.append(f'<row r="{r_idx}">{"".join(cells)}</row>')
    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">'
        '<sheetData>' + "".join(row_xml) + '</sheetData></worksheet>'
    )


def write_xlsx(path: Path, sheets: Dict[str, List[List]]) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    with ZipFile(path, "w", compression=ZIP_DEFLATED) as z:
        z.writestr("[Content_Types].xml", (
            '<?xml version="1.0" encoding="UTF-8"?>'
            '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
            '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
            '<Default Extension="xml" ContentType="application/xml"/>'
            '<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>' +
            "".join(f'<Override PartName="/xl/worksheets/sheet{i}.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>' for i in range(1, len(sheets)+1)) +
            '</Types>'
        ))
        z.writestr("_rels/.rels", (
            '<?xml version="1.0" encoding="UTF-8"?>'
            '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
            '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>'
            '</Relationships>'
        ))
        z.writestr("xl/workbook.xml", (
            '<?xml version="1.0" encoding="UTF-8"?>'
            '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">'
            '<sheets>' + "".join(f'<sheet name="{saxutils.escape(name[:31])}" sheetId="{i}" r:id="rId{i}"/>' for i, name in enumerate(sheets.keys(), start=1)) + '</sheets>'
            '</workbook>'
        ))
        z.writestr("xl/_rels/workbook.xml.rels", (
            '<?xml version="1.0" encoding="UTF-8"?>'
            '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
            "".join(f'<Relationship Id="rId{i}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet{i}.xml"/>' for i in range(1, len(sheets)+1)) +
            '</Relationships>'
        ))
        for i, (_, rows) in enumerate(sheets.items(), start=1):
            z.writestr(f"xl/worksheets/sheet{i}.xml", sheet_xml(rows))


def run(output_xlsx: Path, output_svg: Path, fallback_path: Path, offline: bool) -> None:
    fallback = load_fallback(fallback_path)
    sheets = {
        "AI收入占比": [["Company", "Ticker", "LatestYear", "Revenue", "GrossMargin", "NetIncome", "AIRevenueRatio", "AIRevenue", "Tags", "SharesOutstanding"]],
        "利润弹性情景表": [["Company", "Scenario", "PriceIncrease", "AIRevenue", "IncrementalProfit", "NetIncome_Base", "NetIncome_Scenario", "NetIncome_2026E", "EPS_Scenario", "NetMargin_Base", "NetMargin_Scenario"]],
        "2026E净利润变化": [["Company", "+50%", "+100%", "+150%", "+200%"]],
        "历史财报": [["Company", "Ticker", "Year", "Revenue", "GrossProfit", "NetIncome", "GrossMargin", "NetMargin"]],
        "PE Sensitivity": [["Company", "ASP Increase", "Base NI", "New NI", "EPS", "Forward PE"]],
    }
    eps_points = []
    ni_map: Dict[str, Dict[str, float]] = {}

    for cfg in COMPANIES:
        annual = None
        summary = ""
        shares = 0.0
        if not offline:
            raw = fetch_yahoo_summary(cfg.ticker)
            if raw:
                annual = parse_income_history(raw)
                summary = raw.get("assetProfile", {}).get("longBusinessSummary", "")
                shares = float(raw.get("defaultKeyStatistics", {}).get("sharesOutstanding", {}).get("raw") or 0.0)

        if not annual:
            fb = fallback[cfg.name]
            annual = fb["annual"]
            summary = fb.get("summary", "")
            shares = fb.get("shares_outstanding", 0.0)
            print(f"[WARN] {cfg.name}: 使用离线样本数据")
        else:
            print(f"[OK] {cfg.name}: 在线数据抓取成功")

        built = build_rows(cfg, annual, shares, summary)
        sheets["AI收入占比"].extend(built["AI收入占比"])
        sheets["利润弹性情景表"].extend(built["利润弹性情景表"])
        sheets["历史财报"].extend(built["历史财报"])
        sheets["PE Sensitivity"].extend(built["PE Sensitivity"])
        eps_points.extend(built["EPS"])

        ni_map[cfg.name] = {r[1]: r[2] for r in built["2026E净利润变化_long"]}

    for c in [cfg.name for cfg in COMPANIES]:
        row = [c, ni_map[c].get("+50%", 0.0), ni_map[c].get("+100%", 0.0), ni_map[c].get("+150%", 0.0), ni_map[c].get("+200%", 0.0)]
        sheets["2026E净利润变化"].append(row)

    write_xlsx(output_xlsx, sheets)
    write_svg_eps(eps_points, output_svg)
    print(f"输出 Excel: {output_xlsx}")
    print(f"输出图表: {output_svg}")


def main() -> None:
    p = argparse.ArgumentParser(description="AI产业链投研数据库系统")
    p.add_argument("--output", default="output.xlsx")
    p.add_argument("--chart", default="output/eps_sensitivity.svg")
    p.add_argument("--fallback", default="data/fallback_financials.json")
    p.add_argument("--offline", action="store_true", help="强制使用离线样本数据")
    args = p.parse_args()

    run(Path(args.output), Path(args.chart), Path(args.fallback), args.offline)


if __name__ == "__main__":
    main()
