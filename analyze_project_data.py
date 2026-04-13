"""
analyze_project_data.py  — v2
==============================
Reads an Excel file with project lessons-learned data and produces:
  1. An enriched Excel workbook  (Raw Data + 5 insight sheets + embedded charts)
  2. A PowerPoint presentation   (title, agenda, chart slides, tables, takeaways)

Usage:
    python analyze_project_data.py  my_data.xlsx
    python analyze_project_data.py  my_data.xlsx  --sheet "Sheet2"  --out-dir ./results

Required Python packages:
    pip install pandas openpyxl python-pptx matplotlib

Optional (enables word stemming + word-cloud):
    pip install nltk wordcloud pillow

Column recognition is automatic — headings are matched case-insensitively.
Recognised headings include:
  project name, organization, project type, project stage,
  original category, new category, lesson learned title,
  threat/opportunity, description of threat/opportunity,
  response of threat/opportunity, lessons learned
"""

import sys, re, argparse, warnings
from collections import Counter
from pathlib import Path

import pandas as pd
import numpy as np
import matplotlib; matplotlib.use("Agg")
import matplotlib.pyplot as plt

from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.drawing.image import Image as XLImage

from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN

warnings.filterwarnings("ignore")

# ── Optional NLP (graceful fallback if not installed) ─────────────────────────
try:
    import nltk
    from nltk.corpus import stopwords as nltk_sw
    from nltk.stem import WordNetLemmatizer
    for _p in ["punkt","stopwords","wordnet","omw-1.4","punkt_tab"]:
        try: nltk.download(_p, quiet=True)
        except: pass
    NLTK_OK = True
except ImportError:
    NLTK_OK = False

try:
    from wordcloud import WordCloud
    WC_OK = True
except ImportError:
    WC_OK = False

# ── Design constants ──────────────────────────────────────────────────────────
PAL = {
    "primary":     "1C4E80",
    "secondary":   "0091D5",
    "accent1":     "EA6A47",
    "accent2":     "A5D8DD",
    "accent3":     "488A99",
    "light":       "F0F4F8",
    "dark":        "1A2B3C",
    "threat":      "C0392B",
    "opportunity": "27AE60",
}
OPXL_HEX  = {k: f"FF{v}" for k, v in PAL.items()}            # openpyxl ARGB
PPTX_RGB  = {k: RGBColor.from_string(v) for k, v in PAL.items()}  # pptx RGB

# ── Column aliases (fuzzy, case-insensitive) ──────────────────────────────────
ALIASES = {
    "project_name":      ["project name", "project"],
    "organization":      ["organization", "org", "department"],
    "project_type":      ["project type", "type"],
    "project_stage":     ["project stage", "stage", "phase"],
    "original_category": ["original category", "orig category", "orig cat"],
    "new_category":      ["new category", "revised category", "new cat"],
    "lesson_title":      ["lesson learned title", "lesson title", "title"],
    "threat_opp_type":   ["threat/opportunity", "threat or opportunity",
                          "risk type", "type of risk"],
    "description":       ["description of threat/opportunity", "description",
                          "threat description", "opportunity description", "desc"],
    "response":          ["response of threat/opportunity", "response",
                          "mitigation", "action"],
    "lessons_learned":   ["lessons learned", "lesson learned", "lessons",
                          "key takeaway", "takeaway"],
}
EXTRA_SW = {
    "project","projects","use","used","using","need","needs","ensure","must",
    "also","may","well","one","two","make","team","time","new","work","include",
    "process","provide","within","would","could","should","will","many","often",
    "however","therefore","thus","due","per","key","based","required",
    "high","low","good","important",
}

# =============================================================================
# DATA LOADING
# =============================================================================

def load_data(path: str, sheet: str = None) -> pd.DataFrame:
    xl = pd.ExcelFile(path)
    df = pd.read_excel(xl, sheet_name=sheet or xl.sheet_names[0], dtype=str).fillna("")
    df.columns = df.columns.str.strip()
    return df


def map_columns(df: pd.DataFrame) -> dict:
    lower = {c.lower().strip(): c for c in df.columns}
    result = {}
    for canon, aliases in ALIASES.items():
        for alias in aliases:
            if alias in lower:
                result[canon] = lower[alias]
                break
    return result


# =============================================================================
# TEXT UTILITIES
# =============================================================================

def build_stopwords() -> set:
    sw = EXTRA_SW.copy()
    if NLTK_OK:
        sw |= set(nltk_sw.words("english"))
    else:
        sw |= {
            "the","a","an","and","or","but","in","on","at","to","for","of",
            "with","by","from","is","are","was","were","be","been","being",
            "have","has","had","do","does","did","not","no","it","its",
            "they","them","their","we","our","you","your","he","she",
            "who","which","what","how","when","where","if","as",
            "this","that","these","those","then","here","there",
        }
    return sw


def tokenize(text: str, sw: set) -> list:
    tokens = re.sub(r"[^a-z\s]", " ", str(text).lower()).split()
    if NLTK_OK:
        lem = WordNetLemmatizer()
        tokens = [lem.lemmatize(t) for t in tokens]
    return [t for t in tokens if t not in sw and len(t) > 2]


def top_keywords(texts: list, sw: set, n: int = 20) -> list:
    c = Counter()
    for t in texts:
        c.update(tokenize(t, sw))
    return c.most_common(n)


def freq_table(df: pd.DataFrame, col: str, label: str) -> pd.DataFrame:
    counts = df[col].value_counts().reset_index()
    counts.columns = [label, "Count"]
    counts["Percentage"] = (counts["Count"] / counts["Count"].sum() * 100).round(1)
    return counts


def split_threats_opps(df: pd.DataFrame, col_map: dict):
    threats, opps = [], []
    dc = col_map.get("description", "")
    tc = col_map.get("threat_opp_type", "")
    if not dc:
        return threats, opps
    for _, row in df.iterrows():
        txt = str(row.get(dc, ""))
        lbl = str(row.get(tc, "")).lower() if tc else ""
        (opps if "opportunit" in lbl else threats).append(txt)
    return threats, opps


def stage_cat_pivot(df: pd.DataFrame, col_map: dict):
    sc = col_map.get("project_stage")
    cc = col_map.get("new_category") or col_map.get("original_category")
    if not sc or not cc:
        return None
    return pd.crosstab(df[sc], df[cc])


# =============================================================================
# CHART GENERATION
# =============================================================================

CHART_DIR = Path("/tmp/pld_charts")
CHART_DIR.mkdir(exist_ok=True)

MULTI_COLORS = [
    f"#{PAL['primary']}", f"#{PAL['secondary']}", f"#{PAL['accent1']}",
    f"#{PAL['accent2']}", f"#{PAL['accent3']}", "#7EC8E3", "#FFD166",
    "#EF476F", "#06D6A0", "#118AB2"
] * 4


def _setup(fig, ax):
    fig.patch.set_facecolor("#FAFAFA")
    ax.set_facecolor("#FAFAFA")
    ax.spines[["top", "right"]].set_visible(False)


def make_bar(labels, values, title, fname, color=None, horiz=False, figsize=(8, 4)) -> str:
    color = color or f"#{PAL['secondary']}"
    fig, ax = plt.subplots(figsize=figsize)
    _setup(fig, ax)
    pos = range(len(labels))
    if horiz:
        ax.barh(pos, values, color=color, edgecolor="white", linewidth=0.5)
        ax.set_yticks(pos)
        ax.set_yticklabels(labels, fontsize=9)
        ax.invert_yaxis()
        ax.set_xlabel("Count", fontsize=9)
        for i, v in enumerate(values):
            ax.text(v + max(values) * 0.01, i, str(v), va="center", fontsize=8)
    else:
        ax.bar(pos, values, color=color, edgecolor="white", linewidth=0.5)
        ax.set_xticks(pos)
        ax.set_xticklabels(labels, rotation=35, ha="right", fontsize=9)
        ax.set_ylabel("Count", fontsize=9)
        for i, v in enumerate(values):
            ax.text(i, v + max(values) * 0.01, str(v), ha="center", fontsize=8)
    ax.set_title(title, fontsize=11, fontweight="bold", color=f"#{PAL['dark']}", pad=10)
    plt.tight_layout()
    p = CHART_DIR / fname
    plt.savefig(p, dpi=150, bbox_inches="tight")
    plt.close()
    return str(p)


def make_pie(labels, values, title, fname) -> str:
    fig, ax = plt.subplots(figsize=(7, 5))
    fig.patch.set_facecolor("#FAFAFA")
    wedges, _, autotexts = ax.pie(
        values, colors=MULTI_COLORS[:len(values)], autopct="%1.1f%%",
        startangle=140, wedgeprops=dict(edgecolor="white", linewidth=1.5))
    for at in autotexts:
        at.set_fontsize(8)
    ax.legend(wedges, labels, loc="center left", bbox_to_anchor=(1, 0, 0.5, 1), fontsize=8)
    ax.set_title(title, fontsize=11, fontweight="bold", color=f"#{PAL['dark']}", pad=10)
    plt.tight_layout()
    p = CHART_DIR / fname
    plt.savefig(p, dpi=150, bbox_inches="tight")
    plt.close()
    return str(p)


def make_kw_chart(kw_list, title, fname, color=None) -> str:
    if not kw_list:
        return None
    labels = [k for k, _ in kw_list[:15]]
    values = [v for _, v in kw_list[:15]]
    return make_bar(labels, values, title, fname, color=color, horiz=True, figsize=(8, 5))


def make_wordcloud(texts, fname) -> str:
    if not WC_OK or not texts:
        return None
    sw = build_stopwords()
    combined = " ".join(" ".join(tokenize(t, sw)) for t in texts)
    if not combined.strip():
        return None
    wc = WordCloud(width=800, height=400, background_color="white",
                   colormap="Blues", max_words=80,
                   stopwords=sw, collocations=False).generate(combined)
    fig, ax = plt.subplots(figsize=(8, 4))
    ax.imshow(wc, interpolation="bilinear")
    ax.axis("off")
    plt.tight_layout(pad=0)
    p = CHART_DIR / fname
    plt.savefig(p, dpi=150, bbox_inches="tight")
    plt.close()
    return str(p)


def make_stacked_bar(pivot_df, title, fname) -> str:
    fig, ax = plt.subplots(figsize=(10, 5))
    _setup(fig, ax)
    bottom = np.zeros(len(pivot_df))
    for i, col in enumerate(pivot_df.columns):
        ax.bar(pivot_df.index, pivot_df[col], bottom=bottom,
               label=col, color=MULTI_COLORS[i % len(MULTI_COLORS)],
               edgecolor="white", linewidth=0.5)
        bottom += pivot_df[col].values
    ax.set_title(title, fontsize=11, fontweight="bold", color=f"#{PAL['dark']}", pad=10)
    ax.set_ylabel("Count", fontsize=9)
    ax.set_xticklabels(pivot_df.index, rotation=30, ha="right", fontsize=9)
    ax.legend(bbox_to_anchor=(1.01, 1), loc="upper left", fontsize=8)
    plt.tight_layout()
    p = CHART_DIR / fname
    plt.savefig(p, dpi=150, bbox_inches="tight")
    plt.close()
    return str(p)


def generate_charts(df: pd.DataFrame, col_map: dict, sw: set) -> dict:
    charts = {}

    # Distribution bars
    for field, fname, color in [
        ("project_type",  "type_bar.png",     f"#{PAL['primary']}"),
        ("project_stage", "stage_bar.png",    f"#{PAL['secondary']}"),
        ("new_category",  "category_bar.png", f"#{PAL['accent1']}"),
        ("organization",  "org_bar.png",       f"#{PAL['accent3']}"),
    ]:
        col = col_map.get(field)
        if not col:
            continue
        df_f   = freq_table(df, col, "")
        labels = df_f.iloc[:, 0].tolist()[:12]
        values = df_f["Count"].tolist()[:12]
        title  = f"Distribution by {field.replace('_', ' ').title()}"
        horiz  = max((len(str(l)) for l in labels), default=0) > 12
        charts[fname.replace(".png", "")] = make_bar(
            labels, values, title, fname, color=color, horiz=horiz)

    # Threat/Opportunity pie
    to_col = col_map.get("threat_opp_type")
    if to_col:
        df_to = freq_table(df, to_col, "")
        if len(df_to) >= 1:
            charts["threat_opp_pie"] = make_pie(
                df_to.iloc[:, 0].tolist(), df_to["Count"].tolist(),
                "Threat vs. Opportunity Split", "threat_opp_pie.png")

    # Keyword charts
    all_text_cols = [col_map.get(f) for f in
                     ["description", "response", "lessons_learned", "lesson_title"]
                     if col_map.get(f)]
    all_texts = []
    for c in all_text_cols:
        all_texts += df[c].tolist()

    if all_texts:
        kws = top_keywords(all_texts, sw, 15)
        charts["keyword_all"] = make_kw_chart(
            kws, "Top Keywords — All Text", "keyword_all.png", f"#{PAL['primary']}")

    lc = col_map.get("lessons_learned")
    if lc:
        kws = top_keywords(df[lc].tolist(), sw, 15)
        charts["keyword_lessons"] = make_kw_chart(
            kws, "Top Keywords — Lessons Learned",
            "keyword_lessons.png", f"#{PAL['accent3']}")

    t_texts, o_texts = split_threats_opps(df, col_map)
    if t_texts:
        charts["threat_kw"] = make_kw_chart(
            top_keywords(t_texts, sw, 15),
            "Top Keywords — Threats", "threat_kw.png", f"#{PAL['threat']}")
    if o_texts:
        charts["opp_kw"] = make_kw_chart(
            top_keywords(o_texts, sw, 15),
            "Top Keywords — Opportunities", "opp_kw.png", f"#{PAL['opportunity']}")

    if all_texts:
        charts["wordcloud"] = make_wordcloud(all_texts, "wordcloud.png")

    pivot = stage_cat_pivot(df, col_map)
    if pivot is not None and not pivot.empty:
        charts["stage_cat_stack"] = make_stacked_bar(
            pivot, "Lessons by Stage & Category", "stage_cat_stack.png")

    print(f"  {sum(1 for v in charts.values() if v)} chart(s) generated.")
    return charts


# =============================================================================
# EXCEL BUILDER
# =============================================================================

def _fill(c):
    return PatternFill("solid", fgColor=c)

def _font(bold=False, color="FF000000", size=10):
    return Font(bold=bold, color=color, size=size, name="Calibri")

def _align(h="left", v="center", wrap=False):
    return Alignment(horizontal=h, vertical=v, wrap_text=wrap)

def _bdr():
    s = Side(style="thin", color="FFD0D0D0")
    return Border(left=s, right=s, top=s, bottom=s)


def style_header(ws, row, ncols, bg=None, fg="FFFFFFFF"):
    bg = bg or OPXL_HEX["primary"]
    for c in range(1, ncols + 1):
        cell = ws.cell(row=row, column=c)
        cell.fill = _fill(bg)
        cell.font = _font(True, fg, 10)
        cell.alignment = _align("center")
        cell.border = _bdr()


def style_data(ws, r0, r1, ncols):
    for r in range(r0, r1 + 1):
        bg = OPXL_HEX["light"] if r % 2 == 0 else "FFFFFFFF"
        for c in range(1, ncols + 1):
            cell = ws.cell(row=r, column=c)
            cell.fill = _fill(bg)
            cell.font = _font(size=10)
            cell.alignment = _align(wrap=True)
            cell.border = _bdr()


def write_freq_block(ws, df_f, start_row, title="") -> int:
    if title:
        ws.cell(row=start_row, column=1).value = title
        ws.cell(row=start_row, column=1).font = _font(True, OPXL_HEX["primary"], 12)
        start_row += 1
    for ci, col in enumerate(df_f.columns, 1):
        ws.cell(row=start_row, column=ci).value = col
    style_header(ws, start_row, len(df_f.columns))
    for ri, row in df_f.iterrows():
        for ci, val in enumerate(row, 1):
            ws.cell(row=start_row + ri + 1, column=ci).value = val
    style_data(ws, start_row + 1, start_row + len(df_f), len(df_f.columns))
    return start_row + len(df_f) + 2


def embed_image(ws, path, anchor, w_cm=14, h_cm=7):
    if not path or not Path(path).exists():
        return
    img = XLImage(path)
    img.width  = int(w_cm * 37.795)
    img.height = int(h_cm * 37.795)
    ws.add_image(img, anchor)


def build_excel(df: pd.DataFrame, col_map: dict, charts: dict, sw: set, out_path: str):
    wb = Workbook()

    # ── Sheet 1: Raw Data ──────────────────────────────────────────────────────
    ws = wb.active
    ws.title = "Raw Data"
    ws.freeze_panes = "A2"
    for ci, col in enumerate(df.columns, 1):
        ws.cell(row=1, column=ci).value = col
    style_header(ws, 1, len(df.columns))
    for ri, row in df.iterrows():
        for ci, val in enumerate(row, 1):
            cell = ws.cell(row=ri + 2, column=ci)
            cell.value     = val
            cell.alignment = _align(wrap=True)
            cell.border    = _bdr()
            if ri % 2 == 1:
                cell.fill = _fill(OPXL_HEX["light"])
    for col in ws.columns:
        w = max((len(str(c.value or "")) for c in col), default=10)
        ws.column_dimensions[col[0].column_letter].width = min(w + 4, 45)

    # ── Sheet 2: Summary Dashboard ─────────────────────────────────────────────
    ws2 = wb.create_sheet("Summary Dashboard")
    ws2.merge_cells("A1:H2")
    c = ws2["A1"]
    c.value      = "PROJECT LESSONS LEARNED  —  Insights Dashboard"
    c.font       = _font(True, "FFFFFFFF", 16)
    c.fill       = _fill(OPXL_HEX["primary"])
    c.alignment  = _align("center", "center")

    tc = col_map.get("project_type")
    sc = col_map.get("project_stage")
    cc = col_map.get("new_category") or col_map.get("original_category")
    kpis = [
        ("Total Records",  len(df),                              OPXL_HEX["primary"]),
        ("Project Types",  df[tc].nunique() if tc else "—",     OPXL_HEX["secondary"]),
        ("Stages",         df[sc].nunique() if sc else "—",     OPXL_HEX["accent1"]),
        ("Categories",     df[cc].nunique() if cc else "—",     OPXL_HEX["accent3"]),
    ]
    for i, (label, val, col_hex) in enumerate(kpis):
        cs = i * 2 + 1
        ws2.merge_cells(start_row=4, start_column=cs, end_row=4, end_column=cs + 1)
        ws2.merge_cells(start_row=5, start_column=cs, end_row=5, end_column=cs + 1)
        c1 = ws2.cell(row=4, column=cs)
        c1.value = label; c1.font = _font(True, "FF606060", 9); c1.alignment = _align("center")
        c2 = ws2.cell(row=5, column=cs)
        c2.value = val; c2.font = _font(True, col_hex, 20)
        c2.alignment = _align("center"); c2.fill = _fill(OPXL_HEX["light"])

    for key, col_l, row_n in [
        ("type_bar",       "A",  7),  ("stage_bar",      "I",  7),
        ("category_bar",   "A", 29),  ("threat_opp_pie", "I", 29),
        ("keyword_all",    "A", 51),  ("wordcloud",      "I", 51),
        ("stage_cat_stack","A", 73),
    ]:
        embed_image(ws2, charts.get(key), f"{col_l}{row_n}", 13, 8)

    for ltr in "ABCDEFGHIJKLMNOP":
        ws2.column_dimensions[ltr].width = 10

    # ── Sheet 3: Frequency Tables ──────────────────────────────────────────────
    ws3  = wb.create_sheet("Frequency Tables")
    crow = 1
    for field in ["project_type","project_stage","new_category",
                  "original_category","organization","threat_opp_type"]:
        col = col_map.get(field)
        if not col:
            continue
        df_f = freq_table(df, col, field.replace("_", " ").title())
        crow = write_freq_block(
            ws3, df_f, crow, f"Distribution by {field.replace('_', ' ').title()}")
    for ltr in ["A","B","C"]:
        ws3.column_dimensions[ltr].width = 35 if ltr == "A" else 14

    # ── Sheet 4: Keyword Analysis ──────────────────────────────────────────────
    ws4 = wb.create_sheet("Keyword Analysis")
    ws4.merge_cells("A1:E1")
    ws4["A1"].value = "Top Keywords by Text Field"
    ws4["A1"].font  = _font(True, OPXL_HEX["primary"], 13)
    kw_row = 3
    for field, label in [
        ("description",    "Description / Threat-Opportunity"),
        ("response",       "Response / Mitigation"),
        ("lessons_learned","Lessons Learned"),
        ("lesson_title",   "Lesson Titles"),
    ]:
        col = col_map.get(field)
        if not col:
            continue
        kws = top_keywords(df[col].tolist(), sw, 20)
        if not kws:
            continue
        ws4.cell(row=kw_row, column=1).value = f">> {label}"
        ws4.cell(row=kw_row, column=1).font  = _font(True, OPXL_HEX["secondary"], 11)
        kw_row += 1
        for ci, h in enumerate(["Keyword", "Frequency"], 1):
            ws4.cell(row=kw_row, column=ci).value = h
        style_header(ws4, kw_row, 2)
        kw_row += 1
        for word, freq in kws:
            ws4.cell(row=kw_row, column=1).value = word
            ws4.cell(row=kw_row, column=2).value = freq
            for ci in [1, 2]:
                ws4.cell(row=kw_row, column=ci).border = _bdr()
                if kw_row % 2 == 0:
                    ws4.cell(row=kw_row, column=ci).fill = _fill(OPXL_HEX["light"])
            kw_row += 1
        kw_row += 2
    ws4.column_dimensions["A"].width = 28
    ws4.column_dimensions["B"].width = 14

    # ── Sheet 5: Threat vs. Opportunity keywords ───────────────────────────────
    ws5 = wb.create_sheet("Threat vs Opportunity")
    ws5.merge_cells("A1:F1")
    ws5["A1"].value = "Threat vs. Opportunity — Keyword Comparison"
    ws5["A1"].font  = _font(True, OPXL_HEX["primary"], 13)
    t_texts, o_texts = split_threats_opps(df, col_map)
    t_kws = top_keywords(t_texts, sw, 15) if t_texts else []
    o_kws = top_keywords(o_texts, sw, 15) if o_texts else []
    row = 3
    ws5.cell(row=row, column=1).value = "THREAT Keywords"
    ws5.cell(row=row, column=1).font  = _font(True, OPXL_HEX["threat"], 11)
    ws5.cell(row=row, column=4).value = "OPPORTUNITY Keywords"
    ws5.cell(row=row, column=4).font  = _font(True, OPXL_HEX["opportunity"], 11)
    row += 1
    for ci, h in [(1,"Keyword"),(2,"Freq"),(4,"Keyword"),(5,"Freq")]:
        ws5.cell(row=row, column=ci).value = h
    style_header(ws5, row, 5, bg=OPXL_HEX["dark"])
    row += 1
    for i in range(max(len(t_kws), len(o_kws))):
        if i < len(t_kws):
            ws5.cell(row=row+i, column=1).value = t_kws[i][0]
            ws5.cell(row=row+i, column=2).value = t_kws[i][1]
        if i < len(o_kws):
            ws5.cell(row=row+i, column=4).value = o_kws[i][0]
            ws5.cell(row=row+i, column=5).value = o_kws[i][1]
        for ci in [1,2,4,5]:
            ws5.cell(row=row+i, column=ci).border = _bdr()
            if i % 2 == 0:
                ws5.cell(row=row+i, column=ci).fill = _fill(OPXL_HEX["light"])
    for ltr in ["A","D"]: ws5.column_dimensions[ltr].width = 28
    for ltr in ["B","E"]: ws5.column_dimensions[ltr].width = 12

    # ── Sheet 6: Stage × Category Heatmap ─────────────────────────────────────
    pivot = stage_cat_pivot(df, col_map)
    if pivot is not None and not pivot.empty:
        ws6 = wb.create_sheet("Stage x Category Matrix")
        ws6["A1"].value = "Project Stage × Category — Heatmap"
        ws6["A1"].font  = _font(True, OPXL_HEX["primary"], 13)
        pr = pivot.reset_index()
        for ci, col in enumerate(pr.columns, 1):
            ws6.cell(row=3, column=ci).value = str(col)
        style_header(ws6, 3, len(pr.columns))
        num_cols = pr.select_dtypes(include="number")
        max_val  = num_cols.max().max() if not num_cols.empty else 1
        for ri, row in pr.iterrows():
            for ci, val in enumerate(row, 1):
                cell = ws6.cell(row=ri + 4, column=ci)
                cell.value     = val
                cell.border    = _bdr()
                cell.alignment = _align("center")
                if ci > 1 and isinstance(val, (int, float)) and val > 0 and max_val:
                    ix = int(val / max_val * 200)
                    cell.fill = _fill(f"FF{format(255-ix,'02X')}F0F0")
        for col in ws6.columns:
            w = max((len(str(c.value or "")) for c in col), default=10)
            ws6.column_dimensions[col[0].column_letter].width = min(w + 4, 30)

    wb.save(out_path)
    print(f"  Excel saved  ->  {out_path}")


# =============================================================================
# POWERPOINT BUILDER
# =============================================================================

W = Inches(13.33)
H = Inches(7.5)


def _slide(prs):
    return prs.slides.add_slide(prs.slide_layouts[6])


def _bg(slide, color: RGBColor):
    f = slide.background.fill; f.solid(); f.fore_color.rgb = color


def _rect(slide, l, t, w, h, color: RGBColor):
    s = slide.shapes.add_shape(1, l, t, w, h)
    s.fill.solid(); s.fill.fore_color.rgb = color
    s.line.fill.background(); return s


def _textbox(slide, text, l, t, w, h, size=14, bold=False,
             color=None, align=PP_ALIGN.LEFT):
    color = color or RGBColor(0x1A, 0x2B, 0x3C)
    tb    = slide.shapes.add_textbox(l, t, w, h)
    tf    = tb.text_frame; tf.word_wrap = True
    p     = tf.paragraphs[0]; p.alignment = align
    run   = p.add_run(); run.text = text
    run.font.size = Pt(size); run.font.bold = bold
    run.font.color.rgb = color


def slide_header(prs, slide, title, subtitle=""):
    _bg(slide, RGBColor(0xFA, 0xFA, 0xFB))
    _rect(slide, 0, 0, W, Inches(0.85), PPTX_RGB["primary"])
    _textbox(slide, title, Inches(0.3), Inches(0.08), Inches(12.5), Inches(0.65),
             size=22, bold=True, color=RGBColor(0xFF, 0xFF, 0xFF))
    if subtitle:
        _textbox(slide, subtitle, Inches(0.3), Inches(0.72), Inches(12), Inches(0.3),
                 size=10, color=RGBColor(0xBB, 0xCC, 0xDD))


def add_image_slide(prs, path, title, subtitle=""):
    s = _slide(prs)
    slide_header(prs, s, title, subtitle)
    if path and Path(path).exists():
        s.shapes.add_picture(path, Inches(0.5), Inches(1.05),
                             width=Inches(12.3), height=Inches(6.1))


def add_table_slide(prs, df_t, title, col_widths=None):
    s = _slide(prs)
    slide_header(prs, s, title)
    rows, cols = len(df_t) + 1, len(df_t.columns)
    tbl = s.shapes.add_table(
        rows, cols, Inches(0.4), Inches(1.0),
        Inches(12.5), Inches(min(5.8, rows * 0.38))).table
    col_widths = col_widths or [Inches(12.5 / cols)] * cols
    for i, cw in enumerate(col_widths):
        tbl.columns[i].width = cw
    for ci, col in enumerate(df_t.columns):
        cell = tbl.cell(0, ci); cell.text = str(col)
        cell.fill.solid(); cell.fill.fore_color.rgb = PPTX_RGB["primary"]
        p = cell.text_frame.paragraphs[0]; p.alignment = PP_ALIGN.CENTER
        p.runs[0].font.bold = True; p.runs[0].font.size = Pt(10)
        p.runs[0].font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)
    for ri, row in df_t.iterrows():
        bg = RGBColor(0xF0, 0xF4, 0xF8) if ri % 2 == 0 else RGBColor(0xFF, 0xFF, 0xFF)
        for ci, val in enumerate(row):
            cell = tbl.cell(ri + 1, ci); cell.text = str(val)
            cell.fill.solid(); cell.fill.fore_color.rgb = bg
            p = cell.text_frame.paragraphs[0]
            p.alignment = PP_ALIGN.LEFT if ci == 0 else PP_ALIGN.CENTER
            p.runs[0].font.size = Pt(9)
            p.runs[0].font.color.rgb = RGBColor(0x1A, 0x2B, 0x3C)


def build_pptx(df: pd.DataFrame, col_map: dict, charts: dict, sw: set, out_path: str):
    prs = Presentation()
    prs.slide_width  = W
    prs.slide_height = H

    # Title slide
    s = _slide(prs)
    _bg(s, PPTX_RGB["dark"])
    _rect(s, 0, 0,           W, Inches(0.1),  PPTX_RGB["accent1"])
    _rect(s, 0, Inches(7.4), W, Inches(0.1),  PPTX_RGB["accent1"])
    _textbox(s, "PROJECT LESSONS LEARNED",
             Inches(1), Inches(2.0), Inches(11), Inches(1.1),
             size=40, bold=True, color=RGBColor(0xFF,0xFF,0xFF), align=PP_ALIGN.CENTER)
    _textbox(s, "Text Analytics & Insight Report",
             Inches(1), Inches(3.2), Inches(11), Inches(0.7),
             size=20, color=PPTX_RGB["accent2"], align=PP_ALIGN.CENTER)
    _textbox(s, f"Total Records Analysed: {len(df):,}",
             Inches(1), Inches(4.0), Inches(11), Inches(0.5),
             size=13, color=RGBColor(0x88,0xAA,0xCC), align=PP_ALIGN.CENTER)

    # Agenda
    s2 = _slide(prs)
    slide_header(prs, s2, "Agenda")
    items = ["01  Distribution by Project Type",
             "02  Distribution by Project Stage",
             "03  Distribution by Category",
             "04  Threat vs. Opportunity Split",
             "05  Top Keywords — Descriptions",
             "06  Top Keywords — Lessons Learned",
             "07  Word Cloud — Key Themes",
             "08  Stage × Category Matrix",
             "09  Key Takeaways"]
    for i, item in enumerate(items):
        y = Inches(1.1 + i * 0.58)
        _rect(s2, Inches(0.4), y + Inches(0.06), Inches(0.05), Inches(0.4), PPTX_RGB["accent1"])
        _textbox(s2, item, Inches(0.6), y, Inches(11), Inches(0.52), size=14)

    # Chart slides
    for key, title, subtitle in [
        ("type_bar",       "Distribution by Project Type",     "Lessons by project type"),
        ("stage_bar",      "Distribution by Project Stage",    "Which stages generate the most lessons"),
        ("category_bar",   "Distribution by Category",         "Risk/opportunity category breakdown"),
        ("threat_opp_pie", "Threat vs. Opportunity Split",     "Proportion of threats and opportunities"),
        ("keyword_all",    "Top Keywords — All Text",          "Most frequent terms across all text fields"),
        ("keyword_lessons","Top Keywords — Lessons Learned",   "Key terms from lessons learned text"),
        ("threat_kw",      "Top Keywords — Threats",           "Dominant threat-related themes"),
        ("opp_kw",         "Top Keywords — Opportunities",     "Dominant opportunity-related themes"),
        ("wordcloud",      "Word Cloud — Combined Text",       "Visual map of prominent themes"),
        ("stage_cat_stack","Stage × Category Breakdown",       "How categories distribute across stages"),
    ]:
        path = charts.get(key)
        if path and Path(path).exists():
            add_image_slide(prs, path, title, subtitle)

    # Frequency table slides
    for field in ["project_type", "project_stage", "new_category"]:
        col = col_map.get(field)
        if not col:
            continue
        df_f = freq_table(df, col, field.replace("_", " ").title())
        add_table_slide(prs, df_f.head(10),
                        f"Top 10 — {field.replace('_', ' ').title()}",
                        col_widths=[Inches(7), Inches(2.5), Inches(3)])

    # Stage × Category table
    pivot = stage_cat_pivot(df, col_map)
    if pivot is not None and not pivot.empty:
        pr  = pivot.reset_index()
        nc  = len(pr.columns)
        cws = [Inches(2.5)] + [Inches(10 / max(nc - 1, 1))] * (nc - 1)
        add_table_slide(prs, pr.head(10), "Stage × Category Matrix", col_widths=cws)

    # Key takeaways slide
    s_kt = _slide(prs)
    _bg(s_kt, PPTX_RGB["dark"])
    _rect(s_kt, 0, 0, W, Inches(1.0), PPTX_RGB["accent1"])
    _textbox(s_kt, "Key Takeaways",
             Inches(0.3), Inches(0.1), Inches(12), Inches(0.8),
             size=28, bold=True, color=RGBColor(0xFF,0xFF,0xFF))

    tc = col_map.get("project_type"); sc = col_map.get("project_stage")
    cc = col_map.get("new_category") or col_map.get("original_category")
    dc = col_map.get("description"); lc = col_map.get("lessons_learned")
    t_texts, o_texts = split_threats_opps(df, col_map)

    takeaways = []
    if tc:
        top = df[tc].value_counts()
        takeaways.append(f"Most lessons come from '{top.idxmax()}' projects ({top.iloc[0]} records)")
    if sc:
        top = df[sc].value_counts()
        takeaways.append(f"'{top.idxmax()}' is the most represented project stage")
    if cc:
        top = df[cc].value_counts()
        takeaways.append(f"'{top.idxmax()}' is the leading lesson category")
    tn, on = len(t_texts), len(o_texts)
    if tn + on > 0:
        takeaways.append(f"Dataset contains {tn} threat and {on} opportunity entries "
                         f"({tn/(tn+on)*100:.0f}% threats)")
    if dc:
        kws = top_keywords(df[dc].tolist(), sw, 5)
        if kws:
            takeaways.append("Top description keywords: " + ", ".join(k for k,_ in kws))
    if lc:
        kws = top_keywords(df[lc].tolist(), sw, 5)
        if kws:
            takeaways.append("Top lessons learned keywords: " + ", ".join(k for k,_ in kws))
    takeaways.append("Full frequency tables, keyword lists & heatmaps available in the Excel workbook")

    for i, ta in enumerate(takeaways[:8]):
        y = Inches(1.15 + i * 0.75)
        _rect(s_kt, Inches(0.4), y + Inches(0.1), Inches(0.35), Inches(0.38), PPTX_RGB["accent2"])
        _textbox(s_kt, str(i+1), Inches(0.43), y+Inches(0.08), Inches(0.3), Inches(0.38),
                 size=12, bold=True, color=PPTX_RGB["dark"], align=PP_ALIGN.CENTER)
        _textbox(s_kt, ta, Inches(0.9), y, Inches(11.8), Inches(0.65),
                 size=13, color=RGBColor(0xEC,0xE2,0xD0))

    prs.save(out_path)
    print(f"  PowerPoint saved  ->  {out_path}")


# =============================================================================
# MAIN
# =============================================================================

def main():
    parser = argparse.ArgumentParser(
        description="Analyse project lessons-learned Excel data → insights workbook + PowerPoint")
    parser.add_argument("input_file", help="Input .xlsx file")
    parser.add_argument("--sheet",   default=None, help="Sheet name (default: first sheet)")
    parser.add_argument("--out-dir", default=".",  help="Output directory (default: .)")
    args = parser.parse_args()

    input_path = Path(args.input_file)
    if not input_path.exists():
        sys.exit(f"File not found: {input_path}")

    out_dir = Path(args.out_dir)
    out_dir.mkdir(parents=True, exist_ok=True)
    stem = input_path.stem

    print(f"\nLoading:  {input_path}")
    df      = load_data(str(input_path), args.sheet)
    col_map = map_columns(df)
    sw      = build_stopwords()

    print(f"  Rows: {len(df)}  |  Columns: {len(df.columns)}")
    print(f"  Recognised fields: {list(col_map.keys())}\n")

    print("Generating charts …")
    charts = generate_charts(df, col_map, sw)
    print()

    xl_out   = out_dir / f"{stem}_insights.xlsx"
    pptx_out = out_dir / f"{stem}_insights.pptx"

    print("Building Excel workbook …")
    build_excel(df, col_map, charts, sw, str(xl_out))

    print("Building PowerPoint presentation …")
    build_pptx(df, col_map, charts, sw, str(pptx_out))

    print(f"\nDone!  Outputs in: {out_dir.resolve()}")
    print(f"  {xl_out.name}")
    print(f"  {pptx_out.name}\n")


if __name__ == "__main__":
    main()
