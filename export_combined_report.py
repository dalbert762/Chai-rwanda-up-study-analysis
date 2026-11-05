"""
One HTML report holding all three UP-study analyses, as three tabbed pages.

    python3 fetch_kobo_export.py          # refresh all three exports
    python3 export_combined_report.py     # -> Combined_analysis_report.html

  Page 1  Medicine assessment            analyze_medicine_assessment
  Page 2  Prescription assessment        analyze_kobo_prescription
  Page 3  Medicine assessment for CHWs   analyze_medicine_assessment_chw

Imports the three analysis modules and renders their existing `SECTIONS`
rather than recomputing anything, so the HTML can never drift from the three
Excel workbooks. All three expose the same Table interface
(.title/.headers/.rows/.note/.total/.chart/.pct), which is what makes one
renderer work for all of them.

Two rules shape what lands on a page:

  * every table each analysis script produces is rendered, except the
    sections in `DROP_SECTIONS` (README and Findings - provenance and written
    commentary rather than indicators).
  * every indicator appears ONE way: a table that can be charted is rendered
    as its chart alone (each bar carries its own value label), everything
    else as its table. Never both.

So a chart on these pages stands in for its table; the workbook has that
table in full, as it has everything else.

The three pages are one self-contained file, not three: a single document
opens from Box or an email attachment with no broken links between parts, and
the tab bar swaps which page is displayed. With JavaScript off, and when
printing, all three pages render in sequence instead.

Writes two files:
  Combined_analysis_report.html   full standalone document, opens from disk
  <scratch>/combined_artifact.html  same content without the <html> wrapper,
                                    for publishing as an Artifact

The page is black, in one theme, whatever the reader's OS is set to, and the
type is Trebuchet MS throughout - a system font on Windows and macOS, so the
document now fetches nothing and is genuinely self-contained.

Charts are inline SVG - no libraries, so the file works offline. Colours are
the dataviz reference palette's categorical slots 1 (blue, prescription),
2 (orange, stock) and 3 (aqua, CHW), at their DARK steps (#3987e5 / #d95926 /
#199e70). Those steps are validated against the palette's dark chart surface
#1a1a19, which is why cards and charts keep that surface and only the page
plane behind them is pure black - dropping the marks onto #000 would put them
outside the validated result.

openpyxl + the standard library only, like the rest of this project.
"""

import copy
import datetime as dt
import html
import os
import sys

import analyze_kobo_prescription as P
import analyze_medicine_assessment as M
import analyze_medicine_assessment_chw as C

OUT_HTML = "Combined_analysis_report.html"
OUT_FRAGMENT = os.environ.get("COMBINED_FRAGMENT_PATH", "")

BLUE, BLUE_D = "#2a78d6", "#3987e5"          # slot 1 - prescription
ORANGE, ORANGE_D = "#eb6834", "#d95926"      # slot 2 - stock
AQUA, AQUA_D = "#1baf7a", "#199e70"          # slot 3 - CHW


# ---------------------------------------------------------------
# helpers
# ---------------------------------------------------------------
def esc(v):
    return html.escape("" if v is None else str(v), quote=True)


def is_num(v):
    return isinstance(v, (int, float)) and not isinstance(v, bool)


def fmt(v, as_pct=False):
    """Cell text. Percentages to one decimal; counts with thousands separators."""
    if v is None or v == "":
        return ""
    if isinstance(v, (dt.datetime, dt.date)):
        return v.strftime("%Y-%m-%d")
    if is_num(v):
        if as_pct:
            return f"{v * 100:.1f}%"
        if float(v).is_integer() and abs(v) < 1e15:
            return f"{int(v):,}"
        return f"{v:,.2f}"
    return str(v)


def nice_max(hi):
    """A rounded axis maximum, so every gridline names a real number."""
    if hi <= 0:
        return 1, [0, 1]
    mag = 10 ** (len(str(int(hi))) - 1)
    for step in (1, 2, 2.5, 5, 10):
        top = step * mag
        if top >= hi:
            ticks = [i * top / 4 for i in range(5)]
            return top, ticks
    top = 10 * mag
    return top, [i * top / 4 for i in range(5)]


def bar_h(x0, x1, y, h, r=4):
    """Horizontal bar: rounded at the data end only, square on the baseline."""
    r = max(0, min(r, x1 - x0, h / 2))
    return (f"M{x0:.1f},{y:.1f} H{x1 - r:.1f} Q{x1:.1f},{y:.1f} {x1:.1f},{y + r:.1f} "
            f"V{y + h - r:.1f} Q{x1:.1f},{y + h:.1f} {x1 - r:.1f},{y + h:.1f} "
            f"H{x0:.1f} Z")


def bar_v(x, w, y_base, y_top, r=4):
    r = max(0, min(r, w / 2, y_base - y_top))
    return (f"M{x:.1f},{y_base:.1f} V{y_top + r:.1f} Q{x:.1f},{y_top:.1f} "
            f"{x + r:.1f},{y_top:.1f} H{x + w - r:.1f} "
            f"Q{x + w:.1f},{y_top:.1f} {x + w:.1f},{y_top + r:.1f} "
            f"V{y_base:.1f} Z")


def trunc(s, n):
    s = str(s)
    return s if len(s) <= n else s[: n - 1] + "…"


def compact(v):
    """25,584,502 -> 25.6M. Axis and bar labels only; tooltips stay exact."""
    a = abs(v)
    if a >= 1e9:
        return f"{v / 1e9:,.1f}B"
    if a >= 1e6:
        return f"{v / 1e6:,.1f}M"
    if a >= 1e4:
        return f"{v / 1e3:,.0f}k"
    return fmt(v)


# ---------------------------------------------------------------
# inline SVG charts
# ---------------------------------------------------------------
def svg_chart(table, value_label, pct_axis):
    """One single-series bar chart. Horizontal when labels are long."""
    pairs = table.chart_data()
    if not pairs:
        return ""
    kind = table.chart[0]
    pairs = [(str(a), float(b)) for a, b in pairs]
    hi = max(v for _, v in pairs)
    top, ticks = nice_max(hi)

    # Above ~10,000 the full number stops being readable at a glance and starts
    # being a wide label competing with the bar next to it: 25,584,502 becomes
    # 25.6M on the axis and on the bar, and the exact figure stays one hover
    # away in the tooltip. Below that, nothing is abbreviated.
    abbrev = not pct_axis and top >= 1e4

    def tick_text(t):
        if pct_axis:
            return f"{t * 100:.0f}%"
        if abbrev:
            return compact(t)
        return f"{t:,.0f}" if float(t).is_integer() else f"{t:,.1f}"

    def val_text(v):
        if pct_axis:
            return f"{v * 100:.1f}%"
        return compact(v) if abbrev else fmt(v)

    if kind == "h":
        row, gap = 23, 6
        # pad_r must clear the value label sitting past the bar end: the widest
        # is a grouped RWF total, ~10 glyphs at 11.5px.
        lab_w, pad_r, pad_t = 224, 104, 26
        w = 860
        h = pad_t + len(pairs) * (row + gap) + 26
        plot_x0, plot_x1 = lab_w, w - pad_r
        span = plot_x1 - plot_x0
        out = [f'<svg class="chart" viewBox="0 0 {w} {h}" role="img" '
               f'aria-label="{esc(table.title)}">']
        for t in ticks:
            x = plot_x0 + span * (t / top)
            out.append(f'<line x1="{x:.1f}" y1="{pad_t - 8}" x2="{x:.1f}" '
                       f'y2="{h - 26}" class="grid" />')
            out.append(f'<text x="{x:.1f}" y="{h - 10}" class="tick" '
                       f'text-anchor="middle">{esc(tick_text(t))}</text>')
        for i, (lab, v) in enumerate(pairs):
            y = pad_t + i * (row + gap)
            bw = span * (v / top)
            # Each row is one group behind a full-width transparent target, so
            # the pointer picks the row up anywhere along it - a 3px bar for a
            # near-zero value is otherwise almost impossible to hover - and the
            # tooltip carries the exact figure the label may have abbreviated.
            out.append(f'<g class="row"><rect class="hit" x="0" '
                       f'y="{y - gap / 2:.1f}" width="{w}" height="{row + gap}">'
                       f'<title>{esc(lab)}: {esc(fmt(v, pct_axis))}</title></rect>')
            out.append(f'<text x="{lab_w - 10}" y="{y + row * 0.72:.1f}" '
                       f'class="blab" text-anchor="end">{esc(trunc(lab, 30))}</text>')
            out.append(f'<path d="{bar_h(plot_x0, plot_x0 + max(bw, 0.8), y, row)}" '
                       f'class="bar" />')
            out.append(f'<text x="{plot_x0 + bw + 8:.1f}" y="{y + row * 0.72:.1f}" '
                       f'class="vlab">{esc(val_text(v))}</text></g>')
        out.append(f'<line x1="{plot_x0}" y1="{pad_t - 8}" x2="{plot_x0}" '
                   f'y2="{h - 26}" class="axis" />')
        out.append("</svg>")
        return "".join(out)

    # vertical
    n = len(pairs)
    w = 860
    lab_chars = 16
    # x-labels are rotated -35 deg with text-anchor=end, so they descend
    # sin(35) * width below their anchor. Reserve that or they run out of the
    # viewBox - a 16-char label at 12px is ~109px wide, ~63px of descent.
    pad_l, pad_r, pad_t = 62, 16, 24
    pad_b = 16 + int(lab_chars * 6.8 * 0.574) + 14
    plot_h = 210
    h = pad_t + plot_h + pad_b
    span = w - pad_l - pad_r
    slot = span / max(n, 1)
    bw = min(46, slot * 0.62)
    # With many bars neither a value on every bar nor a label on every bar
    # fits; thin the labels rather than overlap them.
    show_vals = n <= 24
    lab_every = 1 if n <= 30 else (2 if n <= 60 else 3)
    out = [f'<svg class="chart" viewBox="0 0 {w} {h}" role="img" '
           f'aria-label="{esc(table.title)}">']
    for t in ticks:
        y = pad_t + plot_h - plot_h * (t / top)
        out.append(f'<line x1="{pad_l}" y1="{y:.1f}" x2="{w - pad_r}" '
                   f'y2="{y:.1f}" class="grid" />')
        out.append(f'<text x="{pad_l - 8}" y="{y + 4:.1f}" class="tick" '
                   f'text-anchor="end">{esc(tick_text(t))}</text>')
    for i, (lab, v) in enumerate(pairs):
        cx = pad_l + slot * i + slot / 2
        x = cx - bw / 2
        bh = plot_h * (v / top)
        y_top = pad_t + plot_h - bh
        # Full-height column target, for the same reason as the rows above -
        # and it is the only way to read a bar whose label was thinned out.
        out.append(f'<g class="row"><rect class="hit" x="{pad_l + slot * i:.1f}" '
                   f'y="{pad_t - 8}" width="{slot:.1f}" height="{plot_h + 8}">'
                   f'<title>{esc(lab)}: {esc(fmt(v, pct_axis))}</title></rect>')
        out.append(f'<path d="{bar_v(x, bw, pad_t + plot_h, y_top)}" class="bar" />')
        if show_vals:
            out.append(f'<text x="{cx:.1f}" y="{y_top - 6:.1f}" class="vlab" '
                       f'text-anchor="middle">{esc(val_text(v))}</text>')
        if i % lab_every == 0:
            ly = pad_t + plot_h + 16
            out.append(f'<text x="{cx:.1f}" y="{ly}" class="blab" '
                       f'text-anchor="end" transform="rotate(-35 {cx:.1f} '
                       f'{ly})">{esc(trunc(lab, lab_chars))}</text>')
        out.append("</g>")
    out.append(f'<line x1="{pad_l}" y1="{pad_t + plot_h}" x2="{w - pad_r}" '
               f'y2="{pad_t + plot_h}" class="axis" />')
    out.append("</svg>")
    return "".join(out)


# ---------------------------------------------------------------
# tables
# ---------------------------------------------------------------
def render_item(t, idx):
    """One indicator, shown ONE way.

    A table that can be charted is shown only as its chart - the chart carries
    a direct value label on every bar, so the numbers are not lost, and showing
    both says the same thing twice at twice the height. Everything else is
    shown as its table. The full table for every chart is in the workbook.
    """
    if t.chart and t.chart_data():
        return render_chart_only(t, idx)
    return render_table(t, idx)


def render_chart_only(t, idx):
    vc = t.chart[2]
    shown, total = len(t.chart_data()), len(t.rows)
    # A capped chart must say so: "top 20" read as "all 130" is a wrong
    # conclusion, not a cosmetic one.
    cap = (f"Chart shows the top {shown:,} of {total:,} rows; the workbook has "
           f"all of them." if shown < total else "")
    note = " ".join(x for x in (t.note or "", cap) if x)
    # The chart replaced its table, so the heading has to say WHAT is plotted -
    # "By province" alone does not tell a reader the bars are availability.
    measure = str(t.headers[vc]).strip()
    return (f'<section class="tbl" id="t{idx}">'
            f'<h4>{esc(t.title)}<span class="measure">{esc(measure)}</span></h4>'
            f'{svg_chart(t, str(t.headers[vc]), vc in t.pct)}'
            + (f'<p class="note">{esc(note)}</p>' if note else "")
            + "</section>")


def render_table(t, idx):
    rows = list(t.rows)
    body = []
    for r in rows:
        tds = []
        for ci, v in enumerate(r):
            cls = "num" if is_num(v) else ""
            tds.append(f'<td class="{cls}">{esc(fmt(v, ci in t.pct))}</td>')
        body.append("<tr>" + "".join(tds) + "</tr>")
    if t.total:
        tds = []
        for ci, v in enumerate(t.total):
            cls = "num" if is_num(v) else ""
            tds.append(f'<td class="{cls}">{esc(fmt(v, ci in t.pct))}</td>')
        body.append('<tr class="total">' + "".join(tds) + "</tr>")

    head = "".join(f"<th>{esc(h)}</th>" for h in t.headers)
    tall = " tall" if len(rows) > 18 else ""
    note = (f'<p class="note">{esc(t.note)}</p>' if t.note else "")
    count = f'<span class="rowcount">{len(rows):,} rows</span>' if len(rows) > 12 else ""
    return (f'<section class="tbl" id="t{idx}">'
            f'<h4>{esc(t.title)}{count}</h4>'
            f'<div class="scroll{tall}"><table>'
            f"<thead><tr>{head}</tr></thead><tbody>{''.join(body)}</tbody>"
            f"</table></div>{note}</section>")


# ---------------------------------------------------------------
# what the dashboard leaves out
# ---------------------------------------------------------------
# Every table each analysis script produces is rendered, EXCEPT the sections
# named here. A blocklist, not an allowlist: a table added to an analysis
# script shows up on these pages on the next build without being registered
# anywhere, which is what keeps the HTML and the workbooks in step.
#
# README is workbook provenance - refresh commands and unit-of-analysis notes
# that belong in the file they describe. Findings is written commentary rather
# than an indicator. Data collectors profiles the enumerators rather than the
# medicines, which is a supervision question, not a supply one.
DROP_SECTIONS = {"README", "Findings", "Data collectors",
                 "Facility scorecard"}

# Tables cut wherever they appear, for the same reason: they break results
# down BY enumerator. "Collector feedback on the supply chain" is not one of
# these - it is what collectors reported about supply, not a read on them.
DROP_TABLES = {"Data collectors", "By data collector",
               # fieldwork logistics rather than medicine supply: when forms
               # arrived, and which version of the form they were filled on
               "Submissions by day", "Form versions in use",
               # these count SUBMISSIONS, not facilities - a facility visited
               # twice counts twice, so they overstate reach. The facility
               # counts are in "Coverage by province" / "by facility type".
               "Facility visits by province", "Facility visits by district",
               "Facility visits by facility type",
               # export mechanics, not findings
               "Constant columns (no analytical value in this export)",
               "Form version by submission"}

# Individual rows cut from a table that is otherwise kept, keyed by table
# title and matched on the row's first cell.
DROP_ROWS = {"Dataset": {"Facility visits (submissions)"}}


def curate(page_id, sections):
    """Drop the non-indicator sections, tables and rows; keep everything else.

    Tables are copied before their rows are touched - the module's own objects
    are what its Excel workbook is built from, and this view must not edit them.
    """
    out = []
    for name, tables in sections:
        if name in DROP_SECTIONS:
            continue
        kept = []
        for t in tables:
            if t.title in DROP_TABLES:
                continue
            cut = DROP_ROWS.get(t.title)
            if cut:
                t = copy.copy(t)
                t.rows = [r for r in t.rows if str(r[0]).strip() not in cut]
            kept.append(t)
        if kept:
            out.append((name, kept))
    return out


def anchor_for(page_id, section_name):
    return f"{page_id}-{section_name.lower().replace(' ', '-').replace('&', 'and')}"


def render_toc(page_id, sections):
    items = "".join(f'<li><a href="#{esc(anchor_for(page_id, n))}">{esc(n)}</a></li>'
                    for n, _ in sections)
    return f'<nav class="toc"><p class="tochead">On this page</p><ul>{items}</ul></nav>'


def render_page(page_id, kicker, title, blurb, stats, panel, sections, start_idx,
                active=False):
    """One tab's worth of document. `panel` is the page's caveat block, which is
    the part a reader must see before quoting any figure from that page."""
    out = [f'<div class="page{" on" if active else ""}" id="{page_id}" '
           f'role="tabpanel" aria-labelledby="tab-{page_id}">',
           f'<header class="parthead"><p class="kicker">{esc(kicker)}</p>',
           f'<h2>{esc(title)}</h2><p class="lede">{blurb}</p></header>',
           f'<div class="stats">{stats}</div>' if stats else "",
           panel or "",
           render_toc(page_id, sections)]
    idx = start_idx
    for name, tables in sections:
        out.append(f'<section class="sec" id="{esc(anchor_for(page_id, name))}">'
                   f"<h3>{esc(name)}</h3>")
        for t in tables:
            out.append(render_item(t, idx))
            idx += 1
        out.append("</section>")
    out.append("</div>")
    return "".join(out), idx


# ---------------------------------------------------------------
# gather both studies
# ---------------------------------------------------------------
P._count_themes()
PRESC_FULL = [(name, fn()) for name, fn in P.SECTIONS]
STOCK_SECTIONS = curate("stock", M.SECTIONS)
PRESC_SECTIONS = curate("presc", PRESC_FULL)
CHW_SECTIONS = curate("chw", C.SECTIONS)

# --- headline figures, read from the modules so they cannot drift
p_rate = P.ITEM_NA / (P.ITEM_AV + P.ITEM_NA) if (P.ITEM_AV + P.ITEM_NA) else 0
# Method 3 (adopted 2026-10-05): a medicine assessed at two strengths is one
# molecule per facility, available if either strength was on the shelf. The
# headline is the clinical question; the availability tables further down this
# page stay at item level, because they describe the instrument's own rows.
sys.path.insert(0, os.path.join(os.path.dirname(os.path.abspath(__file__)),
                                "analytics"))
import molecules as MOL  # noqa: E402

# DEFINITION (study): the denominator is every molecule the EML authorises at
# that level of care, so a medicine the facility is supposed to manage but
# does not manage at all counts as UNAVAILABLE.
M3 = MOL.eml_base(M.MEDROWS)
m_answered = len(M3)
m_unavail = sum(1 for r in M3 if M.no(r[M.AVAIL]))
m_gap = sum(1 for r in M3 if r["_gap"])
m_rate = m_unavail / m_answered if m_answered else 0
m_items = len(M.answered_all)
m_fulfil = M.del_all / M.req_all if M.req_all else 0

P_FACILITIES = len({str(r[P.FACIL]) for r in P.sub})


def stat(value, label, sub=""):
    return (f'<div class="stat"><p class="v">{value}</p>'
            f'<p class="l">{esc(label)}</p>'
            + (f'<p class="s">{esc(sub)}</p>' if sub else "") + "</div>")


HERO_STATS = "".join([
    stat(f"{len(M.FACILITIES)}", "facilities visited", "medicine assessment"),
    stat(f"{P_FACILITIES}", "facilities visited", "prescription assessment"),
    stat(f"{len(C.FACILITIES)}", "health centres", f"{len(C.CHWS)} CHWs assessed"),
    stat(f"{M.loss_all / 1e6:,.1f}M", "RWF value of stock-outs", "medicine assessment"),
])

# --- the tab bar. Order is the reading order: facility stock first, since the
# other two are read against it.
PAGES = [("stock", "Medicine assessment"),
         ("presc", "Prescription"),
         ("chw", "Medicine assessment for CHWs")]
TABS = "".join(
    f'<button class="tab{" on" if i == 0 else ""}" id="tab-{pid}" data-page="{pid}" '
    f'role="tab" aria-controls="{pid}" aria-selected="{"true" if i == 0 else "false"}">'
    f"{esc(label)}</button>" for i, (pid, label) in enumerate(PAGES))

# meter widths: scaled to the larger of the two rates so the bars compare
meter_hi = max(p_rate, m_rate) or 1

# The comparison panel lives on the stock page - it is the finding that decides
# which of the two availability figures may be quoted, and the stock page is the
# one that survives the comparison.
STOCK_PANEL = f"""
<div class="thesis">
  <div class="versus">
    <div class="vs s">
      <p class="who">Medicine stock assessment &mdash; this page</p>
      <p class="big">{m_rate * 100:.1f}%</p>
      <p class="cap">of {m_answered:,} molecules the EML authorises at that
      level of care could not be dispensed &mdash; {m_unavail:,} observations,
      of which {m_gap:,} are medicines the facility does not manage at all and
      {m_unavail - m_gap:,} were managed but empty. Strengths of the same
      medicine are counted once per facility.</p>
      <div class="meter"><i style="width:{m_rate / meter_hi * 100:.1f}%"></i></div>
    </div>
    <div class="vs p">
      <p class="who">Prescription assessment</p>
      <p class="big">{p_rate * 100:.2f}%</p>
      <p class="cap">of {P.N_ITEM:,} prescribed items recorded as not available
      &mdash; {P.ITEM_NA} items in total.</p>
      <div class="meter"><i style="width:{p_rate / meter_hi * 100:.1f}%"></i></div>
    </div>
  </div>
</div>
"""

stock_html, next_idx = render_page(
    "stock", "Page one", "Medicine assessment",
    f"A fixed basket of essential medicines checked against what each facility "
    f"actually holds. {M.N_MED:,} facility-by-medicine observations at "
    f"{len(M.FACILITIES)} facilities; the unit of analysis is the medicine row, "
    f"never the submission.",
    "".join([
        stat(f"{len(M.FACILITIES)}", "facilities visited",
             f"in {len({str(r[M.DIST]) for r in M.SUB})} districts"),
        stat(f"{M.N_MED:,}", "medicine observations", "facility x medicine"),
        stat(f"{(1 - m_rate) * 100:.1f}%", "available on visit day",
             f"{m_unavail:,} absent of {m_answered:,} authorised"),
        stat(f"{m_fulfil * 100:.1f}%", "requisition fulfilment",
             "delivered against requested"),
    ]),
    STOCK_PANEL, STOCK_SECTIONS, 0, active=True)

presc_html, next_idx = render_page(
    "presc", "Page two", "Prescription assessment",
    f"What was prescribed, and whether the facility could dispense it. "
    f"{P.N_PRES:,} prescriptions carrying {P.N_ITEM:,} medicine items, "
    f"collected at {P_FACILITIES} facilities in "
    f"{len({str(r[P.DISTR]) for r in P.sub})} districts.",
    "".join([
        stat(f"{P_FACILITIES}", "facilities visited",
             f"in {len({str(r[P.DISTR]) for r in P.sub})} districts"),
        stat(f"{P.N_PRES:,}", "prescriptions", f"{P.N_ITEM:,} medicine items"),
        stat(f"{P.N_ITEM / P.N_PRES:.2f}", "medicines per encounter",
             "WHO/INRUD indicator"),
        stat(f"{p_rate * 100:.2f}%", "items not available",
             f"{P.ITEM_NA} of {P.N_ITEM:,} items"),
    ]),
    None, PRESC_SECTIONS, next_idx)

chw_html, _ = render_page(
    "chw", "Page three", "Medicine assessment for CHWs",
    f"The same availability and consumption questions, asked of the community "
    f"health worker rather than the facility. {C.N_MED:,} CHW-by-medicine "
    f"observations covering {len(C.CHWS)} CHWs at {len(C.FACILITIES)} parent "
    f"health centres, over the 13-item CHW package.",
    "".join([
        stat(f"{len(C.CHWS)}", "CHWs assessed", f"{len(C.FACILITIES)} health centres"),
        stat(f"{C.N_MED:,}", "CHW x medicine rows", f"{C.N_MANAGED:,} managed"),
        stat(f"{C.N_AVAIL / C.N_MANAGED * 100:.1f}%", "available on visit day",
             f"{C.N_OUT} stocked out of {C.N_MANAGED:,}"),
        stat(f"{C.AMC_ZERO}", "rows with no consumption",
             "Jun-Aug, of the managed rows"),
    ]),
    None, CHW_SECTIONS, next_idx)

CSS = """
<style>
/* One theme, dark, unconditionally - the page is black whatever the reader's
   OS is set to. The surfaces stay at the dataviz reference palette's dark
   chart surface (#1a1a19): the three series steps below were validated
   against THAT surface, so lifting cards off a pure-black plane is what keeps
   the documented contrast result true. */
:root{
  color-scheme:dark;
  --plane:#000000; --surface:#1a1a19; --raise:#222221;
  --ink:#ffffff; --ink2:#c3c2b7; --muted:#898781;
  --grid:#2c2c2a; --axis:#383835; --rule:rgba(255,255,255,.14);
  --presc:#3987e5; --stock:#d95926; --chw:#199e70;
  --accent:var(--stock);
  --good:#0ca30c; --warn:#fab219; --crit:#d03b3b;
  --band:#16202b;
  --font:"Trebuchet MS","Lucida Grande","Lucida Sans Unicode","Lucida Sans",
    Tahoma,sans-serif;
}
*{box-sizing:border-box}
body{
  margin:0; background:var(--plane); color:var(--ink);
  font-family:var(--font);
  font-size:16px; line-height:1.55; -webkit-text-size-adjust:100%;
}
.wrap{max-width:1180px; margin:0 auto; padding-inline:20px; padding-block:0 72px}
h1,h2,h3,h4{font-family:var(--font); text-wrap:balance; margin:0}
/* Trebuchet has no monospace companion, so the figures that used to rely on a
   mono face hold their columns with tabular numerals instead. */
code,.mono,.rowcount,.tick,.vlab{font-family:var(--font);
  font-variant-numeric:tabular-nums}

/* masthead */
.mast{border-bottom:1px solid var(--rule); padding-block:44px 30px; margin-bottom:34px}
.eyebrow{font-size:12.5px; letter-spacing:.14em; text-transform:uppercase;
  color:var(--muted); margin:0 0 14px; font-weight:600}
.mast h1{font-size:clamp(31px,5.2vw,51px); line-height:1.08; letter-spacing:-.015em;
  max-width:19ch}

/* stat row */
.stats{display:grid; grid-template-columns:repeat(auto-fit,minmax(180px,1fr));
  gap:1px; background:var(--rule); border:1px solid var(--rule);
  border-radius:3px; overflow:hidden; margin:30px 0 0}
.stat{background:var(--surface); padding:18px 18px 16px}
.stat .v{margin:0; font-size:28px; font-weight:600; letter-spacing:-.02em;
  font-family:var(--font)}
.stat .l{margin:3px 0 0; font-size:14px; color:var(--ink2)}
.stat .s{margin:2px 0 0; font-size:12.5px; color:var(--muted)}

/* the thesis panel */
.thesis{margin:38px 0 0; border:1px solid var(--rule); border-radius:3px;
  background:var(--surface); overflow:hidden}
/* the panel is now the two number cards alone - no heading above them to
   need a rule between */
.versus{display:grid; grid-template-columns:1fr 1fr; gap:1px;
  background:var(--rule)}
.vs{background:var(--surface); padding:20px 22px}
.vs .who{font-size:12.5px; letter-spacing:.1em; text-transform:uppercase;
  font-weight:600; margin:0 0 10px}
.vs.p .who{color:var(--presc)} .vs.s .who{color:var(--stock)}
.vs .big{font-family:var(--font); font-size:45px;
  font-weight:600; margin:0; line-height:1; letter-spacing:-.025em}
.vs.p .big{color:var(--presc)} .vs.s .big{color:var(--stock)}
.vs .cap{margin:8px 0 0; font-size:14px; color:var(--ink2)}
.vs .meter{height:8px; background:var(--grid); border-radius:99px;
  margin:14px 0 0; overflow:hidden}
.vs .meter i{display:block; height:100%; border-radius:99px}
.vs.p .meter i{background:var(--presc)} .vs.s .meter i{background:var(--stock)}
/* tab bar */
.tabs{display:flex; flex-wrap:wrap; gap:2px; margin:34px 0 0;
  border-bottom:1px solid var(--rule)}
.tab{font:inherit; font-size:14.5px; font-weight:500; color:var(--ink2);
  background:none; border:0; border-bottom:2px solid transparent;
  padding:11px 16px; margin-bottom:-1px; cursor:pointer; white-space:nowrap}
.tab:hover{color:var(--ink)}
.tab.on{color:var(--ink); font-weight:600}
.tab[data-page="stock"].on{border-bottom-color:var(--stock)}
.tab[data-page="presc"].on{border-bottom-color:var(--presc)}
.tab[data-page="chw"].on{border-bottom-color:var(--chw)}
.tab:focus-visible{outline:2px solid var(--accent); outline-offset:-2px}

/* contents */
.toc{display:block; margin:34px 0 0; padding:20px 22px;
  border:1px solid var(--rule); border-radius:3px; background:var(--surface)}
.tochead{margin:0 0 10px; font-size:12.5px; letter-spacing:.1em;
  text-transform:uppercase; font-weight:600; color:var(--accent)}
.toc ul{margin:0; padding:0; list-style:none; columns:3; column-gap:26px}
.toc li{margin:0 0 5px; break-inside:avoid}
.toc a{color:var(--ink2); text-decoration:none; font-size:14.5px;
  border-bottom:1px solid transparent}
.toc a:hover{color:var(--ink); border-bottom-color:var(--axis)}

/* pages */
.js .page{display:none}
.js .page.on{display:block}
.page{margin-top:44px; scroll-margin-top:16px}
.parthead{border-top:2px solid var(--accent); padding-top:20px}
#presc{--accent:var(--presc)} #stock{--accent:var(--stock)}
#chw{--accent:var(--chw)}
.kicker{margin:0 0 6px; font-size:12.5px; letter-spacing:.14em;
  text-transform:uppercase; font-weight:600; color:var(--accent)}
.parthead h2{font-size:clamp(24px,3.4vw,33px); letter-spacing:-.015em}
.lede{color:var(--ink2); max-width:70ch; margin:10px 0 0}
.sec{margin-top:44px; scroll-margin-top:16px}
.sec > h3{font-size:20px; letter-spacing:-.01em; padding-bottom:7px;
  border-bottom:1px solid var(--rule)}

/* tables */
.tbl{margin-top:26px}
.tbl h4{font-size:15.5px; font-weight:600; margin:0 0 9px;
  display:flex; justify-content:space-between; align-items:baseline; gap:14px}
.rowcount{font-size:12px; color:var(--muted); font-weight:400; white-space:nowrap}
.measure{font-size:12px; color:var(--accent); font-weight:600; white-space:nowrap;
  letter-spacing:.04em; text-transform:uppercase}
.scroll{overflow-x:auto; border:1px solid var(--rule); border-radius:3px;
  background:var(--surface)}
.scroll.tall{max-height:460px; overflow-y:auto}
table{border-collapse:collapse; width:100%; font-size:13.5px}
thead th{position:sticky; top:0; background:var(--raise); text-align:left;
  font-weight:600; font-size:12.5px; color:var(--ink2); padding:8px 11px;
  border-bottom:1px solid var(--axis); white-space:nowrap; z-index:1}
td{padding:6px 11px; border-bottom:1px solid var(--grid);
  vertical-align:top; color:var(--ink2)}
td.num{text-align:right; font-variant-numeric:tabular-nums; color:var(--ink);
  white-space:nowrap}
tbody tr:last-child td{border-bottom:0}
tr.total td{font-weight:600; color:var(--ink); background:var(--band);
  border-top:1px solid var(--axis)}
.note{margin:8px 0 0; font-size:13px; color:var(--muted); max-width:86ch;
  line-height:1.6}

/* charts */
.chart{display:block; width:100%; height:auto; margin:0 0 12px;
  background:var(--surface); border:1px solid var(--rule); border-radius:3px;
  padding:6px 4px}
.chart .bar{fill:var(--accent)}
/* the hover layer: a transparent target the size of the whole row or column,
   so a one-pixel bar is still reachable, and the row it belongs to lights up
   as a unit rather than just the mark */
.chart .hit{fill:transparent}
.chart .row:hover .bar{filter:brightness(1.25)}
.chart .row:hover .blab,.chart .row:hover .vlab{fill:var(--ink);
  font-weight:600}
.chart .grid{stroke:var(--grid); stroke-width:1}
.chart .axis{stroke:var(--axis); stroke-width:1}
.chart .tick{fill:var(--muted); font-size:11.5px}
.chart .blab{fill:var(--ink2); font-size:12px;
  font-family:var(--font)}
.chart .vlab{fill:var(--ink); font-size:11.5px; font-variant-numeric:tabular-nums}

@media (max-width:980px){ .toc ul{columns:2} }
@media (max-width:760px){
  .versus{grid-template-columns:1fr}
  .toc ul{columns:1}
  .mast{padding-block:30px 22px}
  .vs .big{font-size:37px}
  .tabs{gap:0}
  .tab{padding:11px 12px; font-size:14px}
}
@media print{
  /* all three pages print in sequence - a tab a reader cannot click is a tab
     they would never know was there. */
  .js .page{display:block !important}
  .tabs{display:none}
  .page{break-before:page}
  .toc{break-inside:avoid} .tbl{break-inside:avoid}
  .scroll.tall{max-height:none}
  /* On paper the theme flips back to ink on white: a black page prints as a
     solid black rectangle, and white type on it disappears entirely. The
     series colours revert to their LIGHT steps, which are the ones validated
     against a light surface. */
  :root{
    --plane:#ffffff; --surface:#ffffff; --raise:#ffffff;
    --ink:#0b0b0b; --ink2:#333230; --muted:#6a6864;
    --grid:#e1e0d9; --axis:#a9a89e; --rule:rgba(11,11,11,.22);
    --presc:#2a78d6; --stock:#eb6834; --chw:#1baf7a;
    --band:#f1f0ec;
  }
  body{background:#fff; color:#0b0b0b}
}
</style>
"""

# Trebuchet MS ships with Windows and macOS, so nothing is fetched at all now -
# the document has no external requests left and renders the same offline.
FONTS = ""

TABS_JS = """
<script>
(function(){
  var ids = ['stock','presc','chw'];
  var pages = {}, tabs = {};
  ids.forEach(function(id){
    pages[id] = document.getElementById(id);
    tabs[id]  = document.getElementById('tab-' + id);
  });
  function show(id){
    if(!pages[id]) return;
    ids.forEach(function(o){
      pages[o].classList.toggle('on', o === id);
      tabs[o].classList.toggle('on', o === id);
      tabs[o].setAttribute('aria-selected', o === id ? 'true' : 'false');
    });
  }
  // Any button carrying data-page switches pages.
  document.addEventListener('click', function(e){
    var b = e.target.closest('[data-page]');
    if(!b) return;
    show(b.dataset.page);
    history.replaceState(null, '', '#' + b.dataset.page);
    window.scrollTo(0, 0);
  });
  // Deep links: #chw opens that page, #chw-overview opens it and scrolls.
  function fromHash(){
    var h = decodeURIComponent(location.hash.slice(1));
    if(!h) return;
    var pid = ids.filter(function(p){ return h === p || h.indexOf(p + '-') === 0; })[0];
    if(!pid) return;
    show(pid);
    if(h !== pid){
      var el = document.getElementById(h);
      if(el) el.scrollIntoView();
    }
  }
  window.addEventListener('hashchange', fromHash);
  fromHash();
})();
</script>
"""

BODY = f"""
<div class="wrap">

<header class="mast">
  <p class="eyebrow">CHAI Rwanda &middot; Unserved prescriptions study &middot;
  Headline indicators</p>
  <h1>Medicine stock and Prescription assessment findings</h1>
  <div class="stats">{HERO_STATS}</div>
</header>

<nav class="tabs" role="tablist" aria-label="Assessments">{TABS}</nav>

{stock_html}
{presc_html}
{chw_html}

</div>
{TABS_JS}
"""

# Marks the document as scripted before the body paints, so the tab CSS only
# hides pages where the tab bar can actually bring them back. Without JS, and
# in a reader that strips scripts, all three pages simply render in sequence.
JS_FLAG = "<script>document.documentElement.classList.add('js')</script>"

CONTENT = ("<title>Medicine stock and Prescription assessment findings</title>" + JS_FLAG + FONTS + CSS + BODY)

# Standalone document for opening off disk. The Artifact fragment written at
# the end is the same content minus the wrapper, since the publish step supplies
# its own <!doctype>/<head>/<body>.
with open(OUT_HTML, "w", encoding="utf-8") as fh:
    fh.write('<!doctype html>\n<html lang="en">\n<head>\n'
             '<meta charset="utf-8">\n'
             '<meta name="viewport" content="width=device-width,initial-scale=1">\n'
             '<title>Medicine stock and Prescription assessment findings</title>\n'
             + JS_FLAG + FONTS + CSS + "</head>\n<body>\n" + BODY
             + "\n</body>\n</html>\n")

ALL_PAGES = (("Medicine assessment", STOCK_SECTIONS),
             ("Prescription", PRESC_SECTIONS),
             ("Medicine assessment for CHWs", CHW_SECTIONS))
size = os.path.getsize(OUT_HTML)
print(f"wrote {OUT_HTML}  ({size:,} bytes)")
for label, secs in ALL_PAGES:
    print(f"  {label}: {len(secs)} sections, "
          f"{sum(len(ts) for _, ts in secs)} tables, "
          f"{sum(1 for _, ts in secs for t in ts if t.chart and t.chart_data())} charts")

if OUT_FRAGMENT:
    with open(OUT_FRAGMENT, "w", encoding="utf-8") as fh:
        fh.write(CONTENT)
    print(f"wrote {OUT_FRAGMENT}  ({os.path.getsize(OUT_FRAGMENT):,} bytes)")
