"""Write one list of content blocks to LaTeX (compiled to PDF) and to Word.

Blocks:
  ("h1"|"h2"|"h3", title)            numbered automatically (no number in title)
  ("p", text)                        **bold** supported
  ("bullets", [text, ...])
  ("table", title, header, rows, widths_cm, note)   numbered "Table n"
  ("fig", path, caption, width_cm)                  numbered "Figure n"
  ("refs", [text, ...])
  ("pagebreak",)
  ("unnumbered_h1", title)          e.g. executive summary, references
"""
import os
import re
import subprocess

TEX = os.path.expanduser("~/Library/TinyTeX/bin/universal-darwin")


def strip_num(t):
    return re.sub(r"^(\d+(\.\d+)*\.?|[A-Z]\.)\s+", "", t)


# ------------------------------------------------------------------ LaTeX
def tx(s):
    s = str(s)
    rep = {"\\": r"\textbackslash{}", "&": r"\&", "%": r"\%", "$": r"\$", "#": r"\#",
           "_": r"\_", "{": r"\{", "}": r"\}", "~": r"\textasciitilde{}",
           "^": r"\textasciicircum{}"}
    s = "".join(rep.get(ch, ch) for ch in s)
    s = (s.replace("\u00b7", r"\textperiodcentered{}").replace("\u00b1", r"$\pm$")
         .replace("\u03c0\u00b2", r"$\pi^2$").replace(">=", r"$\geq$").replace("<=", r"$\leq$")
         .replace("<", r"$<$").replace(">", r"$>$").replace("\"", "''")
         .replace("\u2013", "--").replace("\u2014", "---").replace("\u00e9", r"\'e"))
    s = re.sub(r"\*\*(.+?)\*\*", r"\\textbf{\1}", s)
    s = re.sub(r"(\d) x (\d)", r"\1 $\\times$ \2", s)
    s = s.replace(" x ", r" $\times$ ")
    return s


def to_tex(blocks, path, title, subtitle, meta, figdirs):
    o = [r"""\documentclass[10pt,a4paper]{report}
\usepackage[T1]{fontenc}
\usepackage[utf8]{inputenc}
\usepackage[margin=2.2cm,headheight=14pt]{geometry}
\usepackage{helvet}\renewcommand{\familydefault}{\sfdefault}
\usepackage{booktabs,longtable,array,xcolor,colortbl,enumitem,titlesec,graphicx,caption,float}
\usepackage{fancyhdr,lastpage,tocloft}
\usepackage[hidelinks]{hyperref}
\definecolor{navy}{HTML}{1F3A68}\definecolor{band}{HTML}{EEF2F7}
\titleformat{\chapter}[hang]{\LARGE\bfseries\color{navy}}{\thechapter}{12pt}{}
\titlespacing*{\chapter}{0pt}{-20pt}{14pt}
\titleformat{\section}{\large\bfseries\color{navy}}{\thesection}{8pt}{}
\titleformat{\subsection}{\normalsize\bfseries\color{navy}}{\thesubsection}{8pt}{}
\setlist{nosep,leftmargin=1.3em,topsep=2pt}
\setlength{\parindent}{0pt}\setlength{\parskip}{5pt}
\captionsetup{font=small,labelfont={bf,color=navy},justification=raggedright,singlelinecheck=false}
\captionsetup[longtable]{position=top}
\graphicspath{""" + "".join("{" + os.path.relpath(d, os.path.dirname(path)).rstrip("/") + "/}" for d in figdirs) + r"""}
\pagestyle{fancy}\fancyhf{}
\renewcommand{\chaptermark}[1]{\markboth{\thechapter.\ #1}{}}
\fancyhead[L]{\footnotesize\color{gray} QMIA 2026}
\fancyhead[R]{\footnotesize\color{gray}\nouppercase{\leftmark}}
\fancyfoot[R]{\footnotesize Page \thepage\ of \pageref{LastPage}}
\renewcommand{\headrulewidth}{0.3pt}
\fancypagestyle{plain}{\fancyhf{}\fancyfoot[R]{\footnotesize Page \thepage\ of \pageref{LastPage}}\renewcommand{\headrulewidth}{0pt}}
\begin{document}
\begin{titlepage}\vspace*{2.5cm}
{\color{navy}\rule{\linewidth}{2pt}}\par\vspace{14pt}
{\raggedright\hyphenpenalty=10000\exhyphenpenalty=10000\Huge\bfseries\color{navy} """ + tx(title) + r"""\par}\vspace{10pt}
{\raggedright\hyphenpenalty=10000\Large\color{navy} """ + tx(subtitle) + r"""\par}\vspace{14pt}
{\color{navy}\rule{\linewidth}{0.8pt}}\par\vspace{20pt}
{\normalsize """ + tx(meta) + r"""\par}
\vfill{\small\color{gray} Republic of Rwanda \textperiodcentered{} Ministry of Health \textperiodcentered{} Rwanda Medical Supply \textperiodcentered{} RSSB \textperiodcentered{} Clinton Health Access Initiative\par}
\end{titlepage}
\pagenumbering{roman}
\tableofcontents
\listoftables
\listoffigures
\clearpage\pagenumbering{arabic}
"""]
    for b in blocks:
        k = b[0]
        if k == "unnumbered_h1":
            o.append(r"\chapter*{" + tx(b[1]) + r"}\addcontentsline{toc}{chapter}{" + tx(b[1])
                     + r"}\markboth{" + tx(b[1]) + "}{}")
        elif k == "h2u":
            o.append(r"\section*{" + tx(b[1]) + "}")
        elif k == "h1":
            o.append(r"\chapter{" + tx(strip_num(b[1])) + "}")
        elif k == "h2":
            o.append(r"\section{" + tx(strip_num(b[1])) + "}")
        elif k == "h3":
            o.append(r"\subsection{" + tx(strip_num(b[1])) + "}")
        elif k == "p":
            o.append(tx(b[1]) + "\n")
        elif k == "bullets":
            o.append(r"\begin{itemize}" + "".join(r"\item " + tx(x) + "\n" for x in b[1])
                     + r"\end{itemize}")
        elif k == "table":
            _, title, hdr, rows, wd, note = b
            tot = sum(wd)
            spec = "".join(r">{\raggedright\arraybackslash}p{" + f"{0.96 * x / tot:.3f}"
                           + r"\linewidth}" for x in wd)
            head = " & ".join(r"\textbf{\color{white}" + tx(h) + "}" for h in hdr)
            body = [(r"\rowcolor{band}" if i % 2 else "") + " & ".join(tx(v) for v in r)
                    + r" \\" for i, r in enumerate(rows)]
            o.append(r"{\footnotesize\setlength{\tabcolsep}{3pt}\renewcommand{\arraystretch}{1.15}" "\n"
                     r"\begin{longtable}{" + spec + "}\n"
                     r"\caption{" + tx(title) + r"}\\" "\n"
                     r"\rowcolor{navy}" + head + r" \\" "\n" r"\endfirsthead" "\n"
                     r"\rowcolor{navy}" + head + r" \\" "\n" r"\endhead" "\n"
                     + "\n".join(body) + "\n"
                     + (r"\multicolumn{" + str(len(hdr)) + r"}{p{0.96\linewidth}}{\scriptsize\itshape\color{gray} "
                        + tx(note) + r"}\\" "\n" if note else "")
                     + r"\end{longtable}}")
        elif k == "fig":
            _, fn, cap, wcm = b
            o.append(r"\begin{figure}[H]\centering\includegraphics[width="
                     + f"{min(1.0, wcm / 16.6):.2f}" + r"\linewidth]{" + os.path.basename(fn)
                     + r"}\caption{" + tx(cap) + r"}\end{figure}")
        elif k == "refs":
            o.append(r"\begin{enumerate}[leftmargin=1.6em,itemsep=2pt]\small"
                     + "".join(r"\item " + tx(x) + "\n" for x in b[1]) + r"\end{enumerate}")
        elif k == "pagebreak":
            o.append(r"\clearpage")
    o.append(r"\end{document}")
    open(path, "w", encoding="utf-8").write("\n".join(o))


def compile_pdf(texpath):
    d, f = os.path.split(texpath)
    log = ""
    for _ in range(3):
        r = subprocess.run([os.path.join(TEX, "pdflatex"), "-interaction=nonstopmode",
                            "-halt-on-error", f], cwd=d, capture_output=True, text=True)
        log = r.stdout
        if r.returncode != 0:
            errs = [l for l in log.splitlines() if l.startswith("!")]
            raise SystemExit("pdflatex failed: " + " | ".join(errs[:5]) + "\n" +
                             "\n".join(log.splitlines()[-25:]))
    warns = [l for l in log.splitlines() if "Warning" in l and "Font" not in l]
    return texpath[:-4] + ".pdf", warns


# ------------------------------------------------------------------ Word
def to_docx(blocks, path, title, subtitle, meta, figdirs):
    from docx import Document
    from docx.enum.table import WD_TABLE_ALIGNMENT
    from docx.enum.text import WD_ALIGN_PARAGRAPH
    from docx.oxml import OxmlElement
    from docx.oxml.ns import qn
    from docx.shared import Cm, Pt, RGBColor
    NAVY, GREY = RGBColor(0x1F, 0x3A, 0x68), RGBColor(0x55, 0x55, 0x55)
    doc = Document()
    sec = doc.sections[0]
    sec.page_width, sec.page_height = Cm(21.0), Cm(29.7)
    sec.left_margin = sec.right_margin = sec.top_margin = sec.bottom_margin = Cm(2.2)
    st = doc.styles
    st["Normal"].font.name, st["Normal"].font.size = "Arial", Pt(10)
    st["Normal"].paragraph_format.space_after = Pt(5)
    st["Normal"].element.rPr.rFonts.set(qn("w:eastAsia"), "Arial")
    for name, size in (("Heading 1", 16), ("Heading 2", 12.5), ("Heading 3", 10.5)):
        s_ = st[name]
        s_.font.name, s_.font.size, s_.font.bold = "Arial", Pt(size), True
        s_.font.color.rgb = NAVY
        s_.paragraph_format.space_before = Pt(14 if name != "Heading 3" else 9)
        s_.paragraph_format.space_after = Pt(4)
        s_.paragraph_format.keep_with_next = True

    def runs(par, text, size=None, color=None, italic=False, bold=None):
        for i, part in enumerate(re.split(r"\*\*", str(text))):
            if part:
                r = par.add_run(part)
                r.bold = bool(i % 2) if bold is None else bold
                r.italic = italic
                if size:
                    r.font.size = Pt(size)
                if color:
                    r.font.color.rgb = color
        return par

    def shade(cell, col):
        tcPr = cell._tc.get_or_add_tcPr()
        sh = OxmlElement("w:shd")
        sh.set(qn("w:val"), "clear"); sh.set(qn("w:color"), "auto"); sh.set(qn("w:fill"), col)
        tcPr.append(sh)

    def field(par, instr):
        r = par.add_run()
        for t_, txt in (("begin", None), (None, instr), ("separate", None), (None, "*"), ("end", None)):
            if t_:
                f = OxmlElement("w:fldChar"); f.set(qn("w:fldCharType"), t_); r._r.append(f)
            elif txt == "*":
                t = OxmlElement("w:t"); t.text = "Right-click to update field."; r._r.append(t)
            else:
                it = OxmlElement("w:instrText"); it.set(qn("xml:space"), "preserve")
                it.text = txt; r._r.append(it)

    tp = doc.add_paragraph(); tp.paragraph_format.space_before = Pt(80)
    runs(tp, "**" + title + "**", 24, NAVY)
    runs(doc.add_paragraph(), subtitle, 15, NAVY)
    runs(doc.add_paragraph(), meta, 10, GREY)
    doc.add_page_break()
    runs(doc.add_paragraph(), "**Contents**", 14, NAVY)
    field(doc.add_paragraph(), 'TOC \\o "1-2" \\h \\z \\u')
    doc.add_page_break()
    n1 = n2 = n3 = nt = nf = 0
    for b in blocks:
        k = b[0]
        if k == "unnumbered_h1":
            doc.add_heading(b[1], level=1)
        elif k == "h2u":
            doc.add_heading(b[1], level=2)
        elif k == "h1":
            n1 += 1; n2 = n3 = 0
            doc.add_heading(f"{n1}. {strip_num(b[1])}", level=1)
        elif k == "h2":
            n2 += 1; n3 = 0
            doc.add_heading(f"{n1}.{n2} {strip_num(b[1])}", level=2)
        elif k == "h3":
            n3 += 1
            doc.add_heading(f"{n1}.{n2}.{n3} {strip_num(b[1])}", level=3)
        elif k == "p":
            runs(doc.add_paragraph(), b[1])
        elif k == "bullets":
            for x in b[1]:
                runs(doc.add_paragraph(style="List Bullet"), x)
        elif k == "table":
            _, title_, hdr, rows, wd, note = b
            nt += 1
            cp = doc.add_paragraph(); cp.paragraph_format.keep_with_next = True
            runs(cp, f"**Table {nt}.** {title_}", 9, NAVY)
            t = doc.add_table(rows=1, cols=len(hdr))
            t.style = "Table Grid"; t.alignment = WD_TABLE_ALIGNMENT.CENTER; t.autofit = False
            for j, h in enumerate(hdr):
                c = t.rows[0].cells[j]; c.text = ""
                runs(c.paragraphs[0], h, 8, RGBColor(255, 255, 255), bold=True); shade(c, "1F3A68")
            for i, row in enumerate(rows):
                cells = t.add_row().cells
                for j, v in enumerate(row):
                    cells[j].text = ""; runs(cells[j].paragraphs[0], v, 8)
                    if i % 2:
                        shade(cells[j], "EEF2F7")
            for row in t.rows:
                for j, w_ in enumerate(wd):
                    row.cells[j].width = Cm(w_ * 16.6 / sum(wd))
                for c in row.cells:
                    for par in c.paragraphs:
                        par.paragraph_format.space_after = Pt(1)
            trPr = t.rows[0]._tr.get_or_add_trPr()
            el = OxmlElement("w:tblHeader"); el.set(qn("w:val"), "true"); trPr.append(el)
            if note:
                runs(doc.add_paragraph(), note, 8, GREY, italic=True)
            else:
                doc.add_paragraph().paragraph_format.space_after = Pt(2)
        elif k == "fig":
            _, fn, cap, wcm = b
            nf += 1
            full = fn if os.path.isabs(fn) else next(os.path.join(d, fn) for d in figdirs
                                                     if os.path.exists(os.path.join(d, fn)))
            doc.add_picture(full, width=Cm(min(wcm, 16.6)))
            doc.paragraphs[-1].alignment = WD_ALIGN_PARAGRAPH.CENTER
            runs(doc.add_paragraph(), f"**Figure {nf}.** {cap}", 8.5, GREY)
        elif k == "refs":
            for x in b[1]:
                par = doc.add_paragraph()
                par.paragraph_format.left_indent = Cm(0.6)
                par.paragraph_format.first_line_indent = Cm(-0.6)
                runs(par, x, 8.5)
        elif k == "pagebreak":
            doc.add_page_break()
    fp = sec.footer.paragraphs[0]; fp.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    r = fp.add_run()
    for t_, txt in (("begin", None), (None, "PAGE"), ("end", None)):
        if t_:
            f = OxmlElement("w:fldChar"); f.set(qn("w:fldCharType"), t_); r._r.append(f)
        else:
            it = OxmlElement("w:instrText"); it.set(qn("xml:space"), "preserve")
            it.text = txt; r._r.append(it)
    r.font.size = Pt(8)
    hp = sec.header.paragraphs[0]
    runs(hp, "QMIA 2026 - Medicine Stock and Prescription Assessment", 8, GREY)
    doc.save(path)
