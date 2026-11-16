from docx import Document
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.section import WD_SECTION
from docx.oxml import OxmlElement
from docx.oxml.ns import qn

OUT = "/Users/apple/Desktop/starter repository/Assignment_Submission_u6726116.docx"

doc = Document()
sec = doc.sections[0]
sec.top_margin = Inches(.72)
sec.bottom_margin = Inches(.72)
sec.left_margin = Inches(.82)
sec.right_margin = Inches(.82)

styles = doc.styles
styles["Normal"].font.name = "Aptos"
styles["Normal"]._element.rPr.rFonts.set(qn("w:ascii"), "Aptos")
styles["Normal"]._element.rPr.rFonts.set(qn("w:hAnsi"), "Aptos")
styles["Normal"].font.size = Pt(10.5)
styles["Normal"].paragraph_format.space_after = Pt(6)
for nm in ("Title", "Heading 1", "Heading 2"):
    styles[nm].font.name = "Aptos Display"
    styles[nm]._element.rPr.rFonts.set(qn("w:ascii"), "Aptos Display")
    styles[nm]._element.rPr.rFonts.set(qn("w:hAnsi"), "Aptos Display")
    styles[nm].font.color.rgb = RGBColor(0, 0, 0)
styles["Title"].font.size = Pt(25)
styles["Title"].font.bold = True
title_ppr = styles["Title"]._element.get_or_add_pPr()
title_bdr = title_ppr.find(qn("w:pBdr"))
if title_bdr is not None:
    title_ppr.remove(title_bdr)
styles["Heading 1"].font.size = Pt(16)
styles["Heading 1"].font.bold = True
styles["Heading 2"].font.size = Pt(12)
styles["Heading 2"].font.bold = True

def shade(cell, fill):
    tcPr = cell._tc.get_or_add_tcPr()
    shd = OxmlElement("w:shd")
    shd.set(qn("w:fill"), fill)
    tcPr.append(shd)

def borders(table):
    tblPr = table._tbl.tblPr
    b = OxmlElement("w:tblBorders")
    for edge in ("top", "left", "bottom", "right", "insideH", "insideV"):
        el = OxmlElement("w:" + edge)
        el.set(qn("w:val"), "single")
        el.set(qn("w:sz"), "5")
        el.set(qn("w:color"), "D9D9D9")
        b.append(el)
    tblPr.append(b)

def table(headers, rows, widths=None):
    t = doc.add_table(rows=1, cols=len(headers))
    t.autofit = False
    borders(t)
    for i, h in enumerate(headers):
        c = t.rows[0].cells[i]
        c.text = h
        shade(c, "DCE6F1")
        for r in c.paragraphs[0].runs:
            r.bold = True
        if widths: c.width = Inches(widths[i])
    for row in rows:
        cells = t.add_row().cells
        for i, val in enumerate(row):
            cells[i].text = val
            if widths: cells[i].width = Inches(widths[i])
    doc.add_paragraph().paragraph_format.space_after = Pt(1)
    return t

def mono(text):
    p = doc.add_paragraph()
    p.paragraph_format.left_indent = Inches(.18)
    p.paragraph_format.space_after = Pt(2)
    r = p.add_run(text)
    r.font.name = "Menlo"
    r._element.rPr.rFonts.set(qn("w:ascii"), "Menlo")
    r._element.rPr.rFonts.set(qn("w:hAnsi"), "Menlo")
    r.font.size = Pt(8.5)
    return p

p = doc.add_paragraph(style="Title")
p.add_run("Git and Website Delivery Assignment")
p = doc.add_paragraph()
p.paragraph_format.space_after = Pt(12)
r = p.add_run("Student u6726116  |  Repository evidence and recovery reflection")
r.font.size = Pt(11)
r.font.color.rgb = RGBColor(75, 85, 99)

doc.add_paragraph(
    "I updated the website message on a feature branch, reviewed and merged the change into ci-demo, "
    "and verified the successful delivery run and artifact. The final artifact metadata identifies the "
    "same merge commit shown by Git and the workflow run. The recovery exercise was completed in a "
    "separate disposable sandbox."
)

doc.add_heading("Repository and pull request", level=1)
table(["Item", "Evidence"], [
    ["Repository", "https://github.com/Needyu198/ead_week2 (public)"],
    ["Pull request", "https://github.com/Needyu198/ead_week2/pull/1 — merged into ci-demo"],
    ["Base branch", "ci-demo — receives the proposed change"],
    ["Compare branch", "feature/u6726116 — contains the proposed website message change"],
    ["Feature commit", "b9a85eede69f998482df3a5082a9958ef8f8ce33"],
    ["Merge commit", "dc6d309244dbf0c04891cff3db8d2216b05afaec"],
    ["PR check", "Passed; PR run: https://github.com/Needyu198/ead_week2/actions/runs/43830278398"],
    ["Review", "Solo review recorded in the PR description; review note: “good.”"],
], [1.35, 5.95])
doc.add_paragraph(
    "The compare branch is the source under review; the base branch is where GitHub applies the merge. "
    "The repository is public, so the instructor can view it without a collaborator invitation. No "
    "passwords, tokens, keys, or credentials are included."
)

doc.add_heading("Branch commands and labelled output", level=1)
mono("COMMAND  git switch ci-demo && git pull --ff-only")
mono("OUTPUT   Already on 'ci-demo'  |  Fast-forward 113aee0..dc6d309")
mono("COMMAND  git rev-parse HEAD")
mono("OUTPUT   dc6d309244dbf0c04891cff3db8d2216b05afaec")
mono("COMMAND  rg -n 'Student u6726116|<h1>' src/index.html")
mono("OUTPUT   11: <h1>Enterprise Application Development</h1>")
mono("         12: <p id=\"message\">Student u6726116 delivered this website through a repeatable pipeline.</p>")
mono("COMMAND  git status --short --branch")
mono("OUTPUT   ## ci-demo...origin/ci-demo  (clean and synchronized)")
doc.add_paragraph(
    "The source diff changed only the message paragraph in src/index.html. Staging was limited to "
    "that file, and the informative commit was “Personalize website delivery message.”"
)

doc.add_heading("Successful delivery run and artifact", level=1)
doc.add_paragraph("Post-merge ci-demo push run: https://github.com/Needyu198/ead_week2/actions/runs/43831052006")
table(["Evidence", "Value"], [
    ["Workflow result", "Success; tests, actual-site validation, build, and artifact upload passed"],
    ["Run SHA", "dc6d309244dbf0c04891cff3db8d2216b05afaec"],
    ["Artifact", "website-dc6d309244dbf0c04891cff3db8d2216b05afaec"],
    ["Outer archive", "site.zip and build-info.json"],
    ["Inner site archive", "index.html, style.css, and build-info.json"],
], [1.55, 5.75])
doc.add_paragraph("Excerpt from the downloaded artifact’s build-info.json:")
mono('{')
mono('  "commit": "dc6d309244dbf0c04891cff3db8d2216b05afaec",')
mono('  "branch": "ci-demo",')
mono('  "run_id": "43831052006",')
mono('  "source_files_sha256": {')
mono('    "index.html": "42670a5e4a809b1037102ffb216ea1a93e3cfc21020a078b8213997545f1292c",')
mono('    "style.css": "303eab4649f4fbbcc9846193222238bcdc1533a37f02634f2b1b25187c09a50d"')
mono('  }')
mono('}')
doc.add_paragraph(
    "The artifact commit and run_id match the successful workflow run. The extracted index.html "
    "contains the student-specific message above. The PR run has no artifact because the workflow "
    "uploads only when the event is a push to refs/heads/ci-demo; a pull_request event runs checks and build only."
)

doc.add_heading("Lab 5 discard and restore", level=1)
doc.add_paragraph(
    "I used the separate sandbox at /Users/apple/Desktop/ead-recovery-sandbox-lab5-20261116. "
    "The sandbox history contains V1 d2bd2c2 (Save version one), V2 de7edd8 (Save version two), "
    "and rollback snapshot adea2b9 (Restore version one as a new snapshot). The backup-v2 branch "
    "continues to point to V2."
)
table(["Recovery action", "Observed result"], [
    ["Restore tracked state to V2", "app.txt returned to version two; config.txt returned; scratch.txt was previewed by ordinary clean; ignored cache/ remained."],
    ["Reset sandbox to V1", "app.txt showed version one; config.txt existed; feature.txt was absent because V1 did not contain it."],
    ["Clean ignored output", "git clean -ndx previewed cache/ before the sandbox cleanup removed it."],
    ["Restore V1 content while HEAD was V2", "app.txt changed and feature.txt was staged for deletion; committing created adea2b9 while V2 remained in history."],
], [2.15, 5.15])
doc.add_paragraph(
    "An untracked file such as scratch.txt has never been saved in a commit, so Git has no historical "
    "version to restore. Ordinary git clean removes untracked, non-ignored files; ignored files are "
    "left alone unless the force option also includes ignored files. All reset and clean operations "
    "were confined to the disposable sandbox. No destructive recovery commands were run in the CI repository."
)

doc.add_heading("Evidence files", level=1)
doc.add_paragraph(
    "The downloaded artifacts and extracted files are retained in /Users/apple/Desktop/ead-week2-lab6-evidence. "
    "The assignment artifact is at assignment-merge/website-dc6d309.zip; the build metadata is in "
    "assignment-merge/outer/build-info.json."
)

footer = sec.footer.paragraphs[0]
footer.alignment = WD_ALIGN_PARAGRAPH.RIGHT
rr = footer.add_run("u6726116  |  Git and website delivery assignment")
rr.font.size = Pt(8)
rr.font.color.rgb = RGBColor(100, 100, 100)

doc.save(OUT)
print(OUT)
