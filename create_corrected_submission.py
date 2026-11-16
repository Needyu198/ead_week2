from docx import Document
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn

out = "/Users/apple/Desktop/starter repository/Assignment_Submission_u6726116_Corrected.docx"
d = Document()
s = d.sections[0]
s.top_margin = Inches(.65); s.bottom_margin = Inches(.65)
s.left_margin = Inches(.78); s.right_margin = Inches(.78)
for name in ("Normal", "Title", "Heading 1", "Heading 2"):
    d.styles[name].font.name = "Arial"
    d.styles[name]._element.rPr.rFonts.set(qn("w:ascii"), "Arial")
    d.styles[name]._element.rPr.rFonts.set(qn("w:hAnsi"), "Arial")
    d.styles[name].font.color.rgb = RGBColor(0,0,0)
d.styles["Normal"].font.size = Pt(9.5)
d.styles["Normal"].paragraph_format.space_after = Pt(4)
d.styles["Title"].font.size = Pt(22); d.styles["Title"].font.bold = True
tp=d.styles["Title"]._element.get_or_add_pPr(); b=tp.find(qn("w:pBdr"))
if b is not None: tp.remove(b)
d.styles["Heading 1"].font.size=Pt(14); d.styles["Heading 1"].font.bold=True
d.styles["Heading 2"].font.size=Pt(11); d.styles["Heading 2"].font.bold=True

def borders(t):
    b=OxmlElement("w:tblBorders")
    for e in ("top","left","bottom","right","insideH","insideV"):
        x=OxmlElement("w:"+e); x.set(qn("w:val"),"single"); x.set(qn("w:sz"),"5"); x.set(qn("w:color"),"D9D9D9"); b.append(x)
    t._tbl.tblPr.append(b)
def table(headers, rows, widths):
    t=d.add_table(rows=1, cols=len(headers)); t.autofit=False; borders(t)
    for i,h in enumerate(headers):
        c=t.rows[0].cells[i]; c.text=h; c.width=Inches(widths[i])
        sh=OxmlElement("w:shd"); sh.set(qn("w:fill"),"DCE6F1"); c._tc.get_or_add_tcPr().append(sh)
        for r in c.paragraphs[0].runs:r.bold=True
    for row in rows:
        cells=t.add_row().cells
        for i,v in enumerate(row): cells[i].text=v; cells[i].width=Inches(widths[i])
    d.add_paragraph().paragraph_format.space_after=Pt(1)
def mono(text):
    p=d.add_paragraph(); p.paragraph_format.space_after=Pt(1); p.paragraph_format.left_indent=Inches(.12)
    r=p.add_run(text); r.font.name="Courier New"; r._element.rPr.rFonts.set(qn("w:ascii"),"Courier New"); r.font.size=Pt(8)

d.add_paragraph("Git and Website Delivery Assignment",style="Title")
p=d.add_paragraph("Student u6726116 | Repository, pull request, delivery, and recovery evidence")
p.runs[0].font.color.rgb=RGBColor(70,80,95)
d.add_paragraph("I changed the website message on a feature branch, reviewed and merged it into ci-demo, and verified the successful delivery run and downloaded artifact. The artifact metadata matches the merge commit and workflow run. Lab 5 recovery work was done in a separate disposable sandbox.")

d.add_heading("Repository and pull request",1)
table(["Item","Evidence"],[
 ["Repository","https://github.com/Needyu198/ead_week2 (public; instructor can view)"],
 ["Pull request","https://github.com/Needyu198/ead_week2/pull/1 — merged into ci-demo"],
 ["Base / compare","Base ci-demo receives the change; compare feature/u6726116 contains it."],
 ["Feature commit","b9a85eede69f998482df3a5082a9958ef8f8ce33"],
 ["Merge commit","dc6d309244dbf0c04891cff3db8d2216b05afaec"],
 ["PR check","Passed: https://github.com/Needyu198/ead_week2/actions/runs/43830278398"],
 ["Review","Solo review recorded on the PR; review note: good."],
],[1.35,5.95])
d.add_paragraph("The feature branch was created from updated ci-demo. The reviewed diff changed only the message paragraph in src/index.html; the required single h1 and stylesheet reference remained intact. Only src/index.html was staged. Commit: Personalize website delivery message.")

d.add_heading("Labelled command and history evidence",1)
for line in [
 "COMMAND  git diff -- src/index.html",
 "OUTPUT   - A small website delivered by a repeatable pipeline.",
 "         + Student u6726116 delivered this website through a repeatable pipeline.",
 "COMMAND  git add src/index.html && git diff --staged --stat",
 "OUTPUT   src/index.html | 2 +−  (only this file staged)",
 "COMMAND  git commit -m \"Personalize website delivery message\"",
 "OUTPUT   [feature/u6726116 b9a85ee] Personalize website delivery message",
 "COMMAND  git branch -vv",
 "OUTPUT   ci-demo 113aee0 [origin/ci-demo] Restore required page heading",
 "         feature/u6726116 b9a85ee [origin/feature/u6726116] Personalize website delivery message",
 "COMMAND  git log --oneline --graph --decorate --all",
 "OUTPUT   * b9a85ee (feature/u6726116) Personalize website delivery message",
 "         * 113aee0 (ci-demo before merge) Restore required page heading",
 "         * 15092f7 Demonstrate a missing required heading",
 "         * 013aa45 (main) Clarify repository delivery workflow",
 "COMMAND  git switch ci-demo && git pull --ff-only",
 "OUTPUT   Fast-forward 113aee0..dc6d309",
 "COMMAND  git rev-parse HEAD",
 "OUTPUT   dc6d309244dbf0c04891cff3db8d2216b05afaec",
 "COMMAND  git status --short --branch",
 "OUTPUT   ## ci-demo...origin/ci-demo  (clean and synchronized)",
]: mono(line)

d.add_heading("Successful delivery run and artifact",1)
d.add_paragraph("Post-merge ci-demo run: https://github.com/Needyu198/ead_week2/actions/runs/43831052006. The tests, actual-site validator, build, and upload all succeeded.")
table(["Check","Verified result"],[
 ["Run SHA","dc6d309244dbf0c04891cff3db8d2216b05afaec"],
 ["Artifact","website-dc6d309244dbf0c04891cff3db8d2216b05afaec"],
 ["Outer archive","site.zip and build-info.json"],
 ["Inside site.zip","index.html, style.css, build-info.json"],
 ["Website change","Student u6726116 delivered this website through a repeatable pipeline."],
],[1.45,5.85])
d.add_paragraph("Excerpt from the downloaded artifact build-info.json:")
for line in ['{','  "commit": "dc6d309244dbf0c04891cff3db8d2216b05afaec",','  "branch": "ci-demo",','  "run_id": "43831052006"','}']:mono(line)
d.add_paragraph("The artifact’s commit and run_id match the successful run. The PR run has no artifact because upload is conditioned on a push to ci-demo; a pull_request event runs checks and build only.")

d.add_heading("Lab 5 discard and restore",1)
d.add_paragraph("I used /Users/apple/Desktop/ead-recovery-sandbox-lab5-20261116. V1 is d2bd2c2, V2 is de7edd8, and rollback snapshot adea2b9 records V1’s contents while keeping V2 in history. The backup-v2 branch still points to V2.")
table(["Action","Observed result"],[
 ["Restore tracked state to V2","app.txt returned to version two; config.txt returned; ordinary clean previewed scratch.txt; ignored cache/ remained."],
 ["Reset sandbox to V1","app.txt contained version one; config.txt existed; feature.txt was absent because V1 did not contain it."],
 ["Remove ignored output","git clean -ndx previewed cache/ before cleanup removed it."],
 ["Restore V1 content with HEAD at V2","app.txt changed and feature.txt was staged for deletion; commit adea2b9 saved the rollback without moving HEAD back."],
],[2.05,5.25])
d.add_paragraph("Git cannot restore an uncommitted file such as scratch.txt from history because Git never saved it in a commit. Ordinary git clean removes untracked non-ignored files; ignored cache/ remains unless cleanup explicitly includes ignored files. Reset and clean commands were run only in the disposable sandbox, never in the CI repository.")
d.add_paragraph("The repository is public, so the instructor has access. No passwords, tokens, keys, or credentials are included.")

f=s.footer.paragraphs[0]; f.alignment=WD_ALIGN_PARAGRAPH.RIGHT
r=f.add_run("u6726116 | Assignment evidence"); r.font.size=Pt(8); r.font.color.rgb=RGBColor(100,100,100)
d.save(out)
print(out)
