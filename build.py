"""Build resume (HTML -> PDF) and website (index.html) from data.json.

Usage:  python build.py        # edit data.json first, then run this, then git push
"""
import json, html, subprocess, pathlib, os

ROOT = pathlib.Path(__file__).parent
D = json.loads((ROOT / "data.json").read_text(encoding="utf-8"))
e = html.escape

FONTS = ('<link rel="preconnect" href="https://fonts.googleapis.com"><link rel="preconnect" href="https://fonts.gstatic.com" crossorigin>'
         '<link href="https://fonts.googleapis.com/css2?family=Fraunces:opsz,wght@9..144,400;9..144,600;9..144,700&'
         'family=Manrope:wght@400;500;600;700&display=swap" rel="stylesheet">')

# ---------------------------------------------------------------- RESUME ---
def jobs(lst):
    out = []
    for j in lst:
        pts = "".join(f"<li>{e(p)}</li>" for p in j["points"])
        out.append(f'''<div class="job"><div class="jh"><div><b>{e(j["role"])}</b>
<span class="org">{e(j["org"])}</span></div><span class="dt">{e(j["dates"])}</span></div><ul>{pts}</ul></div>''')
    return "".join(out)

def resume():
    links = "".join(f'<a href="{l["url"]}"><i>{e(l["label"])}</i>{e(l["handle"])}</a>' for l in D["links"])
    stats = "".join(f'<div class="st"><b>{e(s["n"])}</b><span>{e(s["l"])}</span></div>' for s in D["stats"])
    skills = "".join(f'<div class="sk"><h4>{e(g["group"])}</h4><p>{" · ".join(e(i) for i in g["items"])}</p></div>' for g in D["skills"])
    edu = "".join(f'<div class="ed"><b>{e(x["degree"])}</b><span>{e(x["school"])}</span><em>{e(x["when"])}</em></div>' for x in D["education"])
    certs = "".join(f'<div class="ce"><b>{e(c["name"])}</b><span>{e(c["by"])} · {e(c["when"])}</span></div>' for c in D["certifications"])
    pats = "".join(f'<li><span class="pn">0{i+1}</span>{e(p)}</li>' for i, p in enumerate(D["patents"]))
    pubs = "".join(f'<li><span class="pv">{e(p["venue"])}</span><b>{e(p["title"])}</b> <em>{e(p["where"])}</em></li>' for p in D["publications"])
    write = "".join(f'<li><a href="{w["url"]}">{e(w["title"])}</a></li>' for w in D["writing"][:3])
    return f'''<!doctype html><html><head><meta charset="utf-8"><title>{e(D["name"])} — Resume</title>{FONTS}
<style>
@page {{ size: A4; margin: 0; }}
:root {{ --ink:#16202a; --mut:#5b6670; --line:#dcdfe2; --acc:#0e5a5c; --acc2:#b8742a; --side:#f3f1ec; }}
* {{ box-sizing:border-box; margin:0; padding:0; }}
body {{ font-family:Manrope,sans-serif; color:var(--ink); font-size:8.7pt; line-height:1.42; -webkit-print-color-adjust:exact; print-color-adjust:exact; }}
a {{ color:inherit; text-decoration:none; }}
.page {{ width:210mm; height:297mm; display:grid; grid-template-columns:64mm 1fr; overflow:hidden; page-break-after:always; }}
.page:last-child {{ page-break-after:auto; }}
aside {{ background:var(--side); padding:13mm 7mm 9mm 10mm; border-right:1px solid var(--line); }}
main {{ padding:14mm 11mm 12mm 9mm; }}
.name {{ font-family:Fraunces,serif; font-weight:700; font-size:27pt; line-height:.98; letter-spacing:-.5px; color:var(--ink); }}
.name span {{ color:var(--acc); }}
.role {{ margin-top:3mm; font-size:8.4pt; font-weight:600; letter-spacing:.06em; text-transform:uppercase; color:var(--acc2); }}
.tag {{ margin-top:3.5mm; font-family:Fraunces,serif; font-size:10.5pt; line-height:1.38; color:#33404b; }}
h3 {{ font-family:Manrope; font-size:7.6pt; font-weight:700; letter-spacing:.18em; text-transform:uppercase; color:var(--acc);
     margin:5mm 0 2.2mm; display:flex; align-items:center; gap:2.5mm; }}
h3::after {{ content:""; flex:1; height:1px; background:var(--line); }}
aside h3 {{ margin-top:6mm; }}
.contact div, .contact a {{ display:block; margin-bottom:1.3mm; word-break:break-all; }}
.contact i, .links i {{ display:block; font-style:normal; font-size:6.6pt; letter-spacing:.12em; text-transform:uppercase; color:var(--mut); }}
.links a {{ display:block; margin-bottom:1.8mm; font-weight:600; color:var(--acc); }}
.stats {{ display:grid; grid-template-columns:1fr 1fr; gap:3mm 2mm; }}
.st b {{ display:block; font-family:Fraunces,serif; font-size:17pt; color:var(--acc); line-height:1; }}
.st span {{ font-size:7pt; color:var(--mut); line-height:1.25; display:block; margin-top:.8mm; }}
.sk {{ margin-bottom:2.6mm; }} .sk h4 {{ font-size:7.8pt; font-weight:700; }} .sk p {{ color:#3c4852; font-size:7.9pt; }}
.ed, .ce {{ margin-bottom:2.4mm; }} .ed b, .ce b {{ display:block; font-size:8.1pt; }}
.ed span, .ce span {{ display:block; color:#3c4852; font-size:7.7pt; }} .ed em {{ font-style:normal; color:var(--acc2); font-size:7.5pt; font-weight:600; }}
.ce span {{ color:var(--mut); }}
.summary {{ font-size:9pt; color:#2a3540; }}
.job {{ margin-bottom:3.4mm; }}
.jh {{ display:flex; justify-content:space-between; gap:4mm; align-items:baseline; }}
.jh b {{ font-size:9.4pt; }} .org {{ display:block; color:var(--acc); font-weight:600; font-size:8.3pt; }}
.dt {{ white-space:nowrap; font-size:7.6pt; font-weight:600; color:var(--mut); }}
ul {{ margin-top:1.2mm; padding-left:3.6mm; }} li {{ margin-bottom:.7mm; }} li::marker {{ color:var(--acc2); }}
.pat {{ list-style:none; padding:0; }} .pat li {{ display:flex; gap:3mm; margin-bottom:2.6mm; font-weight:500; font-size:9.4pt; }}
.pn {{ font-family:Fraunces,serif; font-weight:700; color:var(--acc2); font-size:10pt; line-height:1.1; }}
.pubs {{ list-style:none; padding:0; }} .pubs li {{ margin-bottom:3.2mm; padding-left:0; font-size:9pt; }}
.pubs b {{ font-weight:600; }} .pubs em {{ color:var(--mut); font-style:normal; font-size:7.8pt; }}
.pv {{ display:inline-block; font-size:6.8pt; font-weight:700; letter-spacing:.05em; color:#fff; background:var(--acc); padding:.25mm 1.6mm; border-radius:1mm; margin-right:1.5mm; }}
.wr {{ padding-left:3.6mm; }} .wr a {{ color:var(--acc); font-weight:600; }}
.foot {{ margin-top:4mm; font-size:7.3pt; color:var(--mut); }}
</style></head><body>
<section class="page">
<aside>
  <div class="name">Mahesh<br><span>Basavaraj</span></div>
  <div class="role">Asst. Professor · AI Researcher</div>
  <h3>Contact</h3>
  <div class="contact"><div><i>Location</i>{e(D["location"])}</div><div><i>Phone</i>{e(D["phone"])}</div>
  {"".join(f'<a href="mailto:{m}"><i>Email</i>{e(m)}</a>' for m in D["emails"])}</div>
  <h3>Online</h3><div class="links">{links}</div>
  <h3>At a glance</h3><div class="stats">{stats}</div>
  <h3>Skills</h3>{skills}
</aside>
<main>
  <div class="tag">{e(D["tagline"])}</div>
  <h3>Profile</h3><p class="summary">{e(D["summary"])}</p>
  <h3>Academic Experience</h3>{jobs(D["academic"])}
  <h3>Industry Experience</h3>{jobs(D["industry"])}
</main>
</section>
<section class="page">
<aside>
  <h3 style="margin-top:0">Education</h3>{edu}
  <h3>Certifications &amp; FDPs</h3>{certs}
  <h3>Writing &amp; Media</h3>
  <ul class="wr">{write}</ul>
  <p class="foot">More essays on Substack and AI explainers on the <b>Echoes of AI</b> YouTube channel.</p>
</aside>
<main>
  <h3 style="margin-top:0">Patents</h3><ul class="pat">{pats}</ul>
  <h3>Research Publications</h3><ul class="pubs">{pubs}</ul>
</main>
</section>
</body></html>'''

# ---------------------------------------------------------------- WEBSITE ---
def site():
    j = json.dumps(D, ensure_ascii=False)
    return (ROOT / "site_template.html").read_text(encoding="utf-8").replace("/*DATA*/null", j).replace("{{FONTS}}", FONTS)

if __name__ == "__main__":
    (ROOT / "resume.html").write_text(resume(), encoding="utf-8")
    (ROOT / "index.html").write_text(site(), encoding="utf-8")
    chrome = r"C:\Program Files\Google\Chrome\Application\chrome.exe"
    subprocess.run([chrome, "--headless=new", "--disable-gpu", "--no-pdf-header-footer", "--virtual-time-budget=8000",
                    f"--print-to-pdf={ROOT / 'Mahesh_Basavaraj_Resume.pdf'}", (ROOT / "resume.html").as_uri()], check=True)
    print("built resume.html, index.html, Mahesh_Basavaraj_Resume.pdf")
