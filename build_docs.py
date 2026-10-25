"""Build the UNSW application packet (3 documents) on the resume design system.

  Mahesh_Basavaraj_Research_Publications.docx
  Mahesh_Basavaraj_Cover_Letter_UNSW.docx
  Mahesh_Basavaraj_Teaching_Methodology.docx

Run: python build_docs.py
"""
import json, pathlib
from docx import Document
from docx.shared import Pt, Mm
from docx.enum.text import WD_ALIGN_PARAGRAPH
import docx_kit as K
from docx_kit import (INK, MUT, BODY, ACC, ACC2, WHITE, TEAL, AMBER, HEAD, SANS,
                      para, run, heading, bullet, numbered, badge, hyperlink, para_border_bottom,
                      shade_run, name_block, contact_block, two_col, single_page, finish, set_cell_bg)

ROOT = pathlib.Path(__file__).parent
D = json.loads((ROOT / "data.json").read_text(encoding="utf-8"))
PHOTO = ROOT / "assets" / "photo_circle.png"

TARGET = "Prepared for the Assistant Professor (Computer Science) role \u2014 UNSW Bengaluru Campus"
DATE = "25 October 2026"

# Verified identifiers (ORCID + Scholar confirmed Oct 2026; Scopus ID as supplied by Mahesh)
IDS = [
    ("ORCID", "0000-0002-2242-212X", "https://orcid.org/0000-0002-2242-212X"),
    ("Google Scholar", "aANFb2EAAAAJ", "https://scholar.google.com/citations?user=aANFb2EAAAAJ"),
    ("Scopus", "60345533200", None),
    ("Website", "maheshbasavaraj.github.io", "https://maheshbasavaraj.github.io/"),
]

JOURNALS = [
    dict(b="Medical Science Educator", t="Comment on \u201cIncorporating Artificial Intelligence in Qualitative Research: Exploring the Role of ChatGPT in Thematic Analysis\u201d",
         w="Vol. 36, Issue 2, pp. 1063\u20131064 \u00b7 2026", doi="10.1007/s40670-025-02618-y",
         co="with S. Banerjee, I. H. Madar"),
    dict(b="J. Neurological Sciences", t="Comment on \u201cImproving care for amyotrophic lateral sclerosis with artificial intelligence and affective computing\u201d",
         w="Vol. 486 \u00b7 2026", doi=None, co="with S. U. Adhyapak, P. Mishra"),
    dict(b="INDJCST", t="Agentic Approach in the Quest for AGI",
         w="Indian Journal of Computer Science and Technology, pp. 279\u2013289 \u00b7 2025", doi="10.59256/indjcst.20250402043",
         co=None),
    dict(b="IRJET", t="A Review of Deep Learning Techniques in Detection of Anomaly in Credit Card Transactions",
         w="Vol. 9, Issue 9 \u00b7 2022", doi=None, co=None),
]

CONFERENCES = [
    dict(b="ICIICS 2026", t="Multi-Scale Spatio-Temporal Network with Frequency\u2013Spatial Parallel Attention for Motor Imagery Classification",
         w="3rd Intl. Conf. on Integrated Intelligence and Communication Systems, pp. 1\u20137 \u00b7 May 2026",
         doi="10.1109/ICIICS67880.2026.11483407", co="with R. A. M. Sait, B. Y. Vineetha, G. Priyadharshini, N. Sitanshu"),
    dict(b="ICIICS 2026", t="Predictive Mobility-Aware Cognitive Medium Access Control for Underwater Vehicle Communication",
         w="3rd Intl. Conf. on Integrated Intelligence and Communication Systems \u00b7 2026", doi=None,
         co="with P. Maderamitla, J. Ulaganathan, G. Bindu, K. Anguraj"),
    dict(b="ICBDML 2026", t="Harnessing Large Language Models for Next-Generation Artificial Intelligence Applications",
         w="2nd Intl. Conf. on Big Data and Machine Learning \u00b7 Feb 2026", doi=None,
         co="with N. Garg, S. Yangalasetty, T. Vangalapat, M. Kumari, M. Gupta"),
    dict(b="IEEE ICMNWC 2025", t="Lightweight Multi-Agent Q-learning Cross-Layer Optimization in Wireless IoT Networks for Balanced Energy Utilization",
         w="IEEE Intl. Conf. on Modelling, Networks and Wireless Communication \u00b7 Dec 2025", doi=None, co=None),
    dict(b="ICEASD 2024", t="Using AI Agents Built on LLMs for Building AGI",
         w="Intl. Conf. on Engineering Advancements and Sustainable Development, JETIR Vol. 11, Issue 5 \u00b7 May 2024",
         doi=None, co=None),
    dict(b="ICAMSnT 2023", t="The Alignment Problem in AI",
         w="Intl. Conf. on Advances in Modern Science and Technology \u00b7 Apr 2023", doi=None, co=None),
    dict(b="ICAPCTA 2022", t="Evaluating OpenAI Codex for Real-World Coding",
         w="Intl. Conf. on Advances in Parallel Computing Technologies & Applications \u00b7 Jan 2022", doi=None, co=None),
]

PATENTS = [
    ("A Bulk Silkworm Cocoon Inspection and Assisted Sorting Machine Using Shallow-Layer Conveyor Spreading", "Utility patent \u00b7 2026"),
    ("Adaptive Quality of Service Management System for Heterogeneous Internet of Things Networks", "Utility patent \u00b7 2026"),
    ("Ethical and Responsible AI: Bias Mitigation and Fairness in Machine Learning Models", "Utility patent \u00b7 2026"),
    ("Cloud-based Social Virtual Reality IoT Network with Device Synchronization Between Real and Virtual Worlds",
     "Published application 202441098998 \u00b7 filed 14 Dec 2024, published 20 Dec 2024"),
]


# ------------------------------------------------------------------ pieces ---
def doc_title(c, title, sub, size=17):
    p = para(c, first=True, after=1, line=1.05)
    run(p, title, size=size, color=INK, bold=True, font=HEAD)
    p = para(c, after=6)
    run(p, sub, size=7.6, color=ACC2, bold=True, caps=True, spacing=12)


def entry(c, b, title, where, doi=None, co=None, fill=TEAL, after=5):
    """One publication / patent entry: venue badge, title, provenance, DOI, co-authors."""
    p = para(c, after=0, line=1.14)
    badge(p, b, fill=fill, size=7)
    run(p, "  " + title, size=9, color=INK, bold=True)
    p = para(c, after=(0 if (doi or co) else after), line=1.1)
    p.paragraph_format.left_indent = Mm(1)
    run(p, where, size=8, color=MUT)
    if doi:
        p = para(c, after=(0 if co else after), line=1.1); p.paragraph_format.left_indent = Mm(1)
        run(p, "DOI ", size=7.2, color=MUT, bold=True)
        hyperlink(p, "https://doi.org/" + doi, doi, size=7.6, bold=False)
    if co:
        p = para(c, after=after, line=1.1); p.paragraph_format.left_indent = Mm(1)
        run(p, co, size=7.6, color=MUT, italic=True)


def stat_grid(c, stats, cols=2, w=Mm(25), big=16):
    t = c.add_table(rows=(len(stats) + cols - 1) // cols, cols=cols); t.autofit = False
    for i, s in enumerate(stats):
        cell = t.cell(i // cols, i % cols)
        K.set_cell_width(cell, w); K.set_cell_margins(cell, 0, 0, 0, 0)
        p = para(cell, first=True, after=0, line=1.0); run(p, s[0], size=big, color=ACC, bold=True, font=HEAD)
        p = para(cell, after=4, line=1.0); run(p, s[1], size=7, color=MUT)


# ============================================================ 1. PUBLICATIONS
def pubs_side1(c):
    name_block(c, PHOTO, "AI Researcher \u00b7 Educator")
    contact_block(c, D, IDS)
    heading(c, "At a glance")
    stat_grid(c, [("11", "peer-reviewed papers"), ("4", "patents & applications"),
                  ("4", "journal articles"), ("7", "conference papers")])
    heading(c, "Research Areas")
    for a in ("Agentic AI & LLM systems", "Spatio-temporal & attention models",
              "EEG & brain\u2013computer interfaces", "Cognitive IoT & cross-layer optimisation",
              "Responsible AI & fairness", "Anomaly detection"):
        p = para(c, after=1.6, line=1.1)
        p.paragraph_format.left_indent = Mm(3); p.paragraph_format.first_line_indent = Mm(-3)
        run(p, "\u25aa\u2002", size=6.5, color=ACC2); run(p, a, size=8.4)


def pubs_main1(c):
    doc_title(c, "Research Publications & Patents", TARGET)
    heading(c, "Research Profile")
    p = para(c, after=5, line=1.2)
    run(p, "My research asks how intelligent systems can be made reliable and useful outside the laboratory: how "
           "agentic systems built on large language models coordinate to solve multi-step problems, how deep models "
           "recover structure from noisy spatio-temporal signals, and how constrained networked devices share "
           "resources fairly. The work is deliberately applied. It has produced three utility patents and a "
           "published application, and it is shaped by eight years of building production software before I moved "
           "into academia. Current threads are the evaluation and reliability of LLM-based agents, attention "
           "architectures for EEG and motor-imagery decoding, cross-layer optimisation in heterogeneous IoT, and "
           "fairness in deployed machine-learning models.", size=9.3, color=BODY)

    heading(c, "Selected Contributions")
    for t in ("Multi-scale spatio-temporal attention for motor-imagery decoding \u2014 a network combining "
              "frequency\u2013spatial parallel attention with multi-scale temporal modelling (IEEE ICIICS 2026).",
              "Lightweight multi-agent Q-learning for wireless IoT \u2014 cross-layer optimisation that balances "
              "energy across heterogeneous devices without central coordination (IEEE ICMNWC 2025).",
              "Agentic LLM systems as a route to general intelligence \u2014 a collaborative multi-agent framework "
              "analysed against SayCan, ChatDev and Devin (INDJCST 2025).",
              "Fairness in deployed models \u2014 a patented framework for bias mitigation and fairness "
              "verification in machine-learning pipelines (2026)."):
        bullet(c, t, size=8.9, after=2)

    heading(c, "Journal Articles")
    for x in JOURNALS:
        entry(c, x["b"], x["t"], x["w"], x["doi"], x["co"])

    heading(c, "Patents")
    for i, (t, meta) in enumerate(PATENTS, 1):
        p = para(c, after=0, line=1.14)
        p.paragraph_format.left_indent = Mm(8); p.paragraph_format.first_line_indent = Mm(-8)
        p.paragraph_format.tab_stops.add_tab_stop(Mm(8))
        run(p, f"0{i}\t", size=10.5, color=ACC2, bold=True)
        run(p, t, size=9.2, color=INK, bold=(i <= 3))
        p = para(c, after=5, line=1.1); p.paragraph_format.left_indent = Mm(8)
        run(p, meta, size=7.8, color=MUT)


def pubs_side2(c):
    heading(c, "Doctoral Research", first=True)
    p = para(c, after=3, line=1.15)
    run(p, "Ph.D. in Computer Science & Engineering, Visvesvaraya Technological University (enrolled 2023), in "
           "applied deep learning and intelligent systems. Supervised through DSATM, Bengaluru.", size=8.4)
    heading(c, "Research Training")
    for n, w in (("Deep Learning \u2014 Elite", "NPTEL \u00b7 2025"), ("Blockchain Technology", "NPTEL \u00b7 2025"),
                 ("Data Analytics with Python", "NPTEL \u00b7 2024"),
                 ("Machine Learning & NLP", "Infosys Springboard FDP \u00b7 2025"),
                 ("Cyber Security", "AICTE \u00b7 VTU FDP \u00b7 2023")):
        p = para(c, after=0, keep=True); run(p, n, size=8.5, color=INK, bold=True)
        p = para(c, after=5); run(p, w, size=7.9, color=MUT)
    heading(c, "Reproducibility")
    p = para(c, after=0, line=1.15)
    run(p, "Research code and teaching systems are public at ", size=8.2, color=MUT)
    hyperlink(p, "https://github.com/" + D["github_user"], "github.com/" + D["github_user"], size=8.2)


def pubs_main2(c):
    heading(c, "Conference Papers", first=True)
    for x in CONFERENCES:
        entry(c, x["b"], x["t"], x["w"], x["doi"], x["co"])

    heading(c, "Research Supervision")
    p = para(c, after=3, line=1.18)
    run(p, "Final-year capstone and project supervision since 2021 (DSATM, City Engineering College, T. John "
           "Institute of Technology), across applied machine learning, data engineering and intelligent systems. "
           "Teams are expected to produce a working artefact, a written report and a defensible evaluation \u2014 "
           "several capstones have grown out of the open-source teaching systems listed overleaf.", size=9, color=BODY)

    heading(c, "Research Communication & Outreach")
    p = para(c, after=3, line=1.15)
    run(p, "I write and speak about AI for a general technical audience, which keeps my research legible outside "
           "the conference circuit.", size=8.9, color=BODY)
    for w in D["writing"]:
        p = para(c, after=2.5, line=1.1)
        p.paragraph_format.left_indent = Mm(3.4); p.paragraph_format.first_line_indent = Mm(-3.4)
        run(p, "\u25aa\u2002", size=7, color=ACC2); hyperlink(p, w["url"], w["title"], size=8.6)
    p = para(c, before=2, after=3, line=1.15)
    run(p, "Video explainers on the ", size=8.9, color=MUT)
    hyperlink(p, "https://www.youtube.com/@MaheshBasavaraj-q4g", "Echoes of AI", size=8.9)
    run(p, " YouTube channel.", size=8.9, color=MUT)

    heading(c, "Collaboration Interests")
    p = para(c, after=0, line=1.18)
    run(p, "Open to joint work on agent evaluation and LLM reliability, spatio-temporal modelling for biomedical "
           "signals, QoS in large-scale IoT deployments, and applied fairness auditing \u2014 particularly with "
           "industry partners who can supply real data and deployment constraints.", size=9, color=BODY)


def build_publications(out):
    doc = Document()
    two_col(doc, [(pubs_side1, pubs_main1), (pubs_side2, pubs_main2)])
    return finish(doc, out, title="Mahesh Basavaraj \u2014 Research Publications & Patents")


# ============================================================ 2. COVER LETTER
LETTER_BODY = [
    ("Dear Members of the Selection Committee,", "salutation"),

    ("I am applying for the position of Assistant Professor, Computer Science, at the UNSW Bengaluru campus. I "
     "am currently Assistant Professor in CSE (Data Science) at DSATM, Bengaluru, and a doctoral researcher at "
     "Visvesvaraya Technological University. My sixteen years split evenly between the two worlds this campus "
     "sits between: eight years building banking and capital-markets software at Infosys and Accenture, and "
     "eight years teaching computer science and publishing applied AI research. What I would most want you to "
     "know about me is simple: when a course needs software to make its ideas visible, I build it.", "p"),

    ("CPU Rush is the clearest example. My Operating Systems students do not read about scheduling, paging, "
     "deadlock and file systems \u2014 they predict how each behaves, commit to an answer, and a server grades "
     "every decision. Data Structure Design Studio turns data structures into a progressive, scored decision "
     "track under changing constraints. Game Forge generates single-file browser games, one concept each, from "
     "an editable design doctrine. ButterCloud Quiz asks students to choose the best stack and architecture for "
     "Eight such tools are public at github.com/maheshbasavaraj, most in daily "
     "classroom use, and every new course I teach keeps producing more \u2014 students decide, then watch the "
     "machine answer, which is where the learning actually happens.", "p"),

    ("I have taught object-oriented programming, data structures, operating systems, database systems and deep "
     "learning at undergraduate and postgraduate level, and I supervise final-year capstone projects end to "
     "end. At DSATM I designed and led a six-month Agentic AI immersion programme with Capabl and run faculty "
     "development programmes for colleagues. I am ready to adopt UNSW's myExperience framework and the "
     "Integrated Curriculum Framework, to work within standards-based, criterion-referenced assessment, and to "
     "complete Ready to Teach alongside colleagues from Kensington.", "p"),

    ("My research — Agentic AI, large language models and spatio-temporal deep learning — has produced three utility patents, a published patent application in synchronised social-VR IoT networks, and eleven peer-reviewed papers, including recent work on multi-scale spatio-temporal attention networks, lightweight multi-agent Q-learning for wireless IoT, and LLM-based agent systems as a route to general intelligence. The campus's focus on artificial intelligence, and its Centre of Excellence in AI and Innovation with Deloitte India, is where I want to work: applied research that produces both publications and capability.", "p"),

    ("Eight years in industry also taught me what employers expect from a computer science graduate. At Infosys I worked on Finacle, the universal banking platform; at Accenture I built capital-markets systems for international clients. I would bring that into work-integrated learning — industry projects, internship pipelines, and assessment built around real repositories, code review and delivery discipline.", "p"),

    ("UNSW Bengaluru offers a top-20 global degree with local consequence, and Progress for All is a mission I can teach to. The chance to help establish a new campus's standards rather than inherit them is rare. I am based in Bengaluru, ready from the first teaching period, and glad to demonstrate CPU Rush live or teach a class on any topic in the curriculum.", "p"),
]


def build_letter(out):
    doc = Document()
    s = single_page(doc, top=14, bottom=13, left=21, right=21)

    # letterhead
    p = para(doc, first=True, after=0, line=1.0)
    run(p, "Mahesh Basavaraj", size=19, color=INK, bold=True, font=HEAD)
    p = para(doc, after=1, line=1.0)
    run(p, "Assistant Professor, Computer Science & Engineering (Data Science) \u00b7 DSATM, Bengaluru", size=8.6, color=MUT)
    p = para(doc, after=3, line=1.0)
    run(p, "Bengaluru, India  \u00b7  ", size=8.6, color=MUT)
    hyperlink(p, "mailto:BMaheshbasavaraj@gmail.com", "BMaheshbasavaraj@gmail.com", size=8.6, bold=False)
    run(p, "  \u00b7  +91 98866 42542  \u00b7  ", size=8.6, color=MUT)
    hyperlink(p, "https://maheshbasavaraj.github.io/", "maheshbasavaraj.github.io", size=8.6, bold=False)
    run(p, "  \u00b7  ", size=8.6, color=MUT)
    hyperlink(p, "https://www.linkedin.com/in/mahesh-basavaraj-57189515/", "LinkedIn", size=8.6, bold=False)
    para_border_bottom(p, TEAL, sz=12, space=6)

    p = para(doc, before=7, after=6, line=1.0)
    run(p, DATE, size=9, color=MUT)

    for line, bold in (("The Selection Committee \u2014 Faculty Recruitment", True),
                       ("UNSW Bengaluru", True),
                       ("Embassy Manyata Business Park, Nagavara", False),
                       ("Bengaluru 560045, Karnataka, India", False)):
        p = para(doc, after=0, line=1.12)
        run(p, line, size=9.2, color=(INK if bold else MUT), bold=bold)

    p = para(doc, before=9, after=7, line=1.15)
    run(p, "Re:  ", size=9.6, color=ACC, bold=True)
    run(p, "Application for Assistant Professor, Computer Science \u2014 UNSW Bengaluru Campus", size=9.6, color=INK, bold=True)

    for text, kind in LETTER_BODY:
        p = para(doc, before=(0 if kind == "salutation" else 5), after=(6 if kind == "salutation" else 5), line=1.18)
        run(p, text, size=9.5, color=BODY)

    p = para(doc, before=7, after=1, line=1.2)
    run(p, "Yours sincerely,", size=9.5, color=BODY)
    p = para(doc, before=11, after=0, line=1.05)
    run(p, "Mahesh Basavaraj", size=12.5, color=INK, bold=True, font=HEAD)
    p = para(doc, after=0, line=1.05)
    run(p, "Assistant Professor, CSE (Data Science) \u00b7 DSATM, Bengaluru", size=8.4, color=MUT)

    p = para(doc, before=9, after=0, line=1.2)
    para_border_bottom(p, "DCDFE2", sz=4, space=4)
    p = para(doc, before=3, after=0, line=1.25)
    run(p, "Enclosures:  ", size=8.2, color=ACC, bold=True, caps=True, spacing=10)
    run(p, "Curriculum Vitae  \u00b7  Research Publications & Patents  \u00b7  Teaching Methodology", size=8.4, color=MUT)

    return finish(doc, out, title="Mahesh Basavaraj \u2014 Cover Letter, UNSW Bengaluru", shrink=False)


# ====================================================== 3. TEACHING METHODOLOGY
def teach_side1(c):
    name_block(c, PHOTO, "Educator \u00b7 AI Researcher")
    contact_block(c, D, IDS[:1] + IDS[3:])
    heading(c, "Teaching Snapshot")
    stat_grid(c, [("8", "years teaching CS"), ("6+", "courses taught"),
                  ("2", "programmes designed"), ("7", "certifications & FDPs")])
    heading(c, "Courses Taught")
    for n, w in (("Operating Systems", "UG \u00b7 core"), ("Data Structures & Algorithms", "UG \u00b7 core"),
                 ("Object-Oriented Programming (Java)", "UG \u00b7 core"),
                 ("Database Management Systems", "UG \u00b7 core"),
                 ("Deep Learning", "PG \u00b7 elective"),
                 ("Data Science & Analytics with Python", "UG/PG \u00b7 elective"),
                 ("Agentic AI (6-month immersion)", "Industry-partnered")):
        p = para(c, after=0, keep=True); run(p, n, size=8.5, color=INK, bold=True)
        p = para(c, after=4.5); run(p, w, size=7.9, color=MUT)
    heading(c, "Teaching Tools")
    p = para(c, after=0, line=1.15)
    run(p, "Built in-house and used in class: CPU Rush (OS), Data Structure Design Studio, Game Forge MCP, "
           "ButterCloud Quiz, MLOps CI/CD lab. All public on GitHub.", size=8.2, color=MUT)


def teach_main1(c):
    doc_title(c, "Teaching Methodology & Philosophy", TARGET)
    heading(c, "Why I Teach the Way I Do")
    p = para(c, after=5, line=1.2)
    run(p, "Students rarely fail computer science because it is too hard. They fail because the first encounter is "
           "abstract: memory is a diagram, a scheduler is a paragraph, a model is a black box. My job is to make "
           "the invisible observable. When a student can predict what the machine will do, and then watch it do "
           "something else, the correction is theirs \u2014 and it stays. I design for that moment deliberately, "
           "and I build the software that makes it possible at class size.", size=9.3, color=BODY)
    p = para(c, after=4, line=1.2)
    run(p, "The second half of the job is honesty about destination. Most of my students in Bengaluru will be "
           "employed as engineers within months of graduating. Theory has to be delivered in a form they can "
           "carry into a codebase \u2014 which is why my courses use real repositories, code review and delivery "
           "discipline, not only examinations.", size=9.3, color=BODY)

    heading(c, "Five Design Principles")
    principles = [
        ("Start from the misconception, not the syllabus.",
         "Before each topic I write down the wrong mental model students usually hold, and design the first "
         "activity to expose it. Deadlock is taught by making them schedule two processes that lock each other "
         "out, not by listing Coffman conditions first."),
        ("Make thinking visible, and therefore gradable.",
         "Every class has decision points where students commit to an answer \u2014 predict, then observe. "
         "Commitments are captured and scored, so I can see whether the room actually understood before the "
         "exam tells me."),
        ("Assess what matters, at the right stakes.",
         "Low-stakes practice every week with instant feedback; high-stakes judgement reserved for a capstone "
         "that looks like real work. Rubrics are criterion-referenced and published in advance, so a grade "
         "means \u201ccan do X\u201d rather than \u201cwas in the top 20%\u201d."),
        ("Treat generative AI as a professional instrument.",
         "Banning it teaches nothing and is unenforceable. I permit it with mandatory disclosure, assess the "
         "reasoning through oral defence, and set tasks where an unverified generated answer fails on its own "
         "evidence \u2014 which is exactly the skill employers now need."),
        ("Build for the whole room.",
         "Every course ships with a progressive track: scaffolded entry points for students who need them, "
         "extension problems for those who do not, and no single bottleneck activity that only one learning "
         "style can pass."),
    ]
    for i, (lead, body) in enumerate(principles, 1):
        p = para(c, after=2, line=1.16)
        p.paragraph_format.left_indent = Mm(6.4); p.paragraph_format.first_line_indent = Mm(-6.4)
        p.paragraph_format.tab_stops.add_tab_stop(Mm(6.4))
        run(p, f"{i}\t", size=10, color=ACC2, bold=True)
        run(p, lead + " ", size=9, color=INK, bold=True)
        run(p, body, size=9, color=BODY)


def teach_side2(c):
    heading(c, "Professional Development", first=True)
    for n, w in (("Deep Learning \u2014 Elite", "NPTEL \u00b7 Oct\u2013Dec 2025"),
                 ("Blockchain Technology", "NPTEL \u00b7 Oct\u2013Dec 2025"),
                 ("Data Analytics with Python", "NPTEL \u00b7 Jan\u2013Apr 2024"),
                 ("Machine Learning & NLP", "Infosys Springboard FDP \u00b7 Jun 2025"),
                 ("Cyber Security", "AICTE \u00b7 VTU FDP \u00b7 May 2023"),
                 ("Amazon Web Services", "AICTE FDP \u00b7 Aug 2022"),
                 ("The Science Behind Branding", "IISc Bangalore \u00b7 Dec 2025")):
        p = para(c, after=0, keep=True); run(p, n, size=8.4, color=INK, bold=True)
        p = para(c, after=4.5); run(p, w, size=7.8, color=MUT)
    heading(c, "Teaching Leadership")
    for t in ("Designed and led a 6-month Agentic AI immersion programme with Capabl (industry partner).",
              "Organised and delivered institutional Faculty Development Programmes.",
              "Supervised final-year capstone teams to working artefacts and defence."):
        p = para(c, after=3.5, line=1.12)
        p.paragraph_format.left_indent = Mm(3); p.paragraph_format.first_line_indent = Mm(-3)
        run(p, "\u25aa\u2002", size=6.5, color=ACC2); run(p, t, size=8.1)


def teach_main2(c):
    heading(c, "How I Run a Course", first=True)
    p = para(c, after=3, line=1.18)
    run(p, "Each week has the same rhythm, so students always know where the learning is being checked.", size=9, color=BODY)
    for lead, body in (
        ("Before class \u2014", "a two-question diagnostic on the platform tells me the real starting point, not the assumed one."),
        ("In class \u2014", "live coding with deliberate errors to catch, a short prediction round, then peer instruction: commit, compare, defend."),
        ("Lab \u2014", "each lab has checkable gates. Attendance is not evidence; a passing gate is."),
        ("Weekly \u2014", "a short low-stakes quiz with immediate feedback, a worked best answer, and a cumulative score that rewards consistency."),
        ("End of course \u2014", "a capstone in a real repository with code review, a demo and an oral defence."),
        ("Every term \u2014", "anonymous mid-semester feedback, one visible change made in response, and one assessment redesigned."),
    ):
        p = para(c, after=2.5, line=1.16)
        p.paragraph_format.left_indent = Mm(4.4); p.paragraph_format.first_line_indent = Mm(-4.4)
        run(p, "\u25aa\u2002", size=7, color=ACC2)
        run(p, lead + " ", size=9, color=INK, bold=True); run(p, body, size=9, color=BODY)

    heading(c, "Evidence: Tools I Have Built and Use")
    p = para(c, after=3, line=1.18)
    run(p, "I do not only design on paper. These are working systems, public on GitHub, that carry the "
           "principles above into a live classroom.", size=9, color=BODY)
    for name, repo, desc in (
        ("CPU Rush", "cpu-rush", "Socratic OS game \u2014 students predict scheduling, paging, deadlock and file-system behaviour; a server grades every move."),
        ("Data Structure Design Studio", "dsds", "Progressive scored decision track from arrays to hash maps, where constraints change and choices have consequences."),
        ("Game Forge (MCP)", "game_MCP", "Generates single-file browser games, each teaching one concept through play."),
        ("ButterCloud Quiz", "SWE_QUIZ", "Scenario platform: pick the best stack and architecture for realistic requirements."),
        ("MLOps CI/CD lab", "mlops-cicd", "Production-style ML API where every push runs tests and model-quality gates."),
    ):
        p = para(c, after=0, line=1.12)
        hyperlink(p, "https://github.com/maheshbasavaraj/" + repo, name, size=9)
        run(p, "  " + desc, size=8.7, color=BODY)

    heading(c, "Alignment with UNSW Bengaluru")
    p = para(c, after=4, line=1.18)
    run(p, "My practice maps directly onto how UNSW teaches. I design at program level rather than course level, "
           "which is the logic of the Integrated Curriculum Framework: learning outcomes, activities and "
           "assessment aligned before any content is written. I am comfortable with standards-based, "
           "criterion-referenced assessment, with the myExperience approach to course design, and with "
           "work-integrated learning as an assessed component rather than an add-on \u2014 my industry years make "
           "the partner conversations realistic. I expect to learn a great deal from the Ready to Teach "
           "programme and from colleagues in Sydney and Canberra, and I am keen to contribute back: the teaching "
           "systems above are reusable, and I would happily adapt them for UNSW courses and share them across "
           "the campus network.", size=9, color=BODY)
    p = para(c, after=3, line=1.18)
    run(p, "The new campus's AI Centre of Excellence with Deloitte India is a natural home for capstone work: "
           "students solving a partner's real problem, assessed on both engineering quality and research "
           "rigour. That is the kind of assessment I want to build here.", size=9, color=BODY)

    heading(c, "First Twelve Months")
    for lead, body in (
        ("Complete Ready to Teach", "and adopt myExperience and the Integrated Curriculum Framework in every course I take on."),
        ("Publish aligned course outlines", "with constructive alignment from program learning outcomes down to each assessment task."),
        ("Introduce one industry-linked assessment", "per course, co-designed with a partner and marked against a published rubric."),
        ("Contribute a teaching case study", "on AI-era assessment integrity to the Bengaluru and Sydney teaching communities."),
        ("Support the AI Centre of Excellence", "by connecting capstone projects to its research themes and supervising them to publication quality."),
    ):
        p = para(c, after=2.5, line=1.16)
        p.paragraph_format.left_indent = Mm(4.4); p.paragraph_format.first_line_indent = Mm(-4.4)
        run(p, "\u25aa\u2002", size=7, color=ACC2)
        run(p, lead + " \u2014 ", size=9, color=INK, bold=True); run(p, body, size=9, color=BODY)


def build_teaching(out):
    doc = Document()
    two_col(doc, [(teach_side1, teach_main1), (teach_side2, teach_main2)])
    return finish(doc, out, title="Mahesh Basavaraj \u2014 Teaching Methodology")


if __name__ == "__main__":
    for fn, name in ((build_publications, "Mahesh_Basavaraj_Research_Publications.docx"),
                     (build_letter, "Mahesh_Basavaraj_Cover_Letter_UNSW.docx"),
                     (build_teaching, "Mahesh_Basavaraj_Teaching_Methodology.docx")):
        print("wrote", fn(ROOT / name).name)
