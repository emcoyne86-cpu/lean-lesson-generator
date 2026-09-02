"""
Lean Lesson Generator — Streamlit App
Run locally:  streamlit run app.py
"""

import streamlit as st
import io
from pipeline import generate_stage1, generate_selector_rows, generate_stage2, generate_stage3
from builders import (
    build_lean_docx, build_scaffolded_docx,
    build_selector_xlsx, read_selector_xlsx,
    build_slides_pptx,
)

# ── Page config ───────────────────────────────────────────────────────────────
st.set_page_config(
    page_title="Lean Lesson Generator",
    page_icon="📚",
    layout="wide",
    initial_sidebar_state="expanded",
)

st.markdown("""
<style>
/* Overall font */
html, body, [class*="css"] { font-family: 'Nunito', sans-serif; }

/* Stage tab pills */
div[data-testid="stTabs"] button {
    font-weight: 700;
    font-size: 0.9rem;
}

/* Download button */
div.stDownloadButton > button {
    background-color: #1F3864;
    color: white;
    font-weight: 700;
    border-radius: 8px;
    padding: 0.5rem 1.2rem;
}
div.stDownloadButton > button:hover { background-color: #4ECEC8; }

/* Primary button */
div.stButton > button[kind="primary"] {
    background-color: #FF6B6B;
    color: white;
    font-weight: 700;
    border-radius: 8px;
}
div.stButton > button[kind="primary"]:hover { background-color: #e05252; }

.badge {
    display: inline-block;
    padding: 2px 10px;
    border-radius: 12px;
    font-size: 0.78rem;
    font-weight: 700;
    margin-right: 4px;
}
</style>
""", unsafe_allow_html=True)

# ── Session state init ────────────────────────────────────────────────────────
for key in ["lesson_data", "lean_docx", "selector_xlsx", "scaffolded_data",
            "scaffolded_docx", "slides_pptx", "api_key"]:
    if key not in st.session_state:
        st.session_state[key] = None

# ── Sidebar ───────────────────────────────────────────────────────────────────
with st.sidebar:
    st.image("https://img.icons8.com/color/96/book.png", width=60)
    st.title("Lean Lesson Generator")
    st.caption("AI-powered differentiated lesson planning")
    st.divider()

    api_key_input = st.text_input(
        "Anthropic API Key",
        type="password",
        value=st.session_state.api_key or "",
        help="Required to generate lessons. Get one at console.anthropic.com",
        placeholder="sk-ant-..."
    )
    if api_key_input:
        st.session_state.api_key = api_key_input
        st.success("✓ API key saved")
    else:
        st.warning("Enter your API key to get started.")

    st.divider()
    st.caption("**Pipeline stages:**")
    stages = {
        "1 · Lean Lesson":       st.session_state.lean_docx is not None,
        "2 · Scaffold Selector": st.session_state.selector_xlsx is not None,
        "3 · Scaffolded Lesson": st.session_state.scaffolded_docx is not None,
        "4 · Student Slides":    st.session_state.slides_pptx is not None,
    }
    for label, done in stages.items():
        icon = "✅" if done else "⬜"
        st.caption(f"{icon}  {label}")

    st.divider()
    if st.button("🔄 Start Over", use_container_width=True):
        for key in ["lesson_data", "lean_docx", "selector_xlsx",
                    "scaffolded_data", "scaffolded_docx", "slides_pptx"]:
            st.session_state[key] = None
        st.rerun()

# ── Main content ──────────────────────────────────────────────────────────────
st.markdown("## 📚 Lean Lesson Generator")
st.caption("Upload a lesson PDF and let AI do the heavy lifting — one stage at a time.")

tab1, tab2, tab3, tab4 = st.tabs([
    "Stage 1 · Lean Lesson",
    "Stage 2 · Scaffold Selector",
    "Stage 3 · Scaffolded Lesson",
    "Stage 4 · Student Slides",
])


# ─────────────────────────────────────────────────────────────────────────────
# STAGE 1 — Lean Lesson
# ─────────────────────────────────────────────────────────────────────────────
with tab1:
    st.subheader("Step 1 · Generate Lean Lesson")
    st.caption("Upload a lesson PDF and enter the lesson details. Claude will extract the essential steps and format them into a clean, printable Word document.")

    col1, col2 = st.columns(2)
    with col1:
        grade = st.selectbox("Grade", ["K", "1", "2", "3", "4", "5", "6", "7", "8"], index=3)
        subject = st.selectbox("Subject", ["ELA", "Math", "Science", "Social Studies"])
        curriculum = st.text_input("Curriculum / Publisher", placeholder="e.g., EL Education, Wonders, Illustrative Math")
    with col2:
        unit = st.text_input("Unit", placeholder="e.g., Module 1, Unit 1")
        lesson = st.text_input("Lesson", placeholder="e.g., Lesson 3")
        pdf_file = st.file_uploader("Lesson PDF", type=["pdf"], label_visibility="visible")

    st.divider()

    can_generate = bool(pdf_file and st.session_state.api_key and curriculum)
    if not st.session_state.api_key:
        st.info("Enter your API key in the sidebar to continue.")
    elif not curriculum:
        st.info("Fill in the curriculum name above.")
    elif not pdf_file:
        st.info("Upload a lesson PDF above.")

    if st.button("✨ Generate Lean Lesson", type="primary", disabled=not can_generate):
        with st.spinner("Claude is reading the lesson… this takes about 30–60 seconds."):
            try:
                lesson_data = generate_stage1(
                    pdf_file.read(), grade, subject, curriculum, unit, lesson,
                    st.session_state.api_key,
                )
                lean_bytes = build_lean_docx(lesson_data)
                selector_rows = generate_selector_rows(lesson_data)
                selector_bytes = build_selector_xlsx(selector_rows, lesson_data)

                st.session_state.lesson_data     = lesson_data
                st.session_state.lean_docx       = lean_bytes
                st.session_state.selector_xlsx   = selector_bytes
                st.success("✅ Lean lesson created!")
            except Exception as e:
                st.error(f"Something went wrong: {e}")

    if st.session_state.lean_docx:
        title_slug = (st.session_state.lesson_data or {}).get("title", "lesson").replace(" ", "_")[:40]
        st.success("Lean lesson is ready to download.")
        st.download_button(
            "⬇️ Download Lean Lesson (.docx)",
            data=st.session_state.lean_docx,
            file_name=f"{title_slug}_Lean.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        )
        with st.expander("Preview lesson structure"):
            ld = st.session_state.lesson_data
            st.markdown(f"**{ld.get('title','')}**  ·  Grade {ld.get('grade','')}  ·  {ld.get('subject','')}  ·  {ld.get('duration','')}")
            st.markdown(f"*{ld.get('objective','')}*")
            total = 0
            for sec in ld.get("sections", []):
                st.markdown(f"**{sec['name']}** — {sec.get('duration','')}")
                for step in sec.get("steps", []):
                    st.markdown(f"  `[{step['tag']}]` {step['content']} _{step.get('timing','')} min_")
                    total += step.get("timing", 0)
            st.caption(f"Total estimated time: {total} min")


# ─────────────────────────────────────────────────────────────────────────────
# STAGE 2 — Scaffold Selector
# ─────────────────────────────────────────────────────────────────────────────
with tab2:
    st.subheader("Step 2 · Fill Out the Scaffold Selector")
    st.caption("Download the selector, decide which steps need scaffolds and for which students, then upload the completed file.")

    if not st.session_state.selector_xlsx:
        st.info("Complete Stage 1 first to generate the selector.")
    else:
        st.markdown("""
**How to fill it out:**
1. Open the downloaded Excel file
2. For each step, set **Scaffold?** to `Yes` or `No`
3. If Yes, mark `Y` under **ALL**, **ELL**, and/or **SCD** — any combination
4. Add optional notes in **Teacher Directive**
5. Save and upload the file below
""")
        title_slug = (st.session_state.lesson_data or {}).get("title", "lesson").replace(" ", "_")[:40]
        st.download_button(
            "⬇️ Download Scaffold Selector (.xlsx)",
            data=st.session_state.selector_xlsx,
            file_name=f"{title_slug}_Scaffold_Selector.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )

        st.divider()
        filled_file = st.file_uploader("Upload completed selector (.xlsx)", type=["xlsx"])
        if filled_file:
            st.success("Filled selector uploaded — proceed to Stage 3.")
            st.session_state["filled_selector_bytes"] = filled_file.read()


# ─────────────────────────────────────────────────────────────────────────────
# STAGE 3 — Scaffolded Lesson
# ─────────────────────────────────────────────────────────────────────────────
with tab3:
    st.subheader("Step 3 · Generate Scaffolded Lesson")
    st.caption("Claude will generate differentiated scaffolds for each step you selected, and weave them into the lesson document.")

    if not st.session_state.lesson_data:
        st.info("Complete Stage 1 first.")
    elif not st.session_state.get("filled_selector_bytes"):
        st.info("Upload your completed scaffold selector in Stage 2 first.")
    else:
        can_scaffold = bool(st.session_state.api_key)
        if not can_scaffold:
            st.info("Enter your API key in the sidebar.")

        if st.button("✨ Generate Scaffolded Lesson", type="primary", disabled=not can_scaffold):
            with st.spinner("Generating scaffolds… Claude is designing supports for your students."):
                try:
                    selector_rows = read_selector_xlsx(st.session_state["filled_selector_bytes"])
                    scaffolded_data = generate_stage2(
                        st.session_state.lesson_data,
                        selector_rows,
                        st.session_state.api_key,
                    )
                    scaffolded_bytes = build_scaffolded_docx(scaffolded_data)
                    st.session_state.scaffolded_data  = scaffolded_data
                    st.session_state.scaffolded_docx  = scaffolded_bytes
                    st.success("✅ Scaffolded lesson created!")
                except Exception as e:
                    st.error(f"Something went wrong: {e}")

        if st.session_state.scaffolded_docx:
            title_slug = (st.session_state.lesson_data or {}).get("title", "lesson").replace(" ", "_")[:40]
            st.success("Scaffolded lesson is ready.")
            st.download_button(
                "⬇️ Download Scaffolded Lesson (.docx)",
                data=st.session_state.scaffolded_docx,
                file_name=f"{title_slug}_Scaffolded.docx",
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            )

            # Preview scaffolds added
            with st.expander("Preview scaffolds added"):
                for sec in st.session_state.scaffolded_data.get("sections", []):
                    for step in sec.get("steps", []):
                        if step.get("scaffolds"):
                            st.markdown(f"**Step {step.get('step_num','')} [{step['tag']}]** — {step['content'][:80]}...")
                            for s in step["scaffolds"]:
                                pop_color = {"ALL": "🟢", "ELL": "🔵", "SCD": "🟣"}.get(s["population"], "⚪")
                                st.markdown(f"  {pop_color} **{s['population']}** · *{s['type']}* — {s['content'][:100]}")


# ─────────────────────────────────────────────────────────────────────────────
# STAGE 4 — Student Slides
# ─────────────────────────────────────────────────────────────────────────────
with tab4:
    st.subheader("Step 4 · Generate Student-Facing Slides")
    st.caption("Claude will create a colorful, student-facing slide deck from the lesson — including scaffold slides for supported steps.")

    source_data = st.session_state.scaffolded_data or st.session_state.lesson_data
    if not source_data:
        st.info("Complete Stage 1 first. You can generate slides from the lean lesson or the scaffolded version.")
    else:
        using = "scaffolded lesson" if st.session_state.scaffolded_data else "lean lesson (no scaffolds)"
        st.caption(f"Slides will be generated from the **{using}**.")

        can_slides = bool(st.session_state.api_key)
        if not can_slides:
            st.info("Enter your API key in the sidebar.")

        if st.button("✨ Generate Student Slides", type="primary", disabled=not can_slides):
            with st.spinner("Designing slides… Claude is building your student deck."):
                try:
                    slides_data  = generate_stage3(source_data, st.session_state.api_key)
                    slides_bytes = build_slides_pptx(slides_data, source_data)
                    st.session_state.slides_pptx = slides_bytes
                    st.success(f"✅ {len(slides_data)} slides created!")
                except Exception as e:
                    st.error(f"Something went wrong: {e}")

        if st.session_state.slides_pptx:
            title_slug = source_data.get("title", "lesson").replace(" ", "_")[:40]
            st.success("Slides are ready.")
            st.download_button(
                "⬇️ Download Student Slides (.pptx)",
                data=st.session_state.slides_pptx,
                file_name=f"{title_slug}_Slides.pptx",
                mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
            )
            st.caption("Open in PowerPoint or Google Slides. Scaffold slides have teacher notes in the notes panel.")
