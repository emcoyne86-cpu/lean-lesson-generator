"""
pipeline.py — Claude API calls for all 4 stages of the Lean Lesson Generator
"""

import json
import re
import anthropic
import pypdf
import io


# ── Helpers ───────────────────────────────────────────────────────────────────

def _client(api_key: str) -> anthropic.Anthropic:
    return anthropic.Anthropic(api_key=api_key)


def _extract_pdf_text(pdf_bytes: bytes) -> str:
    reader = pypdf.PdfReader(io.BytesIO(pdf_bytes))
    pages = []
    for page in reader.pages:
        text = page.extract_text()
        if text:
            pages.append(text)
    return "\n\n".join(pages)


def _parse_json(text: str) -> dict | list:
    """Extract JSON from a Claude response that may have surrounding prose."""
    # Try each strategy independently, catching decode errors along the way
    patterns = [
        r'```json\s*([\s\S]+?)\s*```',  # fenced json block
        r'```\s*([\s\S]+?)\s*```',       # any fenced block
        r'(\{[\s\S]+\})',                 # bare object
        r'(\[[\s\S]+\])',                 # bare array
    ]
    for pattern in patterns:
        match = re.search(pattern, text)
        if match:
            candidate = match.group(1).strip()
            try:
                return json.loads(candidate)
            except json.JSONDecodeError:
                continue
    raise ValueError(f"Could not parse JSON from response:\n{text[:500]}")


# ── Stage 1 — Lean Lesson ─────────────────────────────────────────────────────

STAGE1_PROMPT = """You are an expert curriculum designer specializing in lean lesson planning.

Analyze the lesson text below and produce a Lean Lesson document.

A Lean Lesson captures ONLY what teachers need to teach — no extraneous curriculum text. Organize the lesson into labeled steps using ONLY these tags:
  BACKGROUND — activating prior knowledge or setting context
  OBJECTIVE   — stating the learning goal to students
  KEY QUESTION — central inquiry or discussion prompt
  VOCAB       — vocabulary instruction
  MODEL       — teacher demonstrating a skill or process
  READ ALOUD  — teacher or student reading aloud
  PRACTICE    — guided or independent student practice
  PROTOCOL    — structured discussion or collaborative activity
  CHECK       — checking for understanding
  CLOSE       — closing synthesis or reflection
  HW          — homework

Group steps into sections: Opening, Work Time A, Work Time B, Closing, Homework (omit sections not present).

For each step:
  - tag: one tag from the list above
  - content: concise, action-oriented description of what happens (1-4 sentences, teacher voice)
  - timing: estimated minutes as an integer

Also extract:
  - lesson title
  - standards (CCSS or state standard codes)
  - objective (what students will be able to do, "I can..." format)
  - materials (specific texts, handouts, manipulatives)

Your response must be ONLY a valid JSON object — no explanation, no prose, no markdown fences. Start your response with { and end with }.

{
  "title": "...",
  "grade": "...",
  "subject": "...",
  "curriculum": "...",
  "unit": "...",
  "lesson": "...",
  "duration": "60 minutes",
  "standards": ["CCSS.ELA-LITERACY.RL.4.1"],
  "objective": "I can...",
  "materials": ["..."],
  "sections": [
    {
      "name": "Opening",
      "duration": "10 minutes",
      "steps": [
        {"tag": "BACKGROUND", "content": "...", "timing": 3}
      ]
    }
  ]
}

LESSON TEXT:
"""

def generate_stage1(
    pdf_bytes: bytes,
    grade: str,
    subject: str,
    curriculum: str,
    unit: str,
    lesson: str,
    api_key: str,
) -> dict:
    """
    Returns lesson_data dict with all lesson fields.
    Callers should pass this to build_lean_docx() to get the .docx bytes.
    """
    text = _extract_pdf_text(pdf_bytes)
    # Claude has a large context window; truncate only if enormous
    text = text[:80000]

    client = _client(api_key)
    response = client.messages.create(
        model="claude-sonnet-4-6",
        max_tokens=4096,
        messages=[
            {
                "role": "user",
                "content": (
                    STAGE1_PROMPT
                    + f"\n\nCurriculum: {curriculum}\nGrade: {grade}\nSubject: {subject}\n"
                    f"Unit: {unit}\nLesson: {lesson}\n\n"
                    + text
                ),
            }
        ],
    )
    data = _parse_json(response.content[0].text)

    # Ensure metadata fields are set (Claude may override with extracted values)
    data.setdefault("grade", grade)
    data.setdefault("subject", subject)
    data.setdefault("curriculum", curriculum)
    data.setdefault("unit", unit)
    data.setdefault("lesson", lesson)

    # Add sequential step numbers across all sections
    step_num = 1
    for section in data.get("sections", []):
        for step in section.get("steps", []):
            step["step_num"] = step_num
            step_num += 1

    return data


# ── Stage 2.1 — Scaffold Selector (no Claude needed) ─────────────────────────

def generate_selector_rows(lesson_data: dict) -> list[dict]:
    """
    Flatten all steps from lesson_data into selector rows.
    Returns list of dicts for xlsx_builder.
    """
    rows = []
    for section in lesson_data.get("sections", []):
        for step in section.get("steps", []):
            rows.append({
                "step_num": step.get("step_num", ""),
                "tag": step.get("tag", ""),
                "content": step.get("content", ""),
                "scaffold": "",   # teacher fills
                "all": "",
                "ell": "",
                "scd": "",
                "teacher_directive": "",
            })
    return rows


# ── Stage 2.2 — Scaffolded Lesson ────────────────────────────────────────────

SCAFFOLD_PROMPT = """You are an expert in differentiated instruction.

Generate scaffolds for the lesson steps below. For each step that needs a scaffold, choose the best strategy for each population.

Scaffold type options:
  ALL  → Sentence Frame, Question Stems, Anchor Chart, Graphic Organizer, Gesture/TPR, Partner Roles
  ELL  → Sentence Frame + Word Bank, Preview-View-Review, Home Language Bridge, TPR, Color-Coded Reference Card, Chunking + Glossing
  SCD  → Partially Completed Template, Chunking + Visual Support, Reduced-Choice Response, Step-by-Step Visual Directions, Graphic Organizer, Object/Image Anchor

For each scaffold, produce:
  - population: "ALL", "ELL", or "SCD"
  - type: scaffold type name
  - content: the actual scaffold (e.g., the sentence frame text, anchor chart content, word bank list) — be specific and use lesson vocabulary
  - teacher_note: brief implementation tip (1-2 sentences)

Your response must be ONLY a valid JSON object — no explanation, no prose, no markdown fences. Start your response with { and end with }.

{
  "scaffolded_steps": [
    {
      "step_num": 1,
      "scaffolds": [
        {
          "population": "ALL",
          "type": "Sentence Frame",
          "content": "I notice that _____ because _____.",
          "teacher_note": "Display on board; model once before releasing students."
        }
      ]
    }
  ]
}

Only include steps that have scaffolds. Omit steps with no scaffolds.

LESSON CONTEXT:
Title: {title}
Grade: {grade}  Subject: {subject}

STEPS NEEDING SCAFFOLDS:
{steps_json}
"""

def generate_stage2(
    lesson_data: dict,
    selector_rows: list[dict],
    api_key: str,
) -> dict:
    """
    Returns lesson_data enriched with scaffolds on applicable steps.
    """
    # Find steps that need scaffolds
    scaffold_steps = [r for r in selector_rows if r.get("scaffold", "").lower() == "yes"]
    if not scaffold_steps:
        return lesson_data  # nothing to scaffold

    # Build step context for Claude
    steps_context = []
    for r in scaffold_steps:
        pops = []
        if r.get("all", "").upper() == "Y":
            pops.append("ALL")
        if r.get("ell", "").upper() == "Y":
            pops.append("ELL")
        if r.get("scd", "").upper() == "Y":
            pops.append("SCD")
        steps_context.append({
            "step_num": r["step_num"],
            "tag": r["tag"],
            "content": r["content"],
            "populations": pops,
            "teacher_directive": r.get("teacher_directive", ""),
        })

    prompt = SCAFFOLD_PROMPT.format(
        title=lesson_data.get("title", ""),
        grade=lesson_data.get("grade", ""),
        subject=lesson_data.get("subject", ""),
        steps_json=json.dumps(steps_context, indent=2),
    )

    client = _client(api_key)
    response = client.messages.create(
        model="claude-sonnet-4-6",
        max_tokens=4096,
        messages=[{"role": "user", "content": prompt}],
    )
    scaffold_data = _parse_json(response.content[0].text)

    # Build lookup: step_num → scaffolds
    scaffold_map = {}
    for entry in scaffold_data.get("scaffolded_steps", []):
        scaffold_map[int(entry["step_num"])] = entry.get("scaffolds", [])

    # Deep-copy lesson_data and inject scaffolds
    import copy
    enriched = copy.deepcopy(lesson_data)
    for section in enriched.get("sections", []):
        for step in section.get("steps", []):
            snum = step.get("step_num")
            if snum in scaffold_map:
                step["scaffolds"] = scaffold_map[snum]

    return enriched


# ── Stage 3 — Student Slides ──────────────────────────────────────────────────

SLIDES_PROMPT = """You are an expert at creating student-facing lesson slides for {grade} students.

Create slides for this {subject} lesson. Slides must be:
- Written FOR students, not teachers ("You will..." not "Students will...")
- Concise — each slide makes one point
- Engaging and age-appropriate for Grade {grade}

Generate these slide types in order:
1. title     — lesson title card
2. target    — "I can..." learning target
3. agenda    — brief overview of the lesson (3-5 bullet points)
4. content   — one slide per major lesson step (combine small steps if logical)
5. scaffold  — one slide per scaffold (student-facing visual: sentence frame, anchor chart, etc.)
6. closing   — exit ticket or closing reflection prompt

For each slide:
  - type: one of the types above
  - title: short heading (max 7 words)
  - bullets: list of 1-5 concise bullet strings (or empty list for title slides)
  - featured_text: optional large-display text (e.g., a sentence frame or discussion prompt — leave "" if none)
  - bg_color: one of "CORAL", "TEAL", "PURPLE", "AMBER", "GREEN", "PEACH", "WHITE"
  - speaker_notes: brief teacher note for this slide

Your response must be ONLY a valid JSON array — no explanation, no prose, no markdown fences. Start your response with [ and end with ].

[
  {{
    "type": "title",
    "title": "The Red Wheelbarrow",
    "bullets": [],
    "featured_text": "",
    "bg_color": "TEAL",
    "speaker_notes": "Welcome students and introduce the poem."
  }},
  ...
]

LESSON DATA:
{lesson_json}
"""

def generate_stage3(lesson_data: dict, api_key: str) -> list[dict]:
    """
    Returns a list of slide dicts for pptx_builder.
    """
    prompt = SLIDES_PROMPT.format(
        grade=lesson_data.get("grade", ""),
        subject=lesson_data.get("subject", ""),
        lesson_json=json.dumps(lesson_data, indent=2)[:12000],  # keep prompt manageable
    )

    client = _client(api_key)
    response = client.messages.create(
        model="claude-sonnet-4-6",
        max_tokens=6000,
        messages=[{"role": "user", "content": prompt}],
    )
    slides = _parse_json(response.content[0].text)
    return slides if isinstance(slides, list) else slides.get("slides", [])
