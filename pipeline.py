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
    patterns = [
        r'```json\s*([\s\S]+?)\s*```',  # fenced json block
        r'```\s*([\s\S]+?)\s*```',       # any fenced block
        r'(\{[\s\S]+\})',                 # bare object
        r'(\[[\s\S]+\])',                 # bare array
    ]
    candidates = []
    for pattern in patterns:
        match = re.search(pattern, text)
        if match:
            candidates.append(match.group(1).strip())

    for candidate in candidates:
        # Try as-is
        try:
            return json.loads(candidate)
        except json.JSONDecodeError:
            pass
        # Try wrapping in {} (Claude sometimes omits outer braces)
        try:
            return json.loads('{' + candidate + '}')
        except json.JSONDecodeError:
            pass
        # Try wrapping in [] (Claude sometimes omits outer array brackets)
        try:
            return json.loads('[' + candidate + ']')
        except json.JSONDecodeError:
            pass

    # Last resort: try wrapping the whole response
    stripped = text.strip()
    for wrapped in ['{' + stripped + '}', '[' + stripped + ']']:
        try:
            return json.loads(wrapped)
        except json.JSONDecodeError:
            pass

    raise ValueError(f"Could not parse JSON from Claude response. First 300 chars:\n{text[:300]}")


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
  - anticipated_response: for KEY QUESTION steps only — brief expected student answer (1-2 sentences); omit for all other steps

Also extract:
  - lesson title
  - standards (CCSS or state standard codes)
  - objective (what students will be able to do, "I can..." format — include all targets)
  - key_vocabulary (list of key vocabulary words or phrases introduced in the lesson)
  - materials (specific texts, handouts, manipulatives)

Respond with ONLY a JSON object wrapped in ```json``` fences. No other text.

```json
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
  "key_vocabulary": ["word1", "word2"],
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
                "all": "",
                "ell": "",
                "scd": "",
                "teacher_directive": "",
            })
    return rows


# ── Stage 2.2 — Scaffolded Lesson ────────────────────────────────────────────

def generate_stage2(
    lesson_data: dict,
    selector_rows: list[dict],
    api_key: str,
) -> dict:
    """Returns lesson_data enriched with scaffolds on applicable steps."""
    scaffold_steps = [
        r for r in selector_rows
        if any(r.get(p, "").upper() == "Y" for p in ["all", "ell", "scd"])
    ]
    if not scaffold_steps:
        return lesson_data

    # Build plain-text step descriptions for the prompt
    lines = []
    for r in scaffold_steps:
        pops = [p for p in ["ALL", "ELL", "SCD"] if r.get(p.lower(), "").upper() == "Y"]
        lines.append(
            f"Step {r['step_num']} [{r['tag']}] for {', '.join(pops)}:\n{r['content']}"
        )
    steps_text = "\n\n".join(lines)

    prompt = (
        "You are an expert in differentiated instruction.\n\n"
        f"Lesson: {lesson_data.get('title','')} | Grade {lesson_data.get('grade','')} | {lesson_data.get('subject','')}\n\n"
        "Generate one scaffold per population per step listed below.\n"
        "Scaffold types for ALL: Sentence Frame, Question Stems, Anchor Chart, Graphic Organizer, Partner Roles\n"
        "Scaffold types for ELL: Sentence Frame + Word Bank, Preview-View-Review, TPR, Home Language Bridge, Color-Coded Card\n"
        "Scaffold types for SCD: Partially Completed Template, Chunking + Visual Support, Reduced-Choice Response, Step-by-Step Directions\n\n"
        "Return a JSON array where each element has: step_num (integer), population (ALL/ELL/SCD), "
        "type (scaffold name), content (the actual scaffold text), teacher_note (one sentence).\n\n"
        "STEPS:\n" + steps_text + "\n\n"
        "Return ONLY the JSON array, no other text."
    )

    client = _client(api_key)
    response = client.messages.create(
        model="claude-sonnet-4-6",
        max_tokens=4096,
        messages=[{"role": "user", "content": prompt}],
    )

    raw = response.content[0].text
    result = _parse_json(raw)

    # Normalise: accept both a bare list and {"scaffolds": [...]}
    if isinstance(result, dict):
        for key in ("scaffolds", "scaffolded_steps", "items"):
            if key in result:
                result = result[key]
                break
        else:
            result = list(result.values())[0] if result else []

    # Group by step_num
    scaffold_map: dict = {}
    for item in (result or []):
        snum = int(item.get("step_num", 0))
        scaffold_map.setdefault(snum, []).append({
            "population":   item.get("population", "ALL"),
            "type":         item.get("type", ""),
            "content":      item.get("content", ""),
            "teacher_note": item.get("teacher_note", ""),
        })

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

Respond with ONLY a JSON array wrapped in ```json``` fences. No other text.

```json
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
