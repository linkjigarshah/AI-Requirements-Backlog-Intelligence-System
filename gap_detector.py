import os
from groq import Groq
from docx import Document
from dotenv import load_dotenv

load_dotenv()
groq_client = Groq(api_key=os.environ.get("GROQ_API_KEY"))

# ── Read requirements from docx ──────────────────────────────
def read_requirements_from_docx(filename):
    doc = Document(filename)
    requirements = []
    for paragraph in doc.paragraphs:
        text = paragraph.text.strip()
        style_name = paragraph.style.name if paragraph.style else ""
        if text and not style_name.startswith("Heading"):
            requirements.append(text)
    return requirements

# ── Compress requirements to reduce token usage ───────────────
def compress_requirements(requirements):
    compressed = []
    for req in requirements:
        r = req
        filler = [
            "The system shall ", "The system should ", "The system will ",
            "The system must ", "The system is required to ",
            "It is required that the system ", "The application shall ",
            "Users shall be able to ", "Users should be able to ",
            "The platform shall ", "The solution shall ",
        ]
        for f in filler:
            r = r.replace(f, "")
        if r:
            r = r[0].upper() + r[1:]
        skip_keywords = ["Note:", "Note from", "TBD", "To be confirmed", "?"]
        if any(r.startswith(k) for k in skip_keywords):
            continue
        if len(r) < 20:
            continue
        compressed.append(r.strip())
    seen = set()
    unique = []
    for r in compressed:
        if r not in seen:
            seen.add(r)
            unique.append(r)
    return unique

# ── Load and compress ─────────────────────────────────────────
print("Reading requirements from requirements.docx...")
requirements = read_requirements_from_docx("requirements.docx")
requirements = compress_requirements(requirements)
print(f"Found {len(requirements)} requirements after compression")
print("-" * 50)

numbered = "\n".join([f"{i+1}. {r}" for i, r in enumerate(requirements)])

prompt = f"""
You are a senior Business Analyst and QA expert performing a requirements review.

Analyze each requirement below and identify any of these gap types:
- VAGUE: The requirement is ambiguous or unclear
- MISSING: Something important is not defined
- UNTESTABLE: The requirement cannot be objectively verified
- CONFLICT: The requirement contradicts another requirement

For each gap found, respond in this exact format:

Requirement #: [number]
Requirement: [the original requirement text]
Gap Type: [VAGUE / MISSING / UNTESTABLE / CONFLICT]
Severity: [HIGH / MEDIUM / LOW]
Issue: [one sentence describing the problem]
Recommendation: [one sentence on how to fix it]

---

If a requirement has no gaps, write:
Requirement #: [number]
Status: PASS

---

Requirements to analyze:
{numbered}
"""

print("Analyzing requirements for gaps...")

response = groq_client.chat.completions.create(
    model="llama-3.1-8b-instant",
    messages=[{"role": "user", "content": prompt}]
)

result = response.choices[0].message.content
print(result)

with open("gap_analysis_output.txt", "w") as f:
    f.write(result)

print("-" * 50)
print("Gap analysis saved to gap_analysis_output.txt")