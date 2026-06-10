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
You are a senior Business Analyst. Based on the requirements below, generate user stories in this exact format:

As a [user type], I want to [action], so that [benefit].

Acceptance Criteria:
- [criterion 1]
- [criterion 2]
- [criterion 3]

Generate one user story for each requirement. Be specific and professional.

Requirements:
{numbered}
"""

print("Sending requirements to AI...")

response = groq_client.chat.completions.create(
    model="llama-3.1-8b-instant",
    messages=[{"role": "user", "content": prompt}]
)

result = response.choices[0].message.content
print(result)

with open("user_stories_output.txt", "w") as f:
    f.write(result)

print("-" * 50)
print("Output saved to user_stories_output.txt")