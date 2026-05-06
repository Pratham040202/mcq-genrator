from flask import Flask, render_template, request, send_file
from groq import Groq
from openpyxl import Workbook
import zipfile
import xml.etree.ElementTree as ET
import os
import time
import re

app = Flask(__name__)

UPLOAD_FOLDER = "uploads"
os.makedirs(UPLOAD_FOLDER, exist_ok=True)

# 🔐 Your API key
client = Groq(api_key="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")

# ==============================
# CLEANUP
# ==============================
def clean_old_files():
    for file in os.listdir(UPLOAD_FOLDER):
        os.remove(os.path.join(UPLOAD_FOLDER, file))

    for file in ["easy_mcq.xlsx", "medium_mcq.xlsx", "hard_mcq.xlsx"]:
        if os.path.exists(file):
            os.remove(file)

clean_old_files()

# ==============================
# EXTRACT TEXT FROM DOCX
# ==============================
def extract_text(path):
    text = ""
    with zipfile.ZipFile(path, "r") as z:
        xml_content = z.read("word/document.xml")
        tree = ET.fromstring(xml_content)

        for elem in tree.iter():
            if elem.text:
                text += elem.text + " "
    return text

# ==============================
# GENERATE 50 MCQs (5 × 10)
# ==============================
def generate_mcqs(chunk, difficulty):
    full_output = ""

    for i in range(5):  # 5 batches → 50 questions
        retries = 3

        while retries > 0:
            try:
                prompt = f"""
Generate EXACTLY 10 MCQs.

Difficulty: {difficulty}

Rules:
- Questions must be correct
- 4 options (A-D)
- Only one correct answer

Format STRICTLY:

Q1. Question
A. Option
B. Option
C. Option
D. Option
Answer: A

Text:
{chunk}
"""

                response = client.chat.completions.create(
                    model="llama-3.1-8b-instant",
                    messages=[{"role": "user", "content": prompt}]
                )

                full_output += response.choices[0].message.content + "\n\n"

                time.sleep(3)  # prevent rate limit
                break

            except:
                retries -= 1
                time.sleep(4)

    return full_output

# ==============================
# PARSE MCQs → STRUCTURED
# ==============================
def parse_mcqs(text):
    questions = []
    current = {}

    for line in text.split("\n"):
        line = line.strip()

        if not line:
            continue

        if line.startswith("Q"):
            if current:
                questions.append(current)
                current = {}

            current["question"] = re.sub(r"Q\d+[\.\)]?\s*", "", line)

        elif line.startswith("A."):
            current["A"] = line[2:].strip()

        elif line.startswith("B."):
            current["B"] = line[2:].strip()

        elif line.startswith("C."):
            current["C"] = line[2:].strip()

        elif line.startswith("D."):
            current["D"] = line[2:].strip()

        elif "Answer:" in line:
            current["answer"] = line.split(":")[-1].strip()

    if current:
        questions.append(current)

    return questions[:50]  # ensure max 50

# ==============================
# SAVE TO EXCEL
# ==============================
def save_excel(filename, mcq_text):
    wb = Workbook()
    ws = wb.active
    ws.title = "MCQ Paper"

    # HEADER (only once)
    headers = ["Sr No.", "Question", "Option A", "Option B", "Option C", "Option D", "Answer"]
    ws.append(headers)

    questions = parse_mcqs(mcq_text)

    for i, q in enumerate(questions, start=1):
        ws.append([
            i,
            q.get("question", ""),
            q.get("A", ""),
            q.get("B", ""),
            q.get("C", ""),
            q.get("D", ""),
            q.get("answer", "")
        ])

    wb.save(filename)

# ==============================
# ROUTES
# ==============================
@app.route("/", methods=["GET", "POST"])
def index():
    if request.method == "POST":
        clean_old_files()

        file = request.files["file"]
        path = os.path.join(UPLOAD_FOLDER, file.filename)
        file.save(path)

        text = extract_text(path)
        chunks = [text[i:i+1500] for i in range(0, 4500, 1500)]

        print("Generating Easy...")
        easy = generate_mcqs(chunks[0], "easy")

        print("Generating Medium...")
        medium = generate_mcqs(chunks[1], "medium")

        print("Generating Hard...")
        hard = generate_mcqs(chunks[2], "hard")

        save_excel("easy_mcq.xlsx", easy)
        save_excel("medium_mcq.xlsx", medium)
        save_excel("hard_mcq.xlsx", hard)

        print("Done")

        return render_template("index.html", ready=True)

    return render_template("index.html", ready=False)

@app.route("/download/<level>")
def download(level):
    return send_file(f"{level}_mcq.xlsx", as_attachment=True)

# ==============================
# RUN
# ==============================
if __name__ == "__main__":
    app.run(debug=True)