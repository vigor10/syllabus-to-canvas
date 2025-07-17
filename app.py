
import re
from docx import Document

def extract_course_schedule_from_docx(filepath):
    doc = Document(filepath)
    lines = [para.text.strip() for para in doc.paragraphs if para.text.strip()]

    schedule_index = None
    for idx, line in enumerate(lines):
        if re.match(r"^#\\s*$", line):
            schedule_index = idx
            break

    if schedule_index is None:
        return []

    course_schedule_data = []
    for idx in range(schedule_index + 1, len(lines)):
        if lines[idx].strip() == "":
            break
        course_schedule_data.append(lines[idx])

    structured_sessions = []
    session_pattern = re.compile(r"^(?P<session>\\d+)\\s+(?P<date>[A-Za-z]{3},\\s*\\d+/\\d+)\\s+(?P<topics>.+)$")
    current = {}
    for entry in course_schedule_data:
        match = session_pattern.match(entry)
        if match:
            if current:
                structured_sessions.append(current)
            current = {
                "session": match.group("session"),
                "date": match.group("date"),
                "topics": match.group("topics"),
                "readings": ""
            }
        else:
            if "readings" in current:
                current["readings"] += " " + entry.strip()
    if current:
        structured_sessions.append(current)

    return structured_sessions

def generate_canvas_schedule_html(sessions):
    rows = ""
    for sess in sessions:
        rows += f"""<tr>
            <td>{sess['date']}</td>
            <td><a href='#'>Class #{sess['session']} - {sess['topics'].split(':')[0]}</a></td>
            <td>{sess['topics']}<br>{sess['readings'].strip()}</td>
        </tr>"""

    html = f"""<style>
      table.schedule {{
        border-collapse: collapse;
        width: 100%;
        font-family: Arial, sans-serif;
      }}
      table.schedule th, table.schedule td {{
        border: 1px solid #cccccc;
        padding: 10px;
        vertical-align: top;
      }}
      table.schedule th {{
        background-color: #003366;
        color: white;
        text-align: left;
      }}
      table.schedule td a {{
        color: #0056b3;
        text-decoration: underline;
      }}
      h2.schedule-title {{
        text-align: center;
        font-family: Arial, sans-serif;
      }}
    </style>

    <h2 class='schedule-title'>Course Schedule & Content</h2>
    <table class='schedule'>
      <thead>
        <tr>
          <th>Date</th>
          <th>Session</th>
          <th>Topics</th>
        </tr>
      </thead>
      <tbody>
        {rows}
      </tbody>
    </table>"""
    return html

def run_agent(input_docx_path, output_html_path):
    sessions = extract_course_schedule_from_docx(input_docx_path)
    if not sessions:
        print("No schedule found in the document.")
        return
    html_content = generate_canvas_schedule_html(sessions)
    with open(output_html_path, "w") as f:
        f.write(html_content)
    print(f"Course schedule exported to: {output_html_path}")

# Example usage:
# run_agent("MACC_542_Syllabus.docx", "canvas_schedule.html")
