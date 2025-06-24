import os
import re
import fitz  # PyMuPDF
from reportlab.lib.pagesizes import A4
from reportlab.platypus import SimpleDocTemplate, Paragraph, Table, TableStyle
from reportlab.lib.styles import getSampleStyleSheet
from reportlab.lib import colors

def extract_text_from_pdf(pdf_path):
    doc = fitz.open(pdf_path)
    text = ""
    for page in doc:
        text += page.get_text()
    return text

def extract_date(text):
    match = re.search(r'\b(\d{1,2}[/-]\d{1,2}[/-]\d{4})\b', text)
    return match.group(1) if match else "Unknown Date"

def extract_summary(text):
    start = text.lower().find("summary of analysis")
    if start != -1:
        section = text[start:start + 700]
        lines = section.splitlines()
        # Remove Cu/Fe/C lines
        filtered = [line for line in lines if not re.match(r'^(Cu|Fe|C)\s+\d+', line.strip())]
        return "\n".join(filtered).strip()
    return "Summary not found."

def extract_metallurgical_table(text):
    # Look for Cu/Fe/C values anywhere in the text
    matches = re.findall(r'(Cu|Fe|C)\s+([\d.]+)', text)
    table = [["Observation", "Details"]]
    for metal, value in matches:
        table.append([metal, value])
    if len(table) == 1:
        table.append(["No metallurgical observations found.", ""])
    return table

def categorize_report(text):
    text = text.lower()
    if "material analysis" in text:
        return "Material Analysis"
    elif "failure analysis" in text:
        return "Failure Analysis"
    return "Unknown"

def create_pdf(component, date, summary, table_data, output_filename):
    doc = SimpleDocTemplate(output_filename, pagesize=A4)
    styles = getSampleStyleSheet()
    elements = []

    elements.append(Paragraph(f"<b>Material Analysis of {component}</b>", styles["Title"]))
    elements.append(Paragraph(f"<b>Date:</b> {date}", styles["Normal"]))
    elements.append(Paragraph(f"<b>Summary of Analysis:</b><br/><br/>{summary}", styles["BodyText"]))

    table = Table(table_data, colWidths=[270, 270])
    table.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (-1, 0), colors.grey),
        ("TEXTCOLOR", (0, 0), (-1, 0), colors.whitesmoke),
        ("ALIGN", (0, 0), (-1, -1), "LEFT"),
        ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
        ("BOTTOMPADDING", (0, 0), (-1, 0), 12),
        ("GRID", (0, 0), (-1, -1), 1, colors.black),
    ]))
    elements.append(table)

    doc.build(elements)
    print(f"✅ PDF created: {output_filename}")

def main():
    folder = input("Enter the folder path containing PDFs: ").strip()
    component = input("Enter component name: ").strip().lower()

    for file in os.listdir(folder):
        if file.lower().endswith(".pdf"):
            path = os.path.join(folder, file)
            text = extract_text_from_pdf(path)

            if component in text.lower():
                category = categorize_report(text)
                if category != "Material Analysis":
                    print(f"⚠️ File '{file}' is {category}, not Material Analysis. Skipping.")
                    continue

                date = extract_date(text)
                summary = extract_summary(text)
                table = extract_metallurgical_table(text)

                outname = f"{component.replace(' ', '_')}_material_analysis_summary.pdf"
                create_pdf(component.title(), date, summary, table, outname)
                return

    print("❌ Component not found in any PDF.")

if __name__ == "__main__":
    main()
