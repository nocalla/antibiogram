import matplotlib.pyplot as plt
import pandas as pd
import pdfkit

# Load the Excel file
file_path = "antibiogram.xlsx"
sheet_name = "Antibiogram"

# Read the named sheet
df = pd.read_excel(file_path, sheet_name=sheet_name)

# Generate PDF
html_content = df.to_html()
pdf_file_path = file_path.replace(".xlsx", ".pdf")
pdfkit.from_string(html_content, pdf_file_path)

# Generate JPG
fig, ax = plt.subplots(figsize=(df.shape[1], df.shape[0] // 2))
ax.axis("tight")
ax.axis("off")
ax.table(
    cellText=df.values, colLabels=df.columns, cellLoc="center", loc="center"
)

jpg_file_path = file_path.replace(".xlsx", ".jpg")
plt.savefig(jpg_file_path, format="jpg")

print(f"Generated {pdf_file_path} and {jpg_file_path}")
