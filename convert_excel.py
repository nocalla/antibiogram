import matplotlib.pyplot as plt
import pandas as pd
from reportlab.lib.pagesizes import letter
from reportlab.pdfgen import canvas

# Load the Excel file
file_path = "antibiogram.xlsx"
sheet_name = "Antibiogram"

# Read the named sheet
df = pd.read_excel(file_path, sheet_name=sheet_name)

# Generate PDF
# Generate PDF using reportlab
pdf_file_path = "converted_output.pdf"
c = canvas.Canvas(pdf_file_path, pagesize=letter)
width, height = letter

# Create a string representation of the dataframe
data_str = df.to_string(index=False)

# Set font and write the data string to the PDF
c.setFont("Helvetica", 10)
c.drawString(30, height - 40, "Excel Sheet Data:")
text = c.beginText(30, height - 60)
text.setFont("Helvetica", 8)
text.textLines(data_str)
c.drawText(text)
c.save()

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
