import matplotlib.pyplot as plt
import pandas as pd
from openpyxl import load_workbook
from reportlab.lib.pagesizes import letter
from reportlab.pdfgen import canvas

# Load the Excel file
file_path = "antibiogram.xlsx"
sheet_name = "Antibiogram"

# Read the named sheet with pandas
df = pd.read_excel(file_path, sheet_name=sheet_name)

# Load the workbook and sheet with openpyxl for formatting
wb = load_workbook(file_path, data_only=True)
ws = wb[sheet_name]

# Extract cell styles
styles = {}
for row in ws.iter_rows():
    for cell in row:
        if cell.value is not None:
            styles[(cell.row, cell.column)] = cell

# Create a PDF file using reportlab
pdf_file_path = "antibiogram.pdf"
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

# Generate JPG using matplotlib
fig, ax = plt.subplots(figsize=(df.shape[1] * 1.5, len(df) // 2))
ax.axis("tight")
ax.axis("off")

table = ax.table(
    cellText=df.values, colLabels=df.columns, cellLoc="center", loc="center"
)

# Apply formatting
for (i, j), cell in table.get_celld().items():
    if (i, j) in styles:
        xl_cell = styles[(i, j)]
        # Set background color if not default
        if xl_cell.fill.start_color.index != '00000000' and xl_cell.fill.start_color.rgb is not None:
            color = xl_cell.fill.start_color.rgb[2:]  # Remove the 'FF' prefix
            cell.set_facecolor('#' + color)
        # Set font weight and style
        if xl_cell.font.bold:
            cell.set_text_props(weight='bold')
        if xl_cell.font.italic:
            cell.set_text_props(style='italic')
        # Set font color if specified
        if xl_cell.font.color is not None and xl_cell.font.color.rgb is not None:
            color = xl_cell.font.color.rgb[2:]  # Remove the 'FF' prefix
            cell.set_text_props(color='#' + color)

# Save the JPG file with the same base name as the PDF file
jpg_file_path = "antibiogram.jpg"
plt.savefig(jpg_file_path, format="jpg", bbox_inches='tight')

print(f"Generated {pdf_file_path} and {jpg_file_path}")
