import subprocess

import pandas as pd
from fpdf import FPDF


def read_dataframe(file_path: str, sheet_name: str) -> pd.DataFrame:
    # Read the named sheet with pandas
    df = pd.read_excel(file_path, sheet_name=sheet_name)
    df.fillna(0, inplace=True)
    df["Drug Name "] = df["Drug Name"]
    cols = ["Drug Class", "Drug Name"]
    df = df[cols + [c for c in df.columns if c not in cols]]
    print(df.head())  # debug
    return df


def generate_pdf(file: str, df: pd.DataFrame) -> str:
    # Create a PDF file using fpdf2
    pdf_file_path = f"{file}.pdf"
    df = df.map(str)
    COLUMNS = [list(df)]  # Get list of dataframe columns
    ROWS = df.values.tolist()  # Get list of dataframe rows
    DATA = COLUMNS + ROWS  # Combine columns and rows in one list

    pdf = FPDF(orientation="landscape")
    pdf.add_page()
    pdf.set_font("Helvetica", size=8)
    with pdf.table(
        borders_layout="MINIMAL",
    ) as table:
        for data_row in DATA:
            table.row(data_row)
    pdf.output(pdf_file_path)

    return pdf_file_path


def generate_jpg(file: str, pdf_file_path: str) -> str:
    jpg_file_path = f"{file}.jpg"
    command = ["pdftoppm", "-jpeg", "-singlefile", pdf_file_path, file]
    subprocess.run(command, check=True)
    return jpg_file_path


def main(file: str, sheet_name: str) -> None:
    file_path = f"{file}.xlsx"

    df = read_dataframe(file_path, sheet_name)
    pdf_file_path = generate_pdf(file, df)
    jpg_file_path = generate_jpg(file, pdf_file_path)

    print(f"Generated {pdf_file_path}, {jpg_file_path}")


if __name__ == "__main__":
    # Load the Excel file
    file = "antibiogram"
    sheet_name = "Drug Information"
    main(file, sheet_name)
