import subprocess

import pandas as pd
from fpdf import FPDF


def read_dataframe(file_path: str, sheet_name: str) -> pd.DataFrame:
    """
    Read a named worksheet in an Excel file into a Pandas dataframe.

    :param file_path: Path to Excel file
    :type file_path: str
    :param sheet_name: Name of worksheet to read
    :type sheet_name: str
    :return: Pandas dataframe of Excel data
    :rtype: pd.DataFrame
    """
    # Read the named sheet with pandas
    df = pd.read_excel(file_path, sheet_name=sheet_name)
    df.fillna(0, inplace=True)
    df["Drug Name "] = df["Drug Name"]
    cols = ["Drug Class", "Drug Name"]
    df = df[cols + [c for c in df.columns if c not in cols]]
    print(df.head())  # debug
    return df


def generate_pdf(output_filename: str, df: pd.DataFrame) -> str:
    """
    Create a PDF from a dataframe using the FPDF2 library.

    :param output_filename: Name of the output PDF file (without extension)
    :type output_filename: str
    :param df: Dataframe to convert to PDF table
    :type df: pd.DataFrame
    :return: Path to the PDF output file
    :rtype: str
    """
    # Create a PDF file using fpdf2
    pdf_file_path = f"{output_filename}.pdf"
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


def generate_jpg(output_filename: str, pdf_file_path: str) -> str:
    """
    Creates a single page JPG based on the first page of an existing PDF file.
    NB: requires the installation of poppler-utils to work

    :param output_filename: Name of the output JPG file (without extension)
    :type output_filename: str
    :param pdf_file_path: Path to the PDF file being converted to JPG
    :type pdf_file_path: str
    :return: Path to the JPG output file
    :rtype: str
    """
    jpg_file_path = f"{output_filename}.jpg"
    command = [
        "pdftoppm",
        "-jpeg",
        "-singlefile",
        pdf_file_path,
        output_filename,
    ]
    subprocess.run(command, check=True)
    return jpg_file_path


def main(file: str, sheet_name: str) -> None:
    """
    Generate a PDF and JPG version of the data in an Excel worksheet.

    :param file: Name of the input Excel file (without extension)
    :type file: str
    :param sheet_name: Name of the Excel worksheet to read
    :type sheet_name: str
    """
    file_path = f"{file}.xlsx"

    df = read_dataframe(file_path, sheet_name)
    pdf_file_path = generate_pdf(file, df)
    jpg_file_path = generate_jpg(file, pdf_file_path)

    print(f"Generated {pdf_file_path}, {jpg_file_path}")


if __name__ == "__main__":
    file = "antibiogram"
    sheet_name = "Drug Information"
    main(file, sheet_name)
