import subprocess

import pandas as pd
from fpdf import FPDF, FontFace

COLOURS = {
    "Penicillin": (224, 224, 224),
    "Anti-staphylococcal penicillins": (224, 224, 224),
    "Aminopenicillins": (224, 224, 224),
    "Aminopenicillins with beta-lactamase inhibitors": (204, 229, 255),
    "1st-gen cephalosporin": (169, 169, 169),
    "2nd-gen cephalosporin": (169, 169, 169),
    "3rd-gen cephalosporin": (169, 169, 169),
    "4th-gen cephalosporin": (169, 169, 169),
    "5th-gen cephalosporin": (169, 169, 169),
    "Carbapenems": (255, 235, 204),
    "Monobactams": (255, 255, 153),
    "Quinolones": (224, 255, 224),
    "Aminoglycosides": (255, 204, 153),
    "Macrolides": (153, 255, 255),
    "Lincosamide": (255, 255, 204),
    "Tetracyclines": (255, 204, 204),
    "Glycopeptides": (204, 204, 255),
    "Antimetabolite": (153, 255, 153),
    "Nitroimidazoles": (255, 255, 153),
    "Gram positive cocci": (68, 114, 196),
    "Gram negative bacilli": (192, 0, 0),
    "Gram negative cocci": (144, 86, 145),
    "Anaerobes": (128, 96, 0),
    "Atypicals": (128, 128, 128),
    "Drug Class": (180, 180, 180),
    "Drug Name": (180, 180, 180),
    "Drug Class ": (180, 180, 180),
    "Drug Name ": (180, 180, 180),
}


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
    return df


def map_drugs_bugs(
    drug_df: pd.DataFrame, bug_df: pd.DataFrame
) -> pd.DataFrame:
    """
    Combine drug and bacteria information into a multi-level column DataFrame.

    This function takes two DataFrames: one containing drug information (with columns
    "Drug Class" and "Drug Name") and the other containing bacteria information
    (with columns "Group", "Name", and one column for each drug name). It merges these
    DataFrames to create a combined DataFrame with multi-level columns where the first
    two columns are the drug class and drug name, and subsequent columns represent
    bacteria information for each drug.

    :param drug_df : A DataFrame with columns "Drug Class" and "Drug Name".
    :type drug_df: pd.DataFrame
    :param bug_df: A DataFrame with columns "Group", "Name", and one column
                           for each drug name containing corresponding values.
    :type bug_df: pd.DataFrame

    :return: A combined DataFrame with multi-level columns representing the
                  merged drug and bacteria information.
    rtype: pd.DataFrame
    """
    bug_names = bug_df["Name"].to_list()
    print(bug_names)
    # Set the index for bug_df
    bug_df.set_index(["Group", "Name"], inplace=True)

    # Stack the bug_df to reshape it
    stacked_bug_df = bug_df.stack().reset_index()
    stacked_bug_df.columns = ["Group", "Name", "Drug Name", "Value"]

    # Merge the drug_df in to get drug classes
    # TODO fix the sorting issue introduced by this merge?
    # It does sort the drugs in a logical way though so maybe it's okay.
    combined_df = pd.merge(stacked_bug_df, drug_df, on="Drug Name")

    print(combined_df["Name"].to_list())

    # Pivot the combined_df so that Drug Name becomes the index and maintain order
    combined_df = combined_df.pivot_table(
        index=["Drug Class", "Drug Name"],
        columns=["Group", "Name"],
        values="Value",
        sort=False,  # type: ignore
    )

    return combined_df


def generate_pdf(
    output_filename: str, df: pd.DataFrame, colours: dict[str, tuple]
) -> str:
    """
    Create a PDF from a dataframe using the FPDF2 library.

    :param output_filename: Name of the output PDF file (without extension)
    :type output_filename: str
    :param df: Dataframe to convert to PDF table
    :type df: pd.DataFrame
    :param colours: Dictionary of colours for each drug class
    :type colours: dict
    :return: Path to the PDF output file
    :rtype: str
    """
    pdf_file_path = f"{output_filename}.pdf"
    pdf = FPDF(orientation="landscape")
    pdf.add_page()
    pdf.set_font("Helvetica", size=6)
    pdf.set_draw_color(220, 220, 220)
    pdf.set_line_width(0.1)
    background_colour = (255, 255, 255)

    df = df.map(str)
    # Get list of dataframe column headers
    headers = list(df)
    header_0 = [["Drug Class", "Drug Name"] + [h[0] for h in headers] + ["Drug Name "]]  # type: ignore
    header_1 = [["Drug Class", "Drug Name"] + [h[1] for h in headers] + ["Drug Name "]]  # type: ignore

    # map colours for header subclasses
    for col in headers:
        colours[col[1]] = colours.get(col[0], background_colour)  # type: ignore

    row_index = df.index.tolist()
    classes = [i[0] for i in row_index]
    drugs = [i[1] for i in row_index]
    shown_text = (
        header_0[0] + header_1[0] + classes + drugs + [d + " " for d in drugs]
    )

    # Get list of dataframe rows
    rows = df.values.tolist()
    rows = [
        [index[0], index[1], *row, index[1] + " "]
        for index, row in zip(row_index, rows)
    ]

    # Combine headers and rows in one list
    table_data = header_0 + header_1 + rows

    # get merged column widths and row heights
    col_spans = get_spans(header_0[0])
    row_spans = get_spans(classes)
    row_spans["Drug Class"] = 2
    row_spans["Drug Name"] = 2
    row_spans["Drug Name "] = 2

    heading_style = FontFace(
        color=background_colour, emphasis="BOLD", fill_color=(220, 220, 220)
    )
    blank_style = FontFace(
        color=background_colour, fill_color=background_colour
    )
    table_text = list()
    with pdf.table(
        headings_style=heading_style,
        num_heading_rows=2,  # type: ignore
    ) as table:
        for i, data_row in enumerate(table_data):
            drug_class = data_row[0]  # first column is Drug Class
            row = table.row()
            row_colour = colours.get(drug_class, background_colour)
            pdf.set_fill_color(*row_colour)

            for j, datum in enumerate(data_row):
                col_span = col_spans.get(datum, 1)
                row_span = row_spans.get(datum, 1)
                if i < 2:
                    bacteria_class = datum
                    pdf.set_fill_color(
                        *colours.get(bacteria_class, row_colour)
                    )
                if datum == "nan":
                    style = blank_style
                else:
                    style = None
                if datum not in shown_text:
                    datum = ""
                if datum not in table_text:
                    row.cell(
                        text=datum,
                        style=style,
                        colspan=col_span,
                        rowspan=row_span,
                    )
                    if datum != "":
                        table_text.append(datum)

    pdf.output(pdf_file_path)

    return pdf_file_path


def get_spans(items: list[str]) -> dict[str, int]:
    """
    For a given list of strings, return a dict mapping each unique string to the count of contiguous identical strings. Only correctly works if strings are grouped in blocks of unique items.

    :param items: list of items to process
    :type items: list[str]
    :return: dict mapping strings to counts
    :rtype: dict[str, int]
    """
    counts = dict()
    last_item = ""
    count = 1
    for item in items:
        if item == last_item:
            count += 1
        else:
            count = 1
        counts[item] = count
        last_item = item
    return counts


def generate_image(
    output_filename: str, pdf_file_path: str, resolution: int = 300
) -> str:
    """
    Creates a single page PNG based on the first page of an existing PDF file.
    NB: requires the installation of poppler-utils to work

    :param output_filename: Name of the output PNG file (without extension)
    :type output_filename: str
    :param pdf_file_path: Path to the PDF file being converted to PNG
    :type pdf_file_path: str
    :param resolution: Resolution in DPI for the generated image
    :type resolution: int, optional
    :return: Path to the PNG output file
    :rtype: str
    """
    ext = "png"
    img_filepath = f"{output_filename}.{ext}"
    command = [
        "pdftoppm",
        f"-{ext}",
        "-singlefile",
        "-rx",
        str(resolution),
        "-ry",
        str(resolution),
        pdf_file_path,
        output_filename,
    ]
    try:
        subprocess.run(command, stderr=subprocess.DEVNULL, check=True)
    except subprocess.CalledProcessError as e:
        print(f"\ngenerate_image function error: {e}")
    return img_filepath


def main(file: str) -> None:
    """
    Generate a PDF and JPG version of the data in an Excel worksheet.

    :param file: Name of the input Excel file (without extension)
    :type file: str
    :param sheet_name: Name of the Excel worksheet to read
    :type sheet_name: str
    """
    file_path = f"{file}.xlsx"
    drug_df = read_dataframe(file_path, "Drug Information")
    bug_df = read_dataframe(file_path, "Bacteria Information")
    df = map_drugs_bugs(drug_df, bug_df)
    print(df.head())  # debug

    pdf_file_path = generate_pdf(file, df, COLOURS)
    jpg_file_path = generate_image(file, pdf_file_path)
    print(f"\nGenerated {pdf_file_path}, {jpg_file_path}")


if __name__ == "__main__":
    file = "antibiogram"
    main(file)
