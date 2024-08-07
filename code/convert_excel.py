import ast
import subprocess

import pandas as pd
from fpdf import FPDF, FontFace


def read_dataframe(
    file_path: str, sheet_name: str, include_col="Include", include_val="Y"
) -> pd.DataFrame:
    """
    Read a named worksheet in an Excel file into a Pandas dataframe,
    excluding certain rows as specified.

    :param file_path: Path to Excel file
    :type file_path: str
    :param sheet_name: Name of worksheet to read
    :type sheet_name: str
    :return: Pandas dataframe of Excel data
    :rtype: pd.DataFrame
    """
    # Read the named sheet with pandas
    df = pd.read_excel(file_path, sheet_name=sheet_name)
    # Only include rows where the include_col value is include_val
    df = df[df[include_col].str.capitalize() == include_val]
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
    # Set the index for bug_df
    bug_df.set_index(["Group", "Name"], inplace=True)

    # Stack the bug_df to reshape it
    stacked_bug_df = bug_df.stack().reset_index()
    stacked_bug_df.columns = ["Group", "Name", "Drug Name", "Value"]

    # Merge the drug_df in to get drug classes
    # TODO fix the sorting issue introduced by this merge?
    # It does sort the drugs in a logical way though so maybe it's okay.
    combined_df = pd.merge(stacked_bug_df, drug_df, on="Drug Name")

    # Pivot the combined_df so that Drug Name becomes the index and maintain order
    combined_df = combined_df.pivot_table(
        index=["Drug Class", "Drug Name"],
        columns=["Group", "Name"],
        values="Value",
        sort=False,  # type: ignore
    )

    return combined_df


class PDF(FPDF):
    def footer(self) -> None:
        # Position cursor at 1.5 cm from bottom:
        self.set_y(-15)
        # Setting font: helvetica italic 8
        self.set_font("helvetica", "I", 6)
        # Printing page number:
        self.cell(
            0,
            10,
            f"Source: https://nocalla.github.io/antibiogram/",
            align="C",
        )


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
    pdf = PDF(orientation="landscape")
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

    # generate column width list
    width_options = [8, 12, 6]
    col_widths = [
        *width_options[0:2],
        *[width_options[2]] * (len(header_0[0]) - 3),
        width_options[1],
    ]

    heading_style = FontFace(
        color=background_colour, emphasis="BOLD", fill_color=(220, 220, 220)
    )
    blank_style = FontFace(
        color=background_colour, fill_color=background_colour
    )
    table_text = list()
    with pdf.table(
        col_widths=col_widths,
        headings_style=heading_style,
        num_heading_rows=2,  # type: ignore
    ) as table:
        for i, data_row in enumerate(table_data):
            drug_class = data_row[0]  # first column is Drug Class
            row = table.row()
            row_colour = colours.get(drug_class, (180, 180, 180))
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
    output_filename: str, pdf_file_path: str, resolution: int = 400
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
    except (subprocess.CalledProcessError, FileNotFoundError) as e:
        print(f"\tgenerate_image function error: {e}")
    return img_filepath


def get_colours(dfs: list[pd.DataFrame]) -> dict[str, tuple[int]]:
    """
    Create a dictionary mapping the values in the first column of each
    provided dataframe to the value in the "Colour" column.

    :param dfs: List of Pandas dataframes to act on.
    :type dfs: list[pd.DataFrame]
    :return: Dictionary of strings mapped to RGB tuples
    :rtype: dict[str, tuple]
    """
    colour_headers = ["Label", "Colour"]
    colour_dfs = list()

    for df in dfs:
        colour_index = list(df.columns).index("Colour")
        df = df.iloc[:, [0, colour_index]]
        df.columns = colour_headers
        colour_dfs.append(df)

    colour_df = pd.concat(colour_dfs).drop_duplicates(subset=colour_headers[0])
    colour_df.dropna(inplace=True)
    colour_df.set_index(keys=colour_headers[0], inplace=True)
    colour_df[colour_headers[1]] = colour_df[colour_headers[1]].apply(
        lambda x: tuple(map(int, ast.literal_eval(x)))
    )
    colour_dict = colour_df.to_dict(orient="dict")[colour_headers[1]]

    return colour_dict


def main(file: str) -> None:
    """
    Generate a PDF and JPG version of the data in an Excel worksheet.

    :param file: Name of the input Excel file (without extension)
    :type file: str
    :param sheet_name: Name of the Excel worksheet to read
    :type sheet_name: str
    """
    # TODO - use a progress bar instead of successive print messages
    file_path = f"{file}.xlsx"
    print(f"Reading drug information from {file_path}")
    drug_df = read_dataframe(file_path, "Drug Information")
    print(f"Reading bacteria information from {file_path}")
    bug_df = read_dataframe(file_path, "Bacteria Information")
    print("Mapping drugs to bacterial groups")
    df = map_drugs_bugs(
        drug_df.drop("Colour", axis="columns"),
        bug_df.drop("Colour", axis="columns"),
    )
    print(f"Getting colours from {file_path}")
    colours = get_colours([drug_df, bug_df])
    print("Generating PDF file")
    pdf_file_path = generate_pdf(file, df, colours)
    print("Generating image file")
    img_file_path = generate_image(file, pdf_file_path)
    print(f"Done!\n\tGenerated {pdf_file_path}, {img_file_path}")


if __name__ == "__main__":
    file = "docs/assets/antibiogram"
    main(file)
