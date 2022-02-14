from pathlib import Path
import re
from docx import Document
import pandas as pd

# Between regex and paths, pylint sees too many issues with this script
# pylint: disable=anomalous-backslash-in-string

# SOP specific parameters
# - Match 4 digits at start -> 0 or 1 characters -> 0 or 1 spaces
# --> a hyphen
# --> any character except \n
SOP_REGEX = "^(\d{4}\w?\s?)[-](.)+"
SOP_PATH = "P:\SOPs\ALL Files"


def sop_search(search_regex, search_path):
    """Search and parses SOP folder"""
    sops = []

    for path in Path(search_path).rglob("*"):
        if (
            re.search(search_regex, path.name)
            and "Archive" not in path.parts
            and Path.is_file(path)
        ):
            department = path.parts[3].split("-")[1].strip()
            number = re.findall("^\d{4}\w?", path.name)[0]
            title = path.stem.split("-")[1].strip()
            hyperlink = '=HYPERLINK("%s", "%s")' % (path, path)

            # Last Revision Date
            date = None
            if path.suffix == ".docx":
                try:
                    document = Document(path)
                    # I know this nesting looks stupid, but it's what the documentation suggests 😳
                    for table in document.tables:
                        for row in table.rows:
                            for cell in row.cells:
                                for para in cell.paragraphs:
                                    # Matches dates in formats like M/D/YY, MM/DD/YYYY, and some permutations
                                    if re.search(
                                        "^(0?[1-9]|1[0-2])[\/](0?[1-9]|[12]\d|3[01])[\/]((19|20)\d{2}|\d{2})$",
                                        para.text.strip(),
                                    ):
                                        # The last date in the last table *appears* to be the most reliable way of getting last revision
                                        # This is the sketchiest part of the whole thing 😬
                                        date = para.text
                except IOError:
                    print("Could not open file")

            row = {
                "Department": department,
                "Number": number,
                "File Name/Title": title,
                "Link to documents": hyperlink,
                "Last Revision Date": date,
            }

            sops.append(row)
    return sops


def export_to_excel(files_dict, output_filename, sheet_name):
    """Exports a list of dictionaries to excel"""
    df = pd.DataFrame(files_dict)
    df.to_excel(output_filename, sheet_name=sheet_name, index=False)


export_to_excel(sop_search(SOP_REGEX, SOP_PATH), "SOPS.xlsx", "SOPS")
