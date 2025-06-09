"""
:author: Valentin Stoyanov <valentin@suricactus.ltd>
:license: MIT
:date: 2025-05-27
"""

from openpyxl import load_workbook
import xml.etree.ElementTree as ET
import argparse

# TODO: Rework this function to generate the same data structure as in `convert_xls.py`.
def excel_to_flat_xml(excel_path, xml_output_path, sheet_name=None):
    # Load the workbook and select the sheet
    wb = load_workbook(excel_path, data_only=True)
    ws = wb[sheet_name] if sheet_name else wb.active

    if not ws:
        raise Exception(f"Failed to open worksheet {sheet_name}")

    # Read header row
    headers = [cell.value for cell in ws[1]]

    # Create XML root element
    root = ET.Element("TransferData")
    root.set("xmlns", "urn:Transfer")

    accountings_element = ET.SubElement(root, "Accountings")

    # Iterate over rows, skipping header
    for row in ws.iter_rows(min_row=2, values_only=True):
        row_elem = ET.SubElement(accountings_element, "Accounting")
        for header, cell_value in zip(headers, row):
            cell_elem = ET.SubElement(row_elem, str(header))
            cell_elem.text = "" if cell_value is None else str(cell_value)

    # Write the XML to file
    tree = ET.ElementTree(root)
    tree.write(xml_output_path, encoding="utf-8", xml_declaration=True)
    print(f"XML saved to {xml_output_path}")

# Example usage
if __name__ == "__main__":
    parser = argparse.ArgumentParser(
        description="Converts a new-format excel chron report to an XML import file format."
    )
    parser.add_argument("-f", "--filepath", help="Path to the .xls input file")
    parser.add_argument("-o", "--output", help="Path to the output .xml file")
    parser.add_argument("-s", "--sheet", type=int, default=0, help="Sheet index (default: 0)")
    parser.add_argument("-r", "--start-row", type=int, default=1, help="Row index to start reading from, including headers (default: 1)")
    args = parser.parse_args()

    excel_to_flat_xml(args.filename, args.output)
