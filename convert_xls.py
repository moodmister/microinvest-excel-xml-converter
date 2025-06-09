"""
:author: Valentin Stoyanov <valentin@suricactus.ltd>
:license: MIT
:date: 2025-05-27
"""

import argparse
from collections import defaultdict
from datetime import datetime
import xlrd
import xml.etree.ElementTree as ET

EXCEL_HEADERS_TO_XML_ATTRIBUTES = {
    "Контиране": "Number",
    "Дата": "AccountingDate",
    "Дебит сметка": "AccountNumber",
    "Кредит сметка": "AccountNumber",
    "Сума": "Amount",
    "Док. вид": "DocumentType",
    "Док. дата": "Date",
    "Документ №": "Number",
    "Партньор": "Name",
    "ЕИК/ДДС Номер": "Bulstat",
    # "Държава": "", ?
    "Основание": "Term",
    "Забележка": "Reference",
    "Втора забележка": "OptionalReference",
    "Сделка по ЗДДС": "VatTerm",
    "Параграф": "",
    "Месец за експорт": "ViesMonth",
}

DOCUMENT_TYPES_TO_CODES = {
    "Ф-ра": "1",
    "ДИ": "2",
    "КИ": "3",
    "ОП": "6",
    "ПКО": "8",
    "РКО": "9",
    "МО": "10",
    "БИ": "11",
}

def float_format(amount):
    return "{:.6f}".format(round(float(amount), 6))

def excel_date_to_str(excel_date):
    if not excel_date:
        return ""

    return datetime.fromordinal(
        datetime(1900, 1, 1).toordinal() + int(excel_date) - 2
    ).strftime("%Y-%m-%d")

def populate_attributes(xml_element, source_dict):
    for key, value in source_dict.items():
            if type(value) not in (str, float, int):
                continue
            xml_element.set(key, value)

def xls_to_flat_xml(xls_path, xml_output_path, sheet_index=0, start_row=1):
    workbook = xlrd.open_workbook(xls_path)
    sheet = workbook.sheet_by_index(sheet_index)

    headers = sheet.row_values(start_row)

    accountings_by_number = defaultdict(list)
    for row_idx in range(start_row + 1, sheet.nrows):
        row_values = sheet.row_values(row_idx)
        data_structure = dict(zip(headers, row_values))
        accountings_by_number[data_structure["Контиране"]].append({
            "Accounting": {
                "AccountingDate": excel_date_to_str(data_structure["Дата"]),
                "Number": data_structure["Контиране"],
                "Reference": data_structure["Забележка"],
                "OptionalReference": data_structure["Втора забележка"],
                "Term": data_structure["Основание"],
                "Vies": "1",
                "ViesMonth": excel_date_to_str(data_structure["Месец за експорт"]),
            },
            "Document": {
                "Date": excel_date_to_str(data_structure["Док. дата"]),
                "Number": data_structure["Документ №"],
                "DocumentType": DOCUMENT_TYPES_TO_CODES[data_structure["Док. вид"]],
            },
            "Company": {
                "Name": data_structure["Партньор"],
                "Bulstat": data_structure["ЕИК/ДДС Номер"].replace("BG", ""),
                "VatNumber": data_structure["ЕИК/ДДС Номер"] if "BG" in data_structure["ЕИК/ДДС Номер"] else "",
                "BankAccounts": {
                }
            },
            "AccountingDetails": [
                {
                    "Direction": "Debit",
                    "AccountNumber": data_structure["Дебит сметка"],
                    "Amount": float_format(data_structure["Сума"]),
                    "VatTerm": data_structure["Сделка по ЗДДС"][0],
                },
                {
                    "Direction": "Credit",
                    "AccountNumber": data_structure["Кредит сметка"],
                    "Amount": float_format(data_structure["Сума"]),
                    "VatTerm": data_structure["Сделка по ЗДДС"][0],
                },
            ],
        })

    root = ET.Element("TransferData")
    root.set("xmlns", "urn:Transfer")
    accountings_element = ET.SubElement(root, "Accountings")
    for number, accountings in accountings_by_number.items():
        accounting_element = ET.SubElement(accountings_element, "Accounting")
        populate_attributes(accounting_element, accountings[0]["Accounting"])

        document_element = ET.SubElement(accounting_element, "Document")
        populate_attributes(document_element, accountings[0]["Document"])

        company_element = ET.SubElement(accounting_element, "Company")
        populate_attributes(company_element, accountings[0]["Company"])
        bank_account_element = ET.SubElement(company_element, "BankAccounts")

        accounting_details_element = ET.SubElement(accounting_element, "AccountingDetails")
        for accounting in accountings:

            for accounting_detail in accounting["AccountingDetails"]:
                accounting_detail_element = ET.SubElement(accounting_details_element, "AccountingDetail")
                populate_attributes(accounting_detail_element, accounting_detail)

    tree = ET.ElementTree(root)
    ET.indent(tree)
    tree.write(xml_output_path, encoding="utf-8", xml_declaration=True)
    print(f"XML saved to {xml_output_path}")

def main():
    parser = argparse.ArgumentParser(
        description="Converts an old-format excel chron report to an XML import file format."
    )
    parser.add_argument("-f", "--filepath", help="Path to the .xls input file")
    parser.add_argument("-o", "--output", help="Path to the output .xml file")
    parser.add_argument("-s", "--sheet", type=int, default=0, help="Sheet index (default: 0)")
    parser.add_argument("-r", "--start-row", type=int, default=1, help="Row index to start reading from, including headers (default: 1)")
    
    args = parser.parse_args()

    xls_to_flat_xml(args.filepath, args.output, sheet_index=args.sheet, start_row=args.start_row)

if __name__ == "__main__":
    main()
