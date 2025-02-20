import os
import re
import pdfplumber
import pandas as pd

cwd = os.getcwd()

# Define keys to extract
keys_to_extract = {
    "Project:": "Project",
    "Variant:": "Variant",
    "System power:": "System Power (kWp)",
    "Tilt": "Tilt (°)",
    "Nb. of modules": "Number of Modules",
    "Pnom total": "Pnom Total (kWp)",
    "Nb. of units": "Number of Inverters",
    "Produced Energy": "Produced Energy (kWh/year)",
    "Specific production": "Specific Production (kWh/kWp/year)",
    "Perf. Ratio PR": "Performance Ratio (PR %)",
}


def extract_value(text, key, numeric_only=False):
    if key in text:
        value = text.split(key)[1].split("\n")[0].strip()
        if numeric_only:
            numbers = re.findall(r"\d+\.?\d*", value)
            value = float(numbers[0]) if numbers else None
        return value
    return None


def pdf_to_df(pdf_path):
    with pdfplumber.open(pdf_path) as pdf:
        text = "\n".join([page.extract_text() for page in pdf.pages if page.extract_text()])

        # Check if it's a PVsyst report
        if "PVsyst - Simulation report" not in text:
            return None

        data = {"Filename": os.path.basename(pdf_path)}
        for key, label in keys_to_extract.items():
            numeric_only = key in ["System power:", "Tilt", "Nb. of modules", "Pnom total", "Nb. of units",
                                   "Produced Energy", "Specific production", "Perf. Ratio PR"]
            extracted_value = extract_value(text, key, numeric_only)
            if extracted_value is not None:
                data[label] = extracted_value
        return data


def parse_pdfs_in_folder(folder_path, excel_path):
    all_data = []
    for filename in os.listdir(folder_path):
        if filename.endswith(".pdf"):
            pdf_path = os.path.join(folder_path, filename)
            data = pdf_to_df(pdf_path)
            if data:
                all_data.append(data)

    if all_data:
        df = pd.DataFrame(all_data)
        df.to_excel(excel_path, index=False)
        print(f"Excel file saved: {excel_path}")
    else:
        print("No valid PVsyst reports found.")


if __name__ == "__main__":
    folder_path = cwd  # Change this if needed
    excel_path = os.path.join(cwd, "PVsyst_Reports.xlsx")
    parse_pdfs_in_folder(folder_path, excel_path)
