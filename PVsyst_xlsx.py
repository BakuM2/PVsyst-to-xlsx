import os
import re
import pdfplumber
import pandas as pd

cwd = os.getcwd()

# Initialize a dictionary to store extracted data
data = {}

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
            value = re.findall(r"\d+\.?\d*", value)[0]  # Extract only numeric value
            value = float(value)
        return value
    return None


def pdf_to_df(pdf_path):
    with pdfplumber.open(pdf_path) as pdf:
        text = "\n".join([page.extract_text() for page in pdf.pages if page.extract_text()])

        for key, label in keys_to_extract.items():
            numeric_only = key in ["System power:", "Tilt", "Nb. of modules", "Pnom total", "Nb. of units",
                                   "Produced Energy", "Specific production",
                                   "Perf. Ratio PR"]
            extracted_value = extract_value(text, key, numeric_only)
            if extracted_value:
                data[label] = extracted_value
        return data


def df_to_excel(df, excel_path):
    df = pd.DataFrame([data])
    df.to_excel(excel_path, index=False)

    print(f"Excel file saved: {excel_path}")


if __name__ == "__main__":
    filename="Project.pdf"

    pdf_path = cwd + "/" + filename
    excel_path = cwd + "/" + filename[:-3] + "xlsx"
    df = pdf_to_df(pdf_path)
    df_to_excel(df, excel_path)

