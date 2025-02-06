import os
import pdfplumber
import pandas as pd
cwd = os.getcwd()


def pdf_to_excel(pdf_path):
# Initialize a dictionary to store extracted data
    data = {}
    with pdfplumber.open(pdf_path) as pdf:
        text = "\n".join([page.extract_text() for page in pdf.pages if page.extract_text()])
        # print(text)

        # Extract Page 1 Information
        if "Project:" in text:
            project = text.split("Project:")[1].split("\n")[0].strip()
            data["Project"] = project

        if "Variant:" in text:
            variant = text.split("Variant:")[1].split("\n")[0].strip()
            data["Variant"] = variant

        if "System power:" in text:
            system_power = text.split("System power:")[1].split("\n")[0].strip()
            system_power = system_power.split(" ")[0]  # Remove everything after the first space
            data["System Power (kWp)"] = system_power


        if "Tilt" in text:
            tilt = text.split("Tilt")[1].split("\n")[0].strip()

            data["Tilt"] = tilt

        if "Nb. of modules" in text:
            nb_modules = text.split("Nb. of modules")[1].split("\n")[0].strip()
            nb_modules= nb_modules.split(" ")[0]
            data["Number of Modules"] = nb_modules

        if "Pnom total" in text:
            pnom_total = text.split("Pnom total")[1].split("\n")[0].strip()
            data["Pnom Total"] = pnom_total

        if "Nb. of inverters" in text:
            nb_inverters = text.split("Nb. of inverters\n")[1].split("\n")[0].strip()
            data["Number of Inverters"] = nb_inverters

        if "Produced Energy" in text:
            produced_energy = text.split("Produced Energy")[1].split("\n")[0].strip()
            produced_energy = produced_energy.split(" ")[0]
            data["Produced Energy"] = produced_energy

        if "Specific production" in text:
            specific_production = text.split("Specific production")[1].split("\n")[0].strip()
            specific_production = specific_production.split(" ")[0]
            data["Specific Production"] = specific_production

        if "Perf. Ratio PR" in text:
            pr = text.split("Perf. Ratio PR")[1].split("\n")[0].strip()
            pr = pr.split(" ")[0]
            data["Performance Ratio"] = pr


    df = pd.DataFrame([data])
    df.to_excel(excel_path, index=False)

    print(f"Excel file saved: {excel_path}")

if __name__ == "__main__":
    filename="Project.pdf"

    pdf_path = cwd+"/"+filename
    excel_path = cwd+"/"+filename[:-3]+"xlsx"
    pdf_to_excel(pdf_path)

