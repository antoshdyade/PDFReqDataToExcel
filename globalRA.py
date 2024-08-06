"""
Code By : Antosh Madappa Dyade, (antosh@dyade.in) under CC BY-SA 4.0 License.
"""

import pdfplumber
import pandas as pd

def extract_data_from_pdf(pdf_path):
  """
    Extracts PRN, Name, ISE(Obt), ICA(Obt), POE(Obt), and Total(Obt) from a PDF file.

    Args:
    pdf_path (str): Path to the input PDF file.

    Returns:
    tuple: Lists containing PRN, Name, ISE(Obt), ICA(Obt), POE(Obt), and Total(Obt).
    """
    
    prn_list = []
    name_list = []
    ise_obt_list = []
    ica_obt_list = []
    poe_obt_list = []
    total_obt_list = []

    with pdfplumber.open(pdf_path) as pdf:
        prn = None
        name = None

        for page_num, page in enumerate(pdf.pages):
            text = page.extract_text().split('\n')
            for line in text:
                if 'PRN:' in line:
                    prn = line.split('PRN:')[1].split()[0].strip()
                    print(f"Found PRN: {prn} on page {page_num + 1}")
                if 'Name:' in line:
                    name = line.split('Name:')[1].strip()
                    print(f"Found Name: {name} on page {page_num + 1}")
                if prn and name:
                    prn_list.append(prn)
                    name_list.append(name)
                    ise_obt = None
                    ica_obt = None
                    poe_obt = None
                    total_obt = None

                    btn03405_count = 0
                    for subsequent_line in text[text.index(line)+1:]:
                        if 'BTN03405' in subsequent_line:
                            btn03405_count += 1
                            subject_data = subsequent_line.split()
                            print(f"Found BTN03405 line: {subsequent_line} on page {page_num + 1}")
                            print(f"Subject Data: {subject_data}")

                            if btn03405_count == 1:
                                try:
                                    ise_obt = subject_data[3].strip()  # ISE(Obt)
                                except IndexError:
                                    ise_obt = None
                            elif btn03405_count == 2:
                                try:
                                    ica_obt = subject_data[3].strip()  # ICA(Obt)
                                    poe_obt = subject_data[5].strip()  # POE(Obt)
                                except IndexError:
                                    ica_obt = None
                                    poe_obt = None
                            elif btn03405_count == 3:
                                try:
                                    total_obt = subject_data[3].strip()  # Total(Obt)
                                except IndexError:
                                    total_obt = None
                                break
                    
                    ise_obt_list.append(ise_obt)
                    ica_obt_list.append(ica_obt)
                    poe_obt_list.append(poe_obt)
                    total_obt_list.append(total_obt)
                    
                    prn = None
                    name = None

    return prn_list, name_list, ise_obt_list, ica_obt_list, poe_obt_list, total_obt_list

def save_to_excel(data, excel_path):
    df = pd.DataFrame({
        'PRN': data[0],
        'Name': data[1],
        'ISE(Obt)': data[2],
        'ICA(Obt)': data[3],
        'POE(Obt)': data[4],
        'Total(Obt)': data[5]
    })
    df.to_excel(excel_path, index=False)

pdf_path = 'C:/Users/Antosh/Downloads/r7/result.pdf'  # Replace with your PDF file path
excel_path = 'C:/Users/Antosh/Downloads/r7/extracted_data.xlsx'  # Replace with your desired output Excel file path

data = extract_data_from_pdf(pdf_path)
save_to_excel(data, excel_path)

print(f"Data has been successfully extracted from {pdf_path} and written to {excel_path}")
