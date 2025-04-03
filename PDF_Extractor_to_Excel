import os
from openpyxl import Workbook
import pdfplumber
import re
from datetime import datetime
import mysql.connector

patterns = {
    "Número da Fatura": r'INVOICE\s*#\s*(\d+)',
    "Data da Fatura": r'DATE[:\s]*(\d{2}/\d{2}/\d{4})'
}

def execute_insert(cursor, data):
    columns = ', '.join(data[0].keys())
    placeholders = ', '.join(['%s'] * len(data[0]))
    sql = f"INSERT INTO invoice_records ({columns}) VALUES ({placeholders})"

    values = [tuple(d.values()) for d in data]
    cursor.executemany(sql, values)

def extract_data_from_text(pdf_text):
    pdf_text = re.sub(r"\s+", " ", pdf_text) 
    extracted_data = {}

    for label, pattern in patterns.items():
        match = re.search(pattern, pdf_text, re.IGNORECASE)
        extracted_data[label] = match.group(1) if match else "N/A"

    return extracted_data

def process_pdfs(directory):
    if not os.path.exists(directory):
        raise FileNotFoundError(f"The directory '{directory}' was not found.")

    files = [f for f in os.listdir(directory) if f.endswith(".pdf")]
    if not files:
        raise Exception("No PDF files found in the directory.")

    results = []

    for file in files:
        file_path = os.path.join(directory, file)
        try:
            pdf_text = ""
            with pdfplumber.open(file_path) as pdf:
                for page in pdf.pages:
                    text = page.extract_text()
                    if text:
                        pdf_text += text + " "

            extracted_data = extract_data_from_text(pdf_text)
            extracted_data["File Name"] = file
            extracted_data["Status"] = "Completed"

            if all(value == "N/A" for key, value in extracted_data.items() if key != "File Name"):
                raise ValueError(f"No information extracted from the file {file}")

            results.append(extracted_data)

        except Exception as e:
            print(f"Erro ao processar {file}: {e}")
            error_data = {key: "N/A" for key in patterns.keys()}
            error_data["File Name"] = file
            error_data["Status"] = f"Erro: {e}"
            results.append(error_data)

    return results

def save_to_excel(data):
    wb = Workbook()
    ws = wb.active
    ws.title = "Extracted Data"

    headers = list(data[0].keys())
    ws.append(headers)

    for row in data:
        ws.append(list(row.values()))

    ws.auto_filter.ref = ws.dimensions

    timestamp = datetime.now().strftime("%Y-%m-%d_%H-%M-%S")
    filename = f"Extracted_Data_{timestamp}.xlsx"

    wb.save(filename)
    print(f"Excel file saved as {filename}")

def save_to_database(data):
    try:
        with mysql.connector.connect(
            host="localhost",
            user="root",
            password="",
            database="process_invoices"
        ) as db:
            with db.cursor() as cursor:
                execute_insert(cursor, data)
            db.commit()
            print("Data successfully inserted into the database.")
    except mysql.connector.Error as e:
        print(f"Error connecting to the database: {e}")

def main():
    directory = "pdf_directory"

    print("--- Starting processing ---")

    try:
        extracted_data = process_pdfs(directory)
        save_to_excel(extracted_data)
        save_to_database(extracted_data)

        print("--- Processing completed successfully! ---")

    except Exception as e:
        print(f"Erro crítico: {e}")

if __name__ == "__main__":
    main()
