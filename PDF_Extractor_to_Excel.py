# Let's import all the necessary libraries
import os
from openpyxl import Workbook
import pdfplumber
import re
from datetime import datetime
import mysql.connector

# 🔍 Pattern dictionary: 
# 🔄 Modify here to extract specific information from the PDF
patterns = {
    # We have here just two examples.
    "Número da Fatura": r'INVOICE\s*#\s*(\d+)', # Captures the invoice number after ”INVOICE #”
    "Data da Fatura": r'DATE[:\s]*(\d{2}/\d{2}/\d{4})' # Captures the date in dd/mm/yyyy format after ”DATE:”
}
'''📌 Explanation
The patterns dictionary contains the fields we want to extract from the PDFs.
- Each key represents the name of the field.
- Each value contains a regex (regular expression) to find that field in the text.
- To add a new field, simply add a new line to the dictionary.'''

def execute_insert(cursor, data):
    # Performs batch inserts into the database.
    columns = ', '.join(data[0].keys())
    placeholders = ', '.join(['%s'] * len(data[0]))
    sql = f"INSERT INTO invoice_records ({columns}) VALUES ({placeholders})"

    values = [tuple(d.values()) for d in data]
    cursor.executemany(sql, values)
'''📌 Explanation
- Dynamically creates the SQL query based on the keys in the data dictionary.
- Uses executemany to insert several lines at once.
- placeholders define the %s dynamically, avoiding SQL Injection.'''

def extract_data_from_text(pdf_text):
    # Extracts data from PDF text based on defined standards.
    pdf_text = re.sub(r"\s+", " ", pdf_text) 
    extracted_data = {}

    for label, pattern in patterns.items():
        match = re.search(pattern, pdf_text, re.IGNORECASE)
        extracted_data[label] = match.group(1) if match else "N/A"

    return extracted_data
'''📌 Explicação
- Limpa o texto do PDF para evitar problemas com espaços e quebras de linha.
- Percorre cada padrão regex definido no dicionário patterns e tenta encontrar os valores no texto.
- Se o campo não for encontrado, retorna "N/A".'''

def process_pdfs(directory):
    # Processes PDF files and returns the extracted data.
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
                        pdf_text += text + " " # Add the text from all the pages

            extracted_data = extract_data_from_text(pdf_text) # Extract the data from the text
            extracted_data["File Name"] = file
            extracted_data["Status"] = "Completed"

            if all(value == "N/A" for key, value in extracted_data.items() if key != "File Name"):
                raise ValueError(f"No information extracted from the file {file}")

            results.append(extracted_data)

        except Exception as e:
            print(f"Erro ao processar {file}: {e}")
            error_data = {key: "N/A" for key in patterns.keys()} # Sets all values to “N/A”
            error_data["File Name"] = file
            error_data["Status"] = f"Erro: {e}"
            results.append(error_data)

    return results
'''📌 Explanation
- Reads all the PDFs in the directory and extracts the text from each one.
- Concatenates the text of all the pages to ensure that nothing is lost.
- Checks if any data has been extracted before storing.
- In the event of an error, the file is still registered, but with the status “Error”.'''

def save_to_excel(data):
    # Saves the extracted data in an Excel file.
    wb = Workbook()
    ws = wb.active
    ws.title = "Extracted Data"
    # Writing headers
    headers = list(data[0].keys()) # Gets the columns automatically
    ws.append(headers)

    for row in data:
        ws.append(list(row.values())) # Add each row in Excel

    ws.auto_filter.ref = ws.dimensions # Apply automatic filter

    timestamp = datetime.now().strftime("%Y-%m-%d_%H-%M-%S")
    filename = f"Extracted_Data_{timestamp}.xlsx"

    wb.save(filename)
    print(f"Excel file saved as {filename}")
'''📌 Explanation
- Dynamically creates columns in Excel based on the extracted data.
- Automatically applies filters to make it easier to navigate in Excel.
- Saves the file with a timestamp to avoid overlap.'''

def save_to_database(data):
    # Connects to the database and inserts the extracted data.
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
'''📌 Explanation
- Connects to the MySQL database.
- Inserts the extracted data into the invoice_records table.
- Uses with to ensure the connection is closed automatically.'''

def main():
    # Main function for carrying out the entire process.
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
'''📌 Explanation
- Controls the general flow of the program.
- Calls the functions in the correct order:
- Processes the PDFs.
- Saves the data in Excel.
- Inserts the data into the database.
- Catches global exceptions to avoid complete failures.'''
