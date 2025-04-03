import pandas as pd
import json
from docx import Document
import os

def process_spreadsheet(input_file, json_output, doc_output_folder, header_names):
    # Step 1: Read the spreadsheet
    df = pd.read_excel(input_file)

    # Initialize an empty list to store data for JSON
    json_data = []
    # Filter the dataframe to include only the specified header names
    df = df[header_names]
    # Step 2: Loop through each row in the dataframe
    for index, row in df.iterrows():
        # Convert the row to a dictionary
        row_dict = row.to_dict()
        
        # Append the row data to json_data list
        json_data.append(row_dict)

        # Step 3: Create a Word document for this row
        doc = Document()
        doc.add_heading(f"Row {index + 1} Data", 0)

        for column, value in row_dict.items():
            doc.add_paragraph(f"{column}: {value}")

        # Save the Word document
        doc_filename = f"{doc_output_folder}/row_{index + 1}.docx"
        doc.save(doc_filename)

    # Step 4: Write the JSON output to a file
    with open(json_output, 'w') as json_file:
        json.dump(json_data, json_file, indent=4)

    print(f"Processing complete. JSON output saved to {json_output}. Word documents saved in {doc_output_folder}.")

# Example usage
input_file = r"C:\Users\Liliana\dev\Ambience_testproject\omissions_sample_encounters.xlsx"  # Input Excel file path
json_output = 'output_data.json'  # Output JSON file path
doc_output_folder = 'word_documents'  # Folder to save the Word documents

# Create the folder for Word documents if it doesn't exist
if not os.path.exists(doc_output_folder):
    os.makedirs(doc_output_folder)

# Call the function to process the spreadsheet
process_spreadsheet(input_file, json_output, doc_output_folder, header_names=['Encounter ID', 'Transcript', 'HPI', 'Clinical feedback'])
