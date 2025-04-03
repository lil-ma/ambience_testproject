import pandas as pd
import json
from docx import Document
import os

import pandas as pd
from docx import Document
import json
from transformers import (
    Gemma3ForConditionalGeneration,
    BitsAndBytesConfig,
    AutoProcessor,
)
import torch

def process_spreadsheet(input_file, json_output, doc_output_folder, header_names):
    # Step 1: Read the spreadsheet
    df = pd.read_excel(input_file)

    # Initialize an empty list to store data for JSON
    json_data = []

    # Step 2: Loop through each row in the filtered dataframe
    for index, row in df.iterrows():
        # Initialize a dictionary to store the row data
        row_dict = {}
        non_nan_values = row[pd.notna(row)]

        # Check if the row has more non-NaN values than the header names
        if len(non_nan_values) > len(header_names):
            # If more columns than expected, concatenate all but the last two cells
            concatenated_value = " ".join(
                [str(cell) for cell in non_nan_values[:-2] if pd.notna(cell)]
            )  # Concatenate everything except last two
            row_dict[header_names[0]] = "reformatted_index_" + str(
                index
            )  # rename first colum (id)

            row_dict[header_names[1]] = (
                concatenated_value  # Assign concatenated value to the second header column
            )
            # Assign last two cells to the corresponding headers
            for i, column in enumerate(header_names[-2:], start=1):
                row_dict[column] = non_nan_values.iloc[-3 + i]
        else:
            # If the row doesn't have more columns, simply assign values as is
            for column, value in row.items():
                row_dict[column] = value

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
    with open(json_output, "w") as json_file:
        json.dump(json_data, json_file, indent=4)

    print(
        f"Processing complete. JSON output saved to {json_output}. Word documents saved in {doc_output_folder}."
    )
    return json_data


def setup_model(model_id="google/gemma-3-4b-it"):
    model = Gemma3ForConditionalGeneration.from_pretrained(
        model_id, temperature=0.2, device_map="auto"
    ).eval()
    processor = AutoProcessor.from_pretrained(model_id)
    return model, processor


def convert_transcript_to_script(transcript, model, processor):
    prompt = f"""
        You will be given a transcript of a conversation. Convert the transcript to script form.
        Transcript: {transcript}
    """
    messages = [
        {
            "role": "system",
            "content": [{"type": "text", "text": "You are a helpful assistant."}],
        },
        {
            "role": "user",
            "content": [
                {"type": "text", "text": prompt},
            ],
        },
    ]
    input_text = processor.apply_chat_template(messages, add_generation_prompt=True)
    inputs = processor(
        images=None, text=input_text, add_special_tokens=False, return_tensors="pt"
    ).to(model.device)
    with torch.inference_mode():
        output = model.generate(**inputs)
    input_len = inputs["input_ids"].shape[-1]
    output = output[0][input_len:]
    output = processor.decode(output, skip_special_tokens=True)
    return output



# Example usage
input_file = r"C:\Users\Liliana\dev\Ambience_testproject\omissions_sample_encounters.xlsx"  # Input Excel file path
json_output = "output_data.json"  # Output JSON file path
doc_output_folder = "word_documents"  # Folder to save the Word documents
model_id="google/gemma-3-4b-it"

# Create the folder for Word documents if it doesn't exist
if not os.path.exists(doc_output_folder):
    os.makedirs(doc_output_folder)

# Call the function to process the spreadsheet
json_data = process_spreadsheet(
    input_file,
    json_output,
    doc_output_folder,
    header_names=["Encounter ID", "Transcript", "HPI", "Clinical feedback"],
)

model, processor = setup_model(model_id=model_id)

for ind, transcript in enumerate(json_data):
    script = convert_transcript_to_script(transcript['Transcript'], model, processor) #use local models for now because I'm cheap
    json_data[i]['script'] = script
