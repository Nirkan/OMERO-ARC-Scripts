# coding=utf-8
'''
Upload metadata from selected ISA Excel file to OMERO.

This script prompts the user to:
- Choose a file type: Investigation, Study, or Assay
- Provide OMERO credentials (host, username, password)
- Enter Project ID (for Investigation/Study) or Dataset ID (for Assay)
- Extract metadata and upload it as key-value pairs to OMERO
'''

import omero
from omero.gateway import BlitzGateway, MapAnnotationWrapper
from ezomero import post_map_annotation
import pandas as pd
import os
import getpass

def prompt_user_input(prompt, optional=True, hidden=False):
    """Helper function to get user input with optional hidden input."""
    while True:
        if hidden:
            value = getpass.getpass(prompt)
        else:
            value = input(prompt)
        if value or not optional:
            return value
        print("This field is optional. Press Enter to skip.")

def extract_metadata_from_xlsx(file_path, file_type):
    """Extracts metadata from an Excel (.xlsx) file using pandas."""
    metadata = {}
    try:
        xls = pd.ExcelFile(file_path)
        for sheet_name in xls.sheet_names:
            df = xls.parse(sheet_name, dtype=str, header=None).fillna("")
            namespace = ""
            
            for _, row in df.iterrows():
                if row.iloc[0].strip().isupper():  # Ensure first column with uppercase is included as namespace
                    namespace = f"ARC:ISA:{file_type.upper()}:{row.iloc[0].strip()}"
                    metadata[namespace] = {}
                elif namespace:
                    key = row.iloc[0].strip()
                    values = [str(v).strip() if str(v).strip() else '' for v in row[1:]]
                    
                    # Format values
                    if key:
                        # columns as comma(,) seperated and quoted(('') seperated.
                        formatted_values = [f"'{v}'" if v else "' '" for v in values]  # Ensure all values are quoted, even empty ones.
                        #Remove only trailing '' in formatted_values. Peserves leading & middle '' with comma seperated.
                        while formatted_values and formatted_values[-1] == "' '":
                            formatted_values.pop()
                        # Now join the formatted values into a comma-seperated string.    
                        metadata[namespace][key] = ", ".join(formatted_values)  # join the comma seperated and quoted values.


    except Exception as e:
        print(f"Error reading {file_path}: {e}")
    return metadata


def apply_metadata(obj, metadata, conn):
    """Apply metadata as key-value pairs to an OMERO object using ezomero."""
    obj_id = obj.getId()  # Get OMERO object ID
    print(obj_id)
    print(obj.OMERO_CLASS)
    obj_type = obj.OMERO_CLASS
    print(obj_type)
    for namespace, values in metadata.items():
        print(namespace)
        print(values)
        post_map_annotation(conn, object_type=obj_type, object_id=obj_id, kv_dict=values, ns=namespace)  # Ensure correct argument order


def main():
    print("\n--- OMERO Metadata Uploader ---")
    file_type = ""
    while file_type not in ["Investigation", "Study", "Assay"]:
        file_type = input("Select file to upload (Investigation, Study, Assay): ").strip()
    
    host = prompt_user_input("OMERO Host: ", optional=False)
    username = prompt_user_input("OMERO Username: ", optional=False)
    password = prompt_user_input("OMERO Password: ", optional=False, hidden=True)

    # Connect to OMERO
    conn = BlitzGateway(username, password, host=host)
    if not conn.connect():
        print("Error: Could not connect to OMERO. Check credentials.")
        return
    
    file_path = prompt_user_input(f"Path to isa.{file_type.lower()}.xlsx: ", optional=False)
    
    if file_type in ["Investigation", "Study"]:
        object_id = prompt_user_input("OMERO Project ID: ", optional=False)
        omero_object = conn.getObject("Project", object_id)
        print(omero_object)
    else:
        object_id = prompt_user_input("OMERO Dataset ID: ", optional=False)
        omero_object = conn.getObject("Dataset", object_id)
        print(omero_object)
    
    if not omero_object:
        print("Invalid OMERO ID. Exiting.")
        return
    
    if os.path.exists(file_path):
        metadata = extract_metadata_from_xlsx(file_path, file_type)
        apply_metadata(omero_object, metadata, conn)
        print(f"Metadata from {file_path} uploaded to OMERO {file_type} with ID {object_id}.")
    else:
        print("File not found. Exiting.")
    
    print("Upload complete!")
    conn.close()

if __name__ == "__main__":
    main()

