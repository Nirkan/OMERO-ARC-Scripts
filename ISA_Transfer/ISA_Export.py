import omero
from omero.gateway import BlitzGateway
import pandas as pd
import getpass
import re


def fetch_metadata_from_project(conn, object_type, object_id):
    """Retrieve metadata key-value pairs from an OMERO Project as a nested dictionary."""
    obj = conn.getObject(object_type, object_id)
    if not obj:
        print("Invalid Project ID.")
        return {}, {}, {}, {}
    
    metadata_investigation = {}
    metadata_study = {}
    metadata_assay = {}
    other_metadata = {}
    
    isa_namespaces_investigation = [
        "ARC:ISA:INVESTIGATION:ONTOLOGY SOURCE REFERENCE",
        "ARC:ISA:INVESTIGATION:INVESTIGATION",
        "ARC:ISA:INVESTIGATION:INVESTIGATION PUBLICATIONS",
        "ARC:ISA:INVESTIGATION:INVESTIGATION CONTACTS",
        "ARC:ISA:INVESTIGATION:STUDY",
        "ARC:ISA:INVESTIGATION:INVESTIGATION PUBLICATIONS: STUDY DESIGN DESCRIPTORS",
        "ARC:ISA:INVESTIGATION:STUDY PUBLICATIONS",
        "ARC:ISA:INVESTIGATION:STUDY FACTORS",
        "ARC:ISA:INVESTIGATION:STUDY ASSAYS",
        "ARC:ISA:INVESTIGATION:STUDY PROTOCOLS",
        "ARC:ISA:INVESTIGATION:STUDY CONTACTS"
    ]
  
    isa_namespaces_study = [
        "ARC:ISA:STUDY:STUDY",
        "ARC:ISA:STUDY:STUDY DESIGN DESCRIPTORS",
        "ARC:ISA:STUDY:STUDY PUBLICATIONS",
        "ARC:ISA:STUDY:STUDY FACTORS",
        "ARC:ISA:STUDY:STUDY ASSAYS",
        "ARC:ISA:STUDY:STUDY PROTOCOLS",
        "ARC:ISA:STUDY:STUDY CONTACTS"
    ]

    isa_namespaces_assay = [
        "ARC:ISA:ASSAY:ASSAY",
        "ARC:ISA:ASSAY:ASSAY PERFORMERS",
        "ARC:ISA:ASSAY:ASSAY DESIGN DESCRIPTORS",
        "ARC:ISA:ASSAY:ASSAY PUBLICATIONS",
        "ARC:ISA:ASSAY:ASSAY FACTORS",
        "ARC:ISA:ASSAY:ASSAY PROTOCOLS",
        "ARC:ISA:ASSAY:ASSAY DATA FILES",
        "ARC:ISA:ASSAY:ASSAY CONTACTS"
    ]
  
    for ann in obj.listAnnotations():
        namespace = ann.getNs() or "No Namespace"
        metadata_target = None
        
        if namespace in isa_namespaces_investigation:
            metadata_target = metadata_investigation
        elif namespace in isa_namespaces_study:
            metadata_target = metadata_study
        elif namespace in isa_namespaces_assay:
            metadata_target = metadata_assay
        else:
            metadata_target = other_metadata
        
        if namespace not in metadata_target:
            metadata_target[namespace] = {}
        
        value = ann.getValue()
        for item in value:
            key = item[0]
            val = list(item[1:])
            metadata_target[namespace][key] = val
          
    #Sorting the metadata result dict based on isa_namespaces list. 
    metadata_investigation_Ordered = {key: metadata_investigation[key] for key in isa_namespaces_investigation if key in metadata_investigation}
    metadata_study_Ordered = {key: metadata_study[key] for key in isa_namespaces_study if key in metadata_study}
    metadata_assay_Ordered = {key: metadata_assay[key] for key in isa_namespaces_assay if key in metadata_assay}

    #print(metadata_investigation_Ordered)
    #print(other_metadata) 
    return metadata_investigation_Ordered, metadata_study_Ordered, metadata_assay_Ordered, other_metadata



# Function to save metadata to excel
def save_metadata_to_excel(metadata_investigation_Ordered, metadata_study_Ordered, metadata_assay_Ordered, other_metadata):
    """Save all metadata categories into separate Excel files with a second sheet for other metadata."""
    files = {
        "isa.investigation.xlsx": metadata_investigation_Ordered,
        "isa.study.xlsx": metadata_study_Ordered,
        "isa.assay.xlsx": metadata_assay_Ordered
    }

    
    for filename, metadata in files.items():
        if metadata != {} :  # If the metadata is not empty, create an excel sheet.
            # Maximum length of columns. Count of number of values.
            max_len = max(len(re.findall(r"'(.*?)'",values[0])) for keys in metadata.values() for values in keys.values())
            print("Maxium number of columns :", max_len)
            with pd.ExcelWriter(filename) as writer:
                rows = []
                for namespace, kv_pairs in metadata.items():
                    # Get header value and add it as first column row.
                    header = namespace.split(":")[-1]
                    rows.append([header] + [''] * max_len)
                    #Add the keys and values.
                    for key, values in kv_pairs.items():
                        #Extract values inside single quotes as one sigle string.
                        value_str = re.findall(r"'(.*?)'", values[0])
                        # Pad the extracted value string with empty strings to match max_values.
                        value_str += [''] * (max_len - len(value_str))
                        # Create row with the key followed by value_str columns.
                        row = [key] + value_str
                        rows.append(row)
                # Create a dataframe and save it to Excel.        
                df = pd.DataFrame(rows)
                print(df)
                df.to_excel(writer, sheet_name=filename[0:3]+ "_" +filename[4:-5], index=False, header=False)   
            
        elif metadata == {}:
          print("No relevant metadata found for", filename)

    if other_metadata != {}:
            # Maximum length of columns. Count of number of values. Count vaulues inside single quotes.
            max_len = max(len(re.findall(r"'(.*?)'",values[0])) for keys in metadata_investigation_Ordered.values() for values in keys.values())
            print(max_len)
            with pd.ExcelWriter("ExtraMetadata.xlsx") as writer:  
                for namespace, kv_pairs in other_metadata.items():
                    df_other = pd.DataFrame([(key, ", ".join(value)) for key, value in kv_pairs.items()], columns=["Key", "Value"])
                    df_other.to_excel(writer, index=False, header=False)
            print('Additonal key-value pairs in the ExtaMetadata.xlsx')      


#Main function
def main():
    """Main script execution."""
    host = input("OMERO Host: ")
    username = input("OMERO Username: ")
    password = getpass.getpass("OMERO Password: ")
     
    conn = BlitzGateway(username, password, host=host)
    if not conn.connect():
        print("Error: Could not connect to OMERO. Check credentials.")
        return None
    
    object_id = input("Enter OMERO Project ID: ")
    object_type = "Project"
     
    # Call the function and capture the return values
    metadata_investigation_Ordered, metadata_study_Ordered, metadata_assay_Ordered, other_metadata = fetch_metadata_from_project(conn, object_type, object_id)
    
    conn.close()
    
    save_metadata_to_excel(metadata_investigation_Ordered, metadata_study_Ordered, metadata_assay_Ordered, other_metadata)

if __name__ == "__main__":
    main()

