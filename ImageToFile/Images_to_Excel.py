import omero
from omero.gateway import BlitzGateway
import pandas as pd
import getpass


def extract_image_metadata(conn, dataset_id):
    dataset = conn.getObject("Dataset", dataset_id)
    if dataset is None:
        print(f"Dataset with ID {dataset_id} not found.")
        conn.close()
        exit()
    
    data = [] # List to store all image metadata.
  
    for image in dataset.listChildren():
        image_name = image.getName()
        X = image.getSizeX()
        Y = image.getSizeY()
        Z = image.getSizeZ()
        C = image.getSizeC()
        T = image.getSizeT()

        # Extract key-value paris (Map Annotations)
        kv_pairs = {}
        for ann in image.listAnnotations():
          if isinstance(ann, omero.gateway.MapAnnotationWrapper):
              for key, value in ann.getValue():
                  kv_pairs[key] = value

        # Create a dictionary for each image
        row = {
            "ImageName": image_name,
            "PixelSizeX": X,
            "PixelSizeY": Y,
            "PixelSizeZ": Z,
            "Channels": C,
            "TimeAxis": T,
        }
        row.update(kv_pairs)  # Add dynamic key-value pairs

        data.append(row)  # Append the structured dictionary

    return data # Return a list of dictionaries 
    

def main():
    # Prompt user to input OMERO credentials.
    host = input("Enter OMERO host: ")
    username = input("Enter OMERO username: ")
    password = getpass.getpass("Enter OMERO password: ")
  
    # Connect to OMERO
    conn = BlitzGateway(username, password, host=host)
    if conn.connect():
        print("Connected to OMERO successfully!")
        return conn
    else:
        print("Failed to connect to OMERO. Check your credentials.")

    # Get dataset ID.
    dataset_id = input("Enter Dataset ID: ")
    
    # Extract the Image metatadata
    data = extract_image_metadata(conn, dataset_id)
    
    # Save to excel
    df = pd.DataFrame(data)
    dataset = conn.getObject("Dataset", dataset_id)
    dataset_name = dataset.getName().replace(" ","_")   # OMERO dataset name. spaces replaced by underscore.
    excel_filename = f"{dataset_name}.xlsx"
    df.to_excel(excel_filename, index=False)
    print(f"Excel file saved successfully: {excel_filename}")
    conn.close()


if __name__ == "__main__":
    main()
