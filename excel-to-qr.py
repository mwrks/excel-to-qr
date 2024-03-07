import pandas as pd
import qrcode
import sys
import os

if len(sys.argv) != 3:
    print("Usage: python xlsx-to-qr.py <xlsx_file> <output_folder>")
    sys.exit(1)

script_dir = os.path.dirname(os.path.abspath(__file__))

try:
    xlsx_file = os.path.join(script_dir, sys.argv[1])
    output_folder = os.path.join(script_dir, sys.argv[2])
    # Create output folder if it doesn't exist
    if not os.path.exists(output_folder):
        os.makedirs(output_folder)  # Create the folder structure (including subfolders)

    # Print the constructed paths for debugging
    # print(f"XLSX file path: {xlsx_file}")
    # print(f"Output folder path: {output_folder}")
    df = pd.read_excel(xlsx_file)
    # print(df)
    maxRows = df.shape[0]

    for i in range(maxRows):
        try:
            number = df.loc[i, 'Number']
            img = qrcode.make(number)
            filename = os.path.join(output_folder, f"{df.loc[i, 'Name']}.png")
            img.save(filename)
            print(f"{i}.\t{filename} is successfully generated!\n")
        except KeyError:
            print(f"Row {i+1}: 'Number' or 'Name' column not found.")
        except TypeError:
            print(f"Row {i+1}: Invalid data type in 'Number' column.")

except FileNotFoundError:
    print(f"Error: '{xlsx_file}' file not found.")
