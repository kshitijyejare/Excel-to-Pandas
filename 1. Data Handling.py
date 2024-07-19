import pandas as pd
import os
import time
import numpy as np

def process_excel_file():

    start_time = time.time()  # Start the timer

    # Define the input path and output path
    input_path = r"C:\Users\Kshit\Desktop\APMC1.xlsx"
    output_path = os.path.join(os.path.dirname(input_path), "APMC2.xlsx")
    
    # Load the Excel file into a DataFrame
    APMC = pd.read_excel(input_path)

    # Remove the first row
    DF_CLN = APMC.iloc[1:].reset_index(drop=True)

    # Trim spaces from all data
    DF_CLN = DF_CLN.applymap(lambda x: x.strip() if isinstance(x, str) else x)

    # Remove all completely blank rows
    DF_CLN.dropna(how='all', inplace=True)

    # Save the cleaned DataFrame to a new Excel file
    DF_CLN.to_excel(output_path, index=False)

    # Print success message and runtime
    end_time = time.time()  # End the timer
    runtime = end_time - start_time
    print(f"Successfully completed. Runtime: {runtime:.2f} seconds")

# Call the function to process the Excel file
process_excel_file()
