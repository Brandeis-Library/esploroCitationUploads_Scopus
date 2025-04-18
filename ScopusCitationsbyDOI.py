import pandas as pd
import requests
import math
import openpyxl
from tqdm import tqdm
from requests.exceptions import Timeout, RequestException  # NEW IMPORT

# Load the DOIs from a local file
with open('dois.txt', 'r') as file:
    dois = [line.strip() for line in file.readlines()]

# Define your Scopus API key
API_KEY = 'scopus_api_key_goes_here'

# Prepare a list to hold the output data
output_data = []

# Loop through each DOI and fetch citation counts
for doi in tqdm(dois, desc="Processing DOIs", unit="DOI"):
    try:  # START OF NEW ERROR HANDLING BLOCK
        # Make a request to the Scopus API with 10-second timeout
        response = requests.get(
            f'https://api.elsevier.com/content/abstract/doi/{doi}',
            headers={'X-ELS-APIKey': API_KEY, 'Accept': 'application/json'},
            timeout=10  # ADDED TIMEOUT PARAMETER
        )
        
        # Check if the request was successful
        if response.status_code == 200:
            data = response.json()
            # Retrieve citation count
            citation_count = data.get('abstracts-retrieval-response', {}).get('coredata', {}).get('citedby-count', 'N/A')
            # Append DOI and citation count to output data
            output_data.append({'DOI': doi, 'Total Citations': citation_count})
        else:
            # Handle HTTP errors
            output_data.append({'DOI': doi, 'Total Citations': 'Error'})
            
    except Timeout:  # SPECIFIC TIMEOUT HANDLING
        output_data.append({'DOI': doi, 'Total Citations': 'Error'})
    except RequestException:  # HANDLE OTHER REQUEST ERRORS
        output_data.append({'DOI': doi, 'Total Citations': 'Error'})

# REST OF THE ORIGINAL CODE REMAINS THE SAME
# Create a DataFrame
output_df = pd.DataFrame(output_data)

# Calculate the number of files needed
num_files = math.ceil(len(output_df) / 999)

# Split the DataFrame and export to Excel files
for i in range(num_files):
    start_idx = i * 999
    end_idx = min((i + 1) * 999, len(output_df))
    
    # Create a subset of the DataFrame
    subset_df = output_df.iloc[start_idx:end_idx]
    
    # Export to Excel
    subset_df.to_excel(f'citation_counts_{i+1}.xlsx', index=False, engine='openpyxl')

print(f"Data has been split and exported into {num_files} Excel files.")
