# DOI Citation Count Fetcher

This Python script reads a list of DOIs (Digital Object Identifiers) from a text file, queries the Scopus API to retrieve the number of citations for each DOI, and exports the results to one or more Excel files.

## Features

- Reads DOIs from a simple text file (`dois.txt`)
- Fetches citation counts from the Scopus API
- Handles network errors and timeouts gracefully
- Splits results into multiple Excel files if necessary (max 999 entries per file)
- Progress bar for tracking script progress


## Requirements

Before running the script, make sure you have the following Python packages installed:

- `pandas`
- `requests`
- `openpyxl`
- `tqdm`

You can install them using pip:

```bash
pip install pandas requests openpyxl tqdm
```


## Usage

1. **Prepare your DOIs:**
    - Create a file named `dois.txt` in the same directory as the script.
    - Add one DOI per line. Example:

```
10.1016/j.cell.2020.01.001
10.1038/s41586-020-2649-2
```

2. **Set your Scopus API key:**
    - The script uses a hardcoded API key (`API_KEY`). Replace the value with your own Scopus API key if necessary.
3. **Run the script:**
    - In your terminal or command prompt, run:

```bash
python your_script_name.py
```

4. **Check your results:**
    - The script will create one or more Excel files named `citation_counts_1.xlsx`, `citation_counts_2.xlsx`, etc., each containing the DOI and its total citation count.

## Notes

- If a DOI cannot be processed (due to network issues or API errors), the script will record "Error" for its citation count.
- The script splits output into files with up to 999 records each to avoid Excel limitations.
- You need a valid Scopus API key to use this script. You can obtain one from [Elsevier Developer Portal](https://dev.elsevier.com/).


## Example Output

| DOI | Total Citations |
| :-- | :-- |
| 10.1016/j.cell.2020.01.001 | 42 |
| 10.1038/s41586-020-2649-2 | 57 |
| 10.1000/invalid-doi | Error |

## Troubleshooting

- **API Key Issues:** Make sure your API key is valid and has not expired.
- **Network Errors:** Ensure you have a stable internet connection.
- **Missing Packages:** If you get an ImportError, install the required packages with pip.


## License

This project is provided as-is for educational and research purposes.
