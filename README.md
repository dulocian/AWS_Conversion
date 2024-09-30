# AWS_Conversion
Flatten the AWS Excel file and export entries to CSV file.

## Requirements

- Python 3.6 or higher
- `openpyxl` (for reading Excel files)
- `argparse` (for command-line argument parsing)
- `csv` (for writing CSV files)

## Installation

1. Clone the repository or download the script.

    ```bash
    git clone <repo-url>
    cd <repo-directory>
    ```

2. Install the required dependencies.

    If you are using `pip` to manage dependencies, create a virtual environment and install the requirements:

    ```bash
    python -m venv venv
    source venv/bin/activate  # For Windows: venv\Scripts\activate
    pip install openpyxl
    ```

3. Make sure you have the Excel file (`.xlsx`) in the correct format. It should contain 10 columns: ID, entry, morpheme<t>, morphology<t>, ex<original>, ex<t>, reference, POS, domain, sense.

## Usage

The script can be executed via the command line with optional arguments for input and output files.

### Command-line Arguments:

- `-i, --input` (optional): Path to the input Excel file. Default is `data/AWS.xlsx`.
- `-o, --output` (optional): Path to the output CSV file. Default is `aws_converted.csv`.

### Example Usage:

1. **Using default settings:**

    This will read data from the default `data/AWS.xlsx` file and output it to `aws_converted.csv`.

    ```bash
    python main.py
    ```

2. **Specifying custom input and output files:**

    You can specify both the input Excel file and the output CSV file using the `-i` and `-o` flags.

    ```bash
    python main.py -i path/to/input.xlsx -o path/to/output.csv
    ```

3. **Help and Usage Information:**

    To get help or view available arguments, run:

    ```bash
    python main.py --help
    ```

    Example output:

    ```
    usage: main.py [-h] [-i INPUT] [-o OUTPUT]

    Flatten AWS Excel entries and export as CSV.

    optional arguments:
      -h, --help            show this help message and exit
      -i INPUT, --input INPUT
                            Input file: ASW Excel file (default: data/AWS.xlsx)
      -o OUTPUT, --output OUTPUT
                            Output file name (default: aws_converted.csv)
    ```

## How the Script Works

1. **Loading the Excel File:**
    - The script loads the Excel file using `openpyxl`.
    - It reads all the rows from the file and stores them in memory, separating the header row and the data rows.
   
2. **Flattening the Data:**
    - The data is flattened by parsing hierarchical columns and converting them into a flat format using alphabetic characters (`A`, `B`, `C`, etc.) to indicate that values are associated with a given lemma (i.e. lemma A, lemma B).
    - The numerical label indicates the row, or additional values that are also associated the given lemma (e.g. A_1 for all values on row 1 of the entry, A_2, for all values on row 2 of the entry).
   
3. **Writing to CSV:**
    - The flattened data is written to a CSV file with new headers that reflect the flattened structure.
    - For any missing data in the original file, empty strings are inserted in the CSV.

## Output

The output is a CSV file where each hierarchical entry from the Excel sheet is flattened into rows with unique column headers.

## Example Output

For an Excel file with the following structure:

| ID | Entry   | Morpheme | Morphology |
|----|---------|----------|------------|
| 89 | Eerste  | -s       | Eerstes    |
| 90 | Tweede  | -en      | Tweeden    | 
|    |         | -s       | Tweedes    | 

The script will produce a CSV file with headers like:

```csv
ID, Entry_A, Morpheme_A_1, Morphology_A_1, Entry_B, Morpheme_B_1, Morphology_B_1, Morpheme_B_2, Morphology_B_2
```

## Notes

- The script accounts for only the first 10 columns in the Excel file. You can modify this by changing the `max_col=10` parameter in the `load_excel` function.
