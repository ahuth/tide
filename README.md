# tide

Tide data analysis. Outputs the the location of high and low signal, plus a chart.

## Requirements

- Python 3
- 1 or more excel files with data
- Each excel file should have:
  - A sheet named `R73157 RFCURRENT - Data`
  - Columns named `Date` and `Current (mA)`

## Setup

1. Create a virtual env

   ```sh
   python3 -m venv env
   ```

2. Activate the virtual env

   ```sh
   source env/bin/activate
   ```

3. Install dependencies

   ```sh
   pip install -r requirements.txt
   ```

## Usage

Pass a list of files to process to tide.py

```sh
python tide.py file1.xlsx file2.xlsx
```

As a result, these files will be created:
- `result.xlsx`, with all the data and local minima and maxima identifed.
- `tidal_data_plot.png`, with a graph of smoothed data.
