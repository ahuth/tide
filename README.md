# tide

Tide data analysis. Outputs the the location of high and low signal, plus a chart.

## Setup

1. Install Python 3

2. Create a virtual env

   ```sh
   python3 -m venv env
   ```

3. Activate the virtual env

   ```sh
   source env/bin/activate
   ```

4. Install dependencies

   ```sh
   pip install -r requirements.txt
   ```

## Usage

Pipe a list of files to process to tide.py

```sh
file1.xlsx file2.xlsx | tide.py
```

These files will be created:
- result.xlsx, with all the data and local minima and maxima identifed.
- tidal_data_plot.png, with the data and smoothed data on a graph.
