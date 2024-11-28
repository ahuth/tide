import pandas as pd
from scipy.signal import savgol_filter, find_peaks
import matplotlib.pyplot as plt
import argparse

def process_tidal_data(input_files, output_file='result.xlsx'):
    all_data = []

    # Load specified Excel files
    for file in input_files:
        try:
            # Read the relevant sheet, skipping the first 6 rows and using row 7 as header
            df = pd.read_excel(file, sheet_name='R73157 RFCURRENT - Data', header=6, usecols=['Date', 'Current (mA)'])

            # Ensure that 'Date' is in proper datetime format
            if not pd.api.types.is_datetime64_any_dtype(df['Date']):
                df['Date'] = pd.to_datetime(df['Date'], errors='coerce')

            # Add the relevant columns
            df['datetime'] = df['Date']  # The 'Date' column is used as the datetime
            all_data.append(df[['datetime', 'Current (mA)']])
        except Exception as e:
            print(f"Error processing {file}: {e}")

    # Check if any data was successfully loaded
    if not all_data:
        print("No valid data files provided.")
        return

    # Combine and sort data
    combined_data = pd.concat(all_data, ignore_index=True)
    combined_data = combined_data.sort_values(by='datetime').reset_index(drop=True)

    # Apply Savitzky–Golay filter for smoothing
    window_size = 5   # Adjust for smoothing (should be odd). Tune for accuracy.
    poly_order = 2    # Quadratic smoothing
    smoothed_values = savgol_filter(combined_data['Current (mA)'], window_size, poly_order)

    # Identify extrema
    distance = 96 # ~8 hours of 5-min intervals
    peaks, _ = find_peaks(smoothed_values, distance=distance)
    troughs, _ = find_peaks(-smoothed_values, distance=distance)

    # Add extrema labels
    combined_data['Smoothed'] = smoothed_values
    combined_data['Extrema'] = ''
    combined_data.loc[peaks, 'Extrema'] = 'Max'
    combined_data.loc[troughs, 'Extrema'] = 'Min'

    # Visualization
    plt.figure(figsize=(12, 6))
    plt.plot(combined_data['datetime'], combined_data['Current (mA)'], label='Original Data', alpha=0.6, color='blue')
    plt.plot(combined_data['datetime'], smoothed_values, label='Smoothed Data', color='orange', linewidth=2)
    plt.scatter(combined_data['datetime'].iloc[peaks], smoothed_values[peaks], label='Maxima', color='red', marker='o')
    plt.scatter(combined_data['datetime'].iloc[troughs], smoothed_values[troughs], label='Minima', color='green', marker='o')
    plt.title('Tidal Data: Original, Smoothed, and Extrema')
    plt.xlabel('Datetime')
    plt.ylabel('Tide Level')
    plt.legend()
    plt.tight_layout()
    plt.savefig('tidal_data_plot.png')
    plt.show()

    # Save to a new Excel file
    with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
        combined_data.to_excel(writer, index=False, sheet_name='Processed Data')
    print(f"Results saved to {output_file}")

# Main function to handle command-line arguments
if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="Process tidal data from Excel files.")
    parser.add_argument('input_files', nargs='+', help='List of Excel files to process')
    parser.add_argument('--output', default='result.xlsx', help='Output Excel file (default: result.xlsx)')
    args = parser.parse_args()
    print(f'Running script on {len(args.input_files)} files...')
    process_tidal_data(args.input_files, output_file=args.output)
