import pandas as pd
from scipy.signal import savgol_filter, find_peaks
import matplotlib.pyplot as plt
import argparse
import os
import sys

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

    # Apply the formula: y = 4.51x - 36.4
    combined_data['Altitude'] = 4.51 * combined_data['Current (mA)'] - 36.4

    # Apply Savitzky–Golay filter for smoothing
    window_size = 5   # Adjust for smoothing (should be odd). Tune for accuracy.
    poly_order = 2    # Quadratic smoothing
    smoothed_values = savgol_filter(combined_data['Altitude'], window_size, poly_order)

    # Identify extrema
    distance = 96 # ~8 hours of 5-min intervals
    peaks, _ = find_peaks(smoothed_values, distance=distance)
    troughs, _ = find_peaks(-smoothed_values, distance=distance)

    # Add extrema labels
    combined_data['Smoothed'] = smoothed_values
    combined_data['Extrema'] = ''
    combined_data.loc[peaks, 'Extrema'] = 'Max'
    combined_data.loc[troughs, 'Extrema'] = 'Min'

    # Calculate mean times between tides
    high_tides_times = combined_data.loc[peaks, 'datetime']
    low_tides_times = combined_data.loc[troughs, 'datetime']

    high_tide_deltas = high_tides_times.diff().dropna().dt.total_seconds() / 3600  # Convert to hours
    low_tide_deltas = low_tides_times.diff().dropna().dt.total_seconds() / 3600   # Convert to hours

    mean_high_tide_interval = high_tide_deltas.mean()
    mean_low_tide_interval = low_tide_deltas.mean()

    # Minimum tide threshold
    min_tide_threshold = -4

    # Calculate max high tide and min low tide
    max_high_tide_value = combined_data['Altitude'].iloc[peaks].max()
    max_high_tide_time = high_tides_times[combined_data['Altitude'].iloc[peaks].idxmax()]

    # Filter out low tide values below the threshold
    valid_troughs = combined_data['Altitude'].iloc[troughs] >= min_tide_threshold
    valid_troughs_indices = troughs[valid_troughs]

    min_low_tide_value = combined_data['Altitude'].iloc[valid_troughs_indices].min()
    min_low_tide_time = low_tides_times[combined_data['Altitude'].iloc[valid_troughs_indices].idxmin()]

    # Create the stats DataFrame
    stats_data = pd.DataFrame({
        'Metric': [
            'Mean Time Between High Tides (hours)',
            'Mean Time Between Low Tides (hours)',
            'Max High Tide (m) + Date/Time',
            'Min Low Tide (m) + Date/Time'
        ],
        'Value': [
            mean_high_tide_interval,
            mean_low_tide_interval,
            f"{max_high_tide_value:.2f} at {max_high_tide_time}",
            f"{min_low_tide_value:.2f} at {min_low_tide_time}"
        ]
    })

    # Visualization
    plt.figure(figsize=(12, 6))
    plt.plot(combined_data['datetime'], combined_data['Altitude'], label='Original Data', alpha=0.6, color='blue')
    plt.plot(combined_data['datetime'], smoothed_values, label='Smoothed Data', color='orange', linewidth=2)
    plt.scatter(combined_data['datetime'].iloc[peaks], smoothed_values[peaks], label='Maxima', color='red', marker='o')
    plt.scatter(combined_data['datetime'].iloc[troughs], smoothed_values[troughs], label='Minima', color='green', marker='o')
    plt.title('Tidal Data: Original, Smoothed, and Extrema')
    plt.xlabel('Datetime')
    plt.ylabel('Tide Level')
    plt.legend()
    plt.tight_layout()

    image_path = 'tidal_data_plot.png'
    plt.savefig(image_path)

    # Open the image using the system's default image viewer
    if os.name == 'posix':  # macOS or Linux
        os.system(f"open {image_path}" if sys.platform == 'darwin' else f"xdg-open {image_path}")
    elif os.name == 'nt':  # Windows
        os.system(f"start {image_path}")

    plt.close()

    # Save to a new Excel file
    with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
        combined_data.to_excel(writer, index=False, sheet_name='Processed Data')
        stats_data.to_excel(writer, index=False, sheet_name='Tidal Statistics')
    print(f"Results saved to {output_file}")

# Main function to handle command-line arguments
if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="Process tidal data from Excel files.")
    parser.add_argument('input_files', nargs='+', help='List of Excel files to process')
    parser.add_argument('--output', default='result.xlsx', help='Output Excel file (default: result.xlsx)')
    args = parser.parse_args()
    print(f'Running script on {len(args.input_files)} files...')
    process_tidal_data(args.input_files, output_file=args.output)
