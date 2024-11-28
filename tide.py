import pandas as pd
from scipy.signal import savgol_filter, find_peaks
import matplotlib.pyplot as plt
import argparse

def process_tidal_data(input_files, output_file='result.xlsx'):
    # Step 1: Load specified Excel files
    all_data = []
    for file in input_files:
        try:
            # Read the relevant sheet and columns
            df = pd.read_excel(file, sheet_name='R73157 RFCURRENT - Data', usecols=['Date', 'Time', 'Values'])
            # Combine Date and Time into a single datetime column
            df['datetime'] = pd.to_datetime(df['Date'].astype(str) + ' ' + df['Time'].astype(str))
            all_data.append(df[['datetime', 'Values']])
        except Exception as e:
            print(f"Error processing {file}: {e}")

    # Check if any data was successfully loaded
    if not all_data:
        print("No valid data files provided.")
        return

    # Step 2: Combine and sort data
    combined_data = pd.concat(all_data, ignore_index=True)
    combined_data = combined_data.sort_values(by='datetime').reset_index(drop=True)

    # Step 3: Apply Savitzky–Golay filter for smoothing
    window_size = 25  # Adjust for smoothing (should be odd)
    poly_order = 2    # Quadratic smoothing
    smoothed_values = savgol_filter(combined_data['Values'], window_size, poly_order)

    # Step 4: Identify extrema
    peaks, _ = find_peaks(smoothed_values, distance=144)  # 144 = ~12 hours for 5-min data
    troughs, _ = find_peaks(-smoothed_values, distance=144)

    # Step 5: Add extrema labels
    combined_data['Smoothed'] = smoothed_values
    combined_data['Extrema'] = 'None'
    combined_data.loc[peaks, 'Extrema'] = 'Max'
    combined_data.loc[troughs, 'Extrema'] = 'Min'

    # Step 6: Visualization
    plt.figure(figsize=(12, 6))
    plt.plot(combined_data['datetime'], combined_data['Values'], label='Original Data', alpha=0.6, color='blue')
    plt.plot(combined_data['datetime'], smoothed_values, label='Smoothed Data', color='orange', linewidth=2)
    plt.scatter(combined_data['datetime'].iloc[peaks], smoothed_values[peaks], label='Maxima', color='red', marker='o')
    plt.scatter(combined_data['datetime'].iloc[troughs], smoothed_values[troughs], label='Minima', color='green', marker='o')
    plt.title('Tidal Data: Original, Smoothed, and Extrema')
    plt.xlabel('Datetime')
    plt.ylabel('Tide Level')
    plt.legend()
    plt.tight_layout()
    plt.savefig('tidal_data_plot.png')  # Save the plot
    plt.show()  # Display the plot

    # Step 7: Save to a new Excel file
    with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
        combined_data.to_excel(writer, index=False, sheet_name='Processed Data')
    print(f"Results saved to {output_file}")

# Main function to handle command-line arguments
if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="Process tidal data from Excel files.")
    parser.add_argument('input_files', nargs='+', help='List of Excel files to process')
    parser.add_argument('--output', default='result.xlsx', help='Output Excel file (default: result.xlsx)')
    args = parser.parse_args()

    # Call the processing function with the input files
    process_tidal_data(args.input_files, output_file=args.output)
