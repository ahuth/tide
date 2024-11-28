import numpy as np
import pandas as pd
from scipy.fft import fft, ifft, fftfreq
from scipy.signal import find_peaks
import matplotlib.pyplot as plt
import argparse

def process_tidal_data(input_files, output_file='result.xlsx'):
    # Step 1: Load specified Excel files
    all_data = []
    for file in input_files:
        try:
            # Read the relevant sheet, skipping the first 6 rows and using row 7 as header
            df = pd.read_excel(file, sheet_name='R73157 RFCURRENT - Data', header=6, usecols=['Date', 'Current (mA)'])

            # Ensure that 'Date' is in proper datetime format
            if not pd.api.types.is_datetime64_any_dtype(df['Date']):
                df['Date'] = pd.to_datetime(df['Date'], errors='coerce')

            # Use 'Date' as the datetime
            df['datetime'] = df['Date']
            all_data.append(df[['datetime', 'Current (mA)']])
        except Exception as e:
            print(f"Error processing {file}: {e}")

    # Check if any data was successfully loaded
    if not all_data:
        print("No valid data files provided.")
        return

    # Step 2: Combine and sort data
    combined_data = pd.concat(all_data, ignore_index=True)
    combined_data = combined_data.sort_values(by='datetime').reset_index(drop=True)

    # Step 3: Apply Fourier Transform (FFT) to the 'Current (mA)' data
    signal = combined_data['Current (mA)'].values
    n = len(signal)
    sampling_rate = 1 / 5  # 5-minute intervals

    # Perform Fast Fourier Transform (FFT)
    fft_signal = fft(signal)
    freqs = fftfreq(n, d=sampling_rate)

    # Step 4: Identify dominant frequency
    half = n // 2  # FFT result is symmetric, so only keep the first half
    fft_signal = fft_signal[:half]
    freqs = freqs[:half]

    # Let's focus on a specific frequency range (e.g., tidal cycle ~12 hours)
    low_freq = 1 / (12 * 60)  # 12 hours in minutes (low cut frequency)
    high_freq = 1 / (4 * 60)  # 4 hours in minutes (high cut frequency)

    # Filter the signal: Retain components in the desired frequency range
    filtered_signal = np.zeros_like(signal)
    for i in range(half):
        if low_freq <= np.abs(freqs[i]) <= high_freq:
            filtered_signal += np.real(ifft(fft_signal[i] * np.exp(2j * np.pi * freqs[i] * np.arange(n) / n)))

    # Step 5: Detect Extrema (peaks and valleys) on the filtered signal
    smoothed_values = filtered_signal
    peaks, _ = find_peaks(smoothed_values, distance=144)  # ~12 hours for 5-min data
    troughs, _ = find_peaks(-smoothed_values, distance=144)

    # Step 6: Add extrema labels
    combined_data['Filtered'] = smoothed_values
    combined_data['Extrema'] = 'None'
    combined_data.loc[peaks, 'Extrema'] = 'Max'
    combined_data.loc[troughs, 'Extrema'] = 'Min'

    # Step 7: Visualization
    plt.figure(figsize=(12, 6))
    plt.plot(combined_data['datetime'], combined_data['Current (mA)'], label='Original Data', alpha=0.6, color='blue')
    plt.plot(combined_data['datetime'], smoothed_values, label='Filtered Data (Fourier)', color='orange', linewidth=2)
    plt.scatter(combined_data['datetime'].iloc[peaks], smoothed_values[peaks], label='Maxima', color='red', marker='o')
    plt.scatter(combined_data['datetime'].iloc[troughs], smoothed_values[troughs], label='Minima', color='green', marker='o')
    plt.title('Tidal Data: Original, Filtered (Fourier), and Extrema')
    plt.xlabel('Datetime')
    plt.ylabel('Tide Level')
    plt.legend()
    plt.tight_layout()
    plt.savefig('tidal_data_fourier_plot.png')  # Save the plot
    plt.show()  # Display the plot

    # Step 8: Save to a new Excel file
    with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
        combined_data.to_excel(writer, index=False, sheet_name='Processed Data')
    print(f"Results saved to {output_file}")

# Main function to handle command-line arguments
if __name__ == "__main__":
    print("Script is running...")

    parser = argparse.ArgumentParser(description="Process tidal data from Excel files.")
    parser.add_argument('input_files', nargs='+', help='List of Excel files to process')
    parser.add_argument('--output', default='result.xlsx', help='Output Excel file (default: result.xlsx)')
    args = parser.parse_args()

    # Call the processing function with the input files
    process_tidal_data(args.input_files, output_file=args.output)
