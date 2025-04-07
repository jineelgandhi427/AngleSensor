import serial
import time
import re
import csv
from datetime import datetime
import pandas as pd

# Configuration
SERIAL_PORT = 'COM14'
BAUDRATE = 115200 # Communication rate
CYCLE_DURATION_MIN = 1  # Approx duration of one cycle in minutes
ENCODER_PPR = 40000  # Total pulse per rotation of encoder in the system

# ---------------------SET SYSTEM RUN TIME-------------------------------------------------------------------------
RUN_SYSTEM_MIN = 1
# -----------------------------------------------------------------------------------------------------------------

# Regex pattern to parse data
SENSOR_DATA_REGEX = r"\s*(\d+)\s+(-?\d+)\s+(\d+)\s+(\d+)\s+(\d+)\s+(\d+)"
CYCLE_ENDED_INDICATOR = "Measurements ended for the cycle"
CYCLE_START_OK = "You have selected option 2 the main program"
ENCODER_0_360_COUNT_REGEX = r"0°-360° cw\s+(-?\d+)"
ENCODER_360_0_COUNT_REGEX = r"360°-0° ccw\s+(-?\d+)"

# Derived values
NUM_CYCLES = max(1, int(RUN_SYSTEM_MIN / CYCLE_DURATION_MIN))
DATA_FILENAME = f"measurement_log_{datetime.now().strftime('%Y%m%d_%H%M%S')}.csv"
LOG_FILENAME = f"cycle_log_{datetime.now().strftime('%Y%m%d_%H%M%S')}.txt"
ENCODER_FILENAME = f"encoder_errors_{datetime.now().strftime('%Y%m%d_%H%M%S')}.csv"

# Time and Date format of the system
DATE_TIME_FORMAT = "%Y-%m-%d %H:%M:%S"

# Global variables
system_start = time.time()  # The variable is used to note down the system start time, it is not noted here
cycle_count = 0  # To store the current cycle number
notedown_system_start_time = True
encoder_error_after_CW_prev = 0  # To store the previously calculated total error after each CW operation


def parse_sensor_data(line, csv_writer, csvfile):
    global cycle_count

    def fetch_sensor_data():
        match = re.match(SENSOR_DATA_REGEX, line)
        if match:
            return {
                "step": int(match.group(1)),
                "encoder": int(match.group(2)),
                "SIN_P": int(match.group(3)),
                "COS_P": int(match.group(4)),
                "SIN_N": int(match.group(5)),
                "COS_N": int(match.group(6)),
                "cycle": cycle_count
            }
        return None

    parsed_data = fetch_sensor_data()
    if parsed_data:
        timestamp = datetime.now().strftime(DATE_TIME_FORMAT)
        print(f"Step {parsed_data['step']:>2} | Time {timestamp} | "
              f"Enc={parsed_data['encoder']:>5} | SIN_P={parsed_data['SIN_P']} | "
              f"COS_P={parsed_data['COS_P']} | SIN_N={parsed_data['SIN_N']} | "
              f"COS_N={parsed_data['COS_N']} | cycle={cycle_count}")

        csv_writer.writerow([
            parsed_data['step'], timestamp, parsed_data['encoder'],
            parsed_data['SIN_P'], parsed_data['COS_P'],
            parsed_data['SIN_N'], parsed_data['COS_N'],
            cycle_count
        ])
        csvfile.flush()


def wait_for_ack(ser, expected_str):
    """ Wait for Arduino to confirm cycle start """
    global system_start, notedown_system_start_time
    while True:
        line = ser.readline().decode('utf-8', errors='ignore').strip()
        if expected_str in line:
            if notedown_system_start_time:  # only note down the start time once in the begining.
                system_start = time.time()
                notedown_system_start_time = False
            return
        elif line:
            print(f"Waiting for Arduino to start cycle: {cycle_count}...")


def log_event(log_file, message):
    timestamp = datetime.now().strftime(DATE_TIME_FORMAT)
    log_file.write(f"[{timestamp}] {message}\n")


def parse_encoder_errors(ser, csv_encoder_error_writer, encoder_file):
    '''
    Calculation of Encoder error:
    CSV file columns -> [Step0_of_CCW, Step0_of_CW]
    if, Cycle=1 → [0, CCW - PPR]
    else → [ PE, (PE + CCW) - PPR]
    where, PE = Previous Error, PPR = Pulse Per Relovution

    For, example:
    Cycle1 -> [0, (CCW1 - PPR) ]
    Cycle2 -> [ PE, (PE + CCW2) - PPR] (PE = CCW1 + CW1)
    Cycle3 -> [ PE, (PE + CCW3) - PPR] (PE = CCW1 + CW1 + CCW2 + CW2)
    ..................
    '''
    global cycle_count, encoder_error_after_CW_prev
    encoder6_count_stored = False
    encoder7_count_stored = False
    # Run in a loop until encoder values are not received
    while not encoder6_count_stored or not encoder7_count_stored:
        line = ser.readline().decode('utf-8', errors='ignore').strip()
        if match := re.search(ENCODER_0_360_COUNT_REGEX, line):
            current_encoder_error_CCW = int(match.group(1))
            encoder6_count_stored = True
        elif match := re.search(ENCODER_360_0_COUNT_REGEX, line):
            current_encoder_error_CW = int(match.group(1))
            encoder7_count_stored = True

    # Calculating total error after cw rotation and adding previous error, this will be used in next loop
    encoder_error_after_CW = encoder_error_after_CW_prev + current_encoder_error_CCW + current_encoder_error_CW

    # Save the encoder error in a csv file
    if (cycle_count == 1):
        csv_encoder_error_writer.writerow([cycle_count, 0, current_encoder_error_CCW - ENCODER_PPR])  # [0, CCW - PPR]
    else:
        encoder_error_after_CCW = (encoder_error_after_CW_prev + current_encoder_error_CCW) - ENCODER_PPR
        csv_encoder_error_writer.writerow([cycle_count, encoder_error_after_CW_prev,
                                          encoder_error_after_CCW])  # [ PREVIOUS_ERR, (PREVIOUS_ERR + CCW) - PPR]
    encoder_file.flush()
    encoder_error_after_CW_prev = encoder_error_after_CW  # Storing the value of previous error
    return [current_encoder_error_CCW, current_encoder_error_CW]


def fit_encoder_error_to_measurements(encoder_error_file_path, measurements_file_path):
    # Load both CSVs
    encoder_error = pd.read_csv(encoder_error_file_path)
    measurements = pd.read_csv(measurements_file_path)
    # Iterate through each cycle
    for idx, row in encoder_error.iterrows():
        cycle_num = row['cycle']
        error_ccw = row['encoder_error_ccw']
        error_cw = row['encoder_error_cw']
        # Find all step 0 entries for this cycle
        step0_rows = measurements[(measurements['cycle'] == cycle_num) & (measurements['step'] == 0)]
        # Get the indexes
        first_step0_idx = step0_rows.index[0]
        second_step0_idx = step0_rows.index[1]
        # Update the encoder values
        measurements.at[first_step0_idx, 'encoder'] = error_ccw
        measurements.at[second_step0_idx, 'encoder'] = error_cw
    # Overwrite the existing file
    measurements.to_csv(measurements_file_path, index=False)
    print(f"Done...")
    print("")


def main():
    global system_start, cycle_count
    try:
        with serial.Serial(SERIAL_PORT, BAUDRATE, timeout=1) as ser, \
                open(DATA_FILENAME, mode='w', newline='', buffering=1) as csvfile, \
                open(LOG_FILENAME, mode='w') as log_file, \
                open(ENCODER_FILENAME, mode='w', newline='', buffering=1) as encoder_file:

            # Initiate data csv file
            csv_writer = csv.writer(csvfile)
            csv_writer.writerow(["step", "timestamp", "encoder", "SIN_P", "COS_P", "SIN_N", "COS_N", "cycle"])

            # Initiate encoder csv file
            csv_encoder_error_writer = csv.writer(encoder_file)
            csv_encoder_error_writer.writerow(
                ["cycle", "encoder_error_ccw", "encoder_error_cw"])

            print(f"Starting system...\nLogging to: {DATA_FILENAME}")
            print(f"Cycle log: {LOG_FILENAME}")
            print(f"System runtime: {RUN_SYSTEM_MIN} min → Estimated cycles: {NUM_CYCLES}")

            log_event(log_file, f"System started. Runtime goal: {RUN_SYSTEM_MIN} min")

            time.sleep(2)
            ser.reset_input_buffer()
            ser.reset_output_buffer()

            full_system_start = datetime.now()

            while (time.time() - system_start) <= RUN_SYSTEM_MIN * 60:
                # Trigger a new cycle
                ser.write(b'2\n')
                cycle_count += 1
                wait_for_ack(ser, CYCLE_START_OK)

                log_event(log_file, f"Cycle {cycle_count} started")

                print(f"\n{'*'*50}\nStarting cycle {cycle_count}\n{'*'*50}")

                while True:
                    line = ser.readline().decode('utf-8', errors='ignore').strip()
                    if not line:
                        continue

                    parse_sensor_data(line, csv_writer, csvfile)

                    if CYCLE_ENDED_INDICATOR in line:
                        encoder_errors = parse_encoder_errors(ser, csv_encoder_error_writer, encoder_file)
                        log_event(log_file, f"Cycle {cycle_count} ended -> {encoder_errors}")
                        print(f"\n{'*'*50}\nFinished cycle {cycle_count}\n{'*'*50}")
                        break

            full_system_end = datetime.now()
            total_runtime_min = round((full_system_end - full_system_start).total_seconds() / 60, 2)

            print("====================================================")
            print(f"\nSystem completed. Total cycles: {cycle_count}")
            print(f"Data logged to: {DATA_FILENAME}")
            print(f"Logs are saved in: {LOG_FILENAME}")
            print(f"Encoder error file saved as   : {ENCODER_FILENAME}")
            print("====================================================")

            # Final summary to be logged
            log_file.write("\n================ SYSTEM SUMMARY ===============\n")
            log_file.write(f"Total cycles completed: {cycle_count}\n")
            log_file.write(f"Data logged to        : {DATA_FILENAME}\n")
            log_file.write(f"Logs are saved in     : {LOG_FILENAME}\n")
            log_file.write(f"System started at     : {full_system_start.strftime(DATE_TIME_FORMAT)}\n")
            log_file.write(f"System ended at       : {full_system_end.strftime(DATE_TIME_FORMAT)}\n")
            log_file.write(f"Total time elapsed    : {total_runtime_min} minutes\n")

            print("Fitting encoder counts...")
            log_file.write(f"Start fitting Encoder counts -> {time.strftime(DATE_TIME_FORMAT)}\n")
            fit_encoder_error_to_measurements(ENCODER_FILENAME, DATA_FILENAME)
            log_file.write(f"Finished fitting Encoder counts -> {time.strftime(DATE_TIME_FORMAT)}\n")
            log_file.write(f"Encoder error file saved as   : {ENCODER_FILENAME}\n")
            log_file.write("====================================================\n")

    except KeyboardInterrupt:
        print(f"\nInterrupted by user. Data saved to {DATA_FILENAME}")
    except serial.SerialException as e:
        print(f"Serial connection error: {e}")


if __name__ == "__main__":
    main()
