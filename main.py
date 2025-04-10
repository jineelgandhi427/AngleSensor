import serial
import time
import re
import csv
from datetime import datetime
import pandas as pd
from maths import Formula

# Configuration
SERIAL_PORT = 'COM14'
BAUDRATE = 115200  # Communication rate
CYCLE_DURATION_MIN = 1  # Approx duration of one cycle in minutes
ENCODER_PPR = 40000  # Total pulse per rotation of encoder in the system

# ---------------------SET SYSTEM RUN TIME-------------------------------------------------------------------------
RUN_SYSTEM_MIN = 1
# -----------------------------------------------------------------------------------------------------------------

# Regex pattern to parse data (DO NOT CHANGE!)
SENSOR_DATA_REGEX = r"\s*(\d+)\s+(-?\d+)\s+(\d+)\s+(\d+)\s+(\d+)\s+(\d+)"
CYCLE_ENDED_INDICATOR = "Measurements ended for the cycle"
CYCLE_START_OK = "You have selected option 2 the main program"
ENCODER_0_360_COUNT_REGEX = r"0°-360° cw\s+(-?\d+)"
ENCODER_360_0_COUNT_REGEX = r"360°-0° ccw\s+(-?\d+)"
CW_ENCODER_ERROR_REGEX = r"Correction cw:\s*(-?\d+)"
CCW_ENCODER_ERROR_REGEX = r"Correction ccw:\s*(-?\d+)"

# Derived values (DO NOT CHANGE!)
NUM_CYCLES = max(1, int(RUN_SYSTEM_MIN / CYCLE_DURATION_MIN))
TIMESTAMP = datetime.now().strftime('%Y%m%d_%H%M%S')
DATA_FILENAME = f"measurement_log_{TIMESTAMP}.csv"
LOG_FILENAME = f"cycle_log_{TIMESTAMP}.txt"
ENCODER_FILENAME = f"encoder_errors_{TIMESTAMP}.csv"

# Time and Date format of the system (DO NOT CHANGE!)
DATE_TIME_FORMAT = "%Y-%m-%d %H:%M:%S"

# Global variables
system_start = time.time()  # The variable is used to note down the system start time, it is not noted here
cycle_count = 0  # To store the current cycle number
notedown_system_start_time = True
encoder_error_after_CCW_prev = 0  # To store the previously calculated total error after each CW operation
cw_encoder_error_diff = 0
ccw_encoder_error_diff = 0

# Maths module object
# formulas = Formula(o_x_m=O_X_M, o_y_m=O_Y_M, a_x_m=A_X_M, a_y_m=A_Y_M, phi_x_m=PHI_X_M, phi_y_m=PHI_Y_M)


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


def add_dir_col_to_data(file_path):
    df = pd.read_csv(file_path)
    directions = []
    reset_counter = 0
    for idx, row in df.iterrows():
        if row['step'] == 0 and idx != 0:
            reset_counter += 1
        if reset_counter % 2 == 0:
            directions.append('CW')
        else:
            directions.append('CCW')
    df['direction'] = directions
    df.to_csv(file_path, index=False)


def main():
    global system_start, cycle_count, formulas, cw_encoder_error_diff, ccw_encoder_error_diff
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
                ["cycle", "encoder_error_cw", "encoder_error_ccw"])

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

                    match_cw = re.search(CW_ENCODER_ERROR_REGEX, line)
                    if match_cw:
                        cw_encoder_error_diff = match_cw.group(1)
                    match_ccw = re.search(CCW_ENCODER_ERROR_REGEX, line)
                    if match_ccw:
                        ccw_encoder_error_diff = match_ccw.group(1)

                    if CYCLE_ENDED_INDICATOR in line:
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

            print("Preparing the CSV file, it might take some time...")

            # Final summary to be logged
            log_file.write("\n================ SYSTEM SUMMARY ===============\n")
            log_file.write(f"Total cycles completed: {cycle_count}\n")
            log_file.write(f"System started at     : {full_system_start.strftime(DATE_TIME_FORMAT)}\n")
            log_file.write(f"System ended at       : {full_system_end.strftime(DATE_TIME_FORMAT)}\n")
            log_file.write(f"Total time elapsed    : {total_runtime_min} minutes\n")

            # Add direction column (CW/CCW) to the data
            print("Adding direction column...")
            log_file.write(f"Adding direction column to CSV -> {time.strftime(DATE_TIME_FORMAT)}\n")
            add_dir_col_to_data(DATA_FILENAME)
            log_file.write(f"Finished adding direction -> {time.strftime(DATE_TIME_FORMAT)}\n")

            # Calculating and storing angle error
            # log_file.write(f"Started calculating Angle Error -> {time.strftime(DATE_TIME_FORMAT)}\n")
            # formulas.calculate_and_update_angle_errors(DATA_FILENAME)
            # print(f"Angle errors calculated and file updated -> {time.strftime(DATE_TIME_FORMAT)}\n")

            log_file.write(f"Logs are saved in     : {LOG_FILENAME}\n")
            log_file.write(f"Final CSV saved to  : {DATA_FILENAME}\n")
            log_file.write("====================================================\n")

    except KeyboardInterrupt:
        print(f"\nInterrupted by user. Data saved to {DATA_FILENAME}")
    except serial.SerialException as e:
        print(f"Serial connection error: {e}")


if __name__ == "__main__":
    main()
