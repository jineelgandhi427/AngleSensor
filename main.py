import serial
import time
import re
import csv
from datetime import datetime

# Configuration
SERIAL_PORT = '/dev/ttyACM0'
BAUDRATE = 115200
CYCLE_DURATION_MIN = 1  # Approx duration of one cycle in minutes

#---------------------SET SYSTEM RUN TIME-------------------------------------------------------------------------
RUN_SYSTEM_MIN = 1
#-----------------------------------------------------------------------------------------------------------------

# Regex pattern to parse sensor data
SENSOR_DATA_REGEX = r"\s*(\d+)\s+(-?\d+)\s+(\d+)\s+(\d+)\s+(\d+)\s+(\d+)"
CYCLE_ENDED_INDICATOR = "Measurements ended for the cycle"
CYCLE_START_OK = "You have selected option 2 the main program"

# Derived values
NUM_CYCLES = max(1, int(RUN_SYSTEM_MIN / CYCLE_DURATION_MIN))
CSV_FILENAME = f"measurement_log_{datetime.now().strftime('%Y%m%d_%H%M%S')}.csv"
LOG_FILENAME = f"cycle_log_{datetime.now().strftime('%Y%m%d_%H%M%S')}.txt"

# Time and Date format of the system
DATE_TIME_FORMAT = "%Y-%m-%d %H:%M:%S"

# Global variables
system_start = time.time()
cycle_count = 0
notedown_system_start_time = True

def parse_sensor_line(line):
    global cycle_count
    print(f"Line: {line}")
    match = re.match(SENSOR_DATA_REGEX, line)
    if match:
        print("Matched")
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

def wait_for_ack(ser, expected_str):
    """ Wait for Arduino to confirm cycle start """
    global system_start, notedown_system_start_time
    while True:
        line = ser.readline().decode('utf-8', errors='ignore').strip()
        if expected_str in line:
            if notedown_system_start_time: # only note down the start time once in the begining.
                system_start = time.time()
                notedown_system_start_time = False
            return
        elif line:
            print(f"Waiting for Arduino to start cycle: {cycle_count}...")

def log_event(log_file, message):
    timestamp = datetime.now().strftime(DATE_TIME_FORMAT)
    log_file.write(f"[{timestamp}] {message}\n")

def main():
    global system_start, cycle_count
    try:
        with serial.Serial(SERIAL_PORT, BAUDRATE, timeout=1) as ser, \
             open(CSV_FILENAME, mode='w', newline='', buffering=1) as csvfile, \
             open(LOG_FILENAME, mode='w') as log_file:

            csv_writer = csv.writer(csvfile)
            csv_writer.writerow(["step", "timestamp", "encoder", "SIN_P", "COS_P", "SIN_N", "COS_N", "cycle"])

            print(f"Starting system...\nLogging to: {CSV_FILENAME}")
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

                    if CYCLE_ENDED_INDICATOR in line:
                        log_event(log_file, f"Cycle {cycle_count} ended")
                        print(f"\n{'*'*50}\nFinished cycle {cycle_count}\n{'*'*50}")
                        break

                    parsed = parse_sensor_line(line)
                    if parsed:
                        timestamp = datetime.now().strftime(DATE_TIME_FORMAT)
                        print(f"Step {parsed['step']:>2} | Time {timestamp} | "
                              f"Enc={parsed['encoder']:>5} | SIN_P={parsed['SIN_P']} | "
                              f"COS_P={parsed['COS_P']} | SIN_N={parsed['SIN_N']} | "
                              f"COS_N={parsed['COS_N']} | cycle={cycle_count}")

                        csv_writer.writerow([
                            parsed['step'], timestamp, parsed['encoder'],
                            parsed['SIN_P'], parsed['COS_P'],
                            parsed['SIN_N'], parsed['COS_N'],
                            cycle_count
                        ])
                        csvfile.flush()

            full_system_end = datetime.now()
            total_runtime_min = round((full_system_end - full_system_start).total_seconds() / 60, 2)

            print(f"\nSystem completed. Total cycles: {cycle_count}")
            print(f"Data logged to: {CSV_FILENAME}")
            print(f"Logs are saved in: {LOG_FILENAME}")

            # Final summary to be logged
            log_file.write("\n========= SYSTEM SUMMARY =========\n")
            log_file.write(f"Total cycles completed: {cycle_count}\n")
            log_file.write(f"System started at     : {full_system_start.strftime(DATE_TIME_FORMAT)}\n")
            log_file.write(f"System ended at       : {full_system_end.strftime(DATE_TIME_FORMAT)}\n")
            log_file.write(f"Total time elapsed    : {total_runtime_min} minutes\n")
            log_file.write("==================================\n")

    except KeyboardInterrupt:
        print(f"\nInterrupted by user. Data saved to {CSV_FILENAME}")
    except serial.SerialException as e:
        print(f"Serial connection error: {e}")

if __name__ == "__main__":
    main()
