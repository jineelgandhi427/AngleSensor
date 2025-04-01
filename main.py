import serial
import time
import re
import csv
from datetime import datetime

# Configuration
SERIAL_PORT = 'COM8'
BAUDRATE = 115200
RUN_SYSTEM_MIN = 2  # Must be >1 and integer
CYCLE_DURATION_MIN = 1.5  # Approx duration of one cycle in minutes

# Regex pattern to parse sensor data
SENSOR_DATA_REGEX = r"\s*(\d+)\s+(-?\d+)\s+(\d+)\s+(\d+)\s+(\d+)\s+(\d+)\s+([0-9.]+)"
CYCLE_ENDED_INDICATOR = "Measurements ended for cycle:"
CYCLE_START_OK = "You have selected option 2 the main program"

# Derived values
NUM_CYCLES = max(1, int(RUN_SYSTEM_MIN / CYCLE_DURATION_MIN))
CSV_FILENAME = f"measurement_log_{datetime.now().strftime('%Y%m%d_%H%M%S')}.csv"

system_start = time.time()
cycle_count = 0
notedown_system_start_time = True

def parse_sensor_line(line):
    print(f"Line: {line}")
    match = re.match(SENSOR_DATA_REGEX, line)
    if match:
        return {
            "step": int(match.group(1)),
            "encoder": int(match.group(2)),
            "SIN_P": int(match.group(3)),
            "COS_P": int(match.group(4)),
            "SIN_N": int(match.group(5)),
            "COS_N": int(match.group(6)),
            "TEMP": float(match.group(7))
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
            print(f"Waiting... Arduino said: {line}")

def main():
    global system_start, cycle_count
    try:
        with serial.Serial(SERIAL_PORT, BAUDRATE, timeout=1) as ser, \
             open(CSV_FILENAME, mode='w', newline='', buffering=5) as csvfile:

            csv_writer = csv.writer(csvfile)
            csv_writer.writerow(["step", "timestamp", "encoder", "SIN_P", "COS_P", "SIN_N", "COS_N", "TEMP"])

            print(f"Starting system...\nLogging to: {CSV_FILENAME}")
            print(f"System runtime: {RUN_SYSTEM_MIN} min → Estimated cycles: {NUM_CYCLES}")

            time.sleep(2)  # Allow Arduino to reset
            ser.reset_input_buffer()
            ser.reset_output_buffer()

            while (time.time() - system_start) <= RUN_SYSTEM_MIN * 60:
                # Trigger a new cycle
                ser.write(b'2\n')
                wait_for_ack(ser, CYCLE_START_OK)

                cycle_count += 1
                print(f"\n{'*'*50}\nStarting cycle {cycle_count}\n{'*'*50}")

                while True:
                    line = ser.readline().decode('utf-8', errors='ignore').strip()
                    if not line:
                        continue

                    if CYCLE_ENDED_INDICATOR in line:
                        print(f"\n{'*'*50}\nFinished cycle {cycle_count}\n{'*'*50}")
                        break

                    parsed = parse_sensor_line(line)
                    if parsed:
                        timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                        # print(f"Step {parsed['step']:>2} | Time {timestamp} | "
                        #       f"Enc={parsed['encoder']:>5} | SIN_P={parsed['SIN_P']} | "
                        #       f"COS_P={parsed['COS_P']} | SIN_N={parsed['SIN_N']} | "
                        #       f"COS_N={parsed['COS_N']} | Temp={parsed['TEMP']:.2f}°C")

                        csv_writer.writerow([
                            parsed['step'], timestamp, parsed['encoder'],
                            parsed['SIN_P'], parsed['COS_P'],
                            parsed['SIN_N'], parsed['COS_N'],
                            parsed['TEMP']
                        ])
                        csvfile.flush()
                        print(time.time() - system_start)
                    print(f"OUTSIDE: {time.time() - system_start}")

            print(f"\nSystem completed. Total cycles: {cycle_count}")
            print(f"Data logged to: {CSV_FILENAME}")

    except KeyboardInterrupt:
        print(f"\nInterrupted by user. Data saved to {CSV_FILENAME}")
    except serial.SerialException as e:
        print(f"Serial connection error: {e}")

if __name__ == "__main__":
    main()
