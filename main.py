import serial
import time
import re
import csv
from datetime import datetime

# Define port and baudrate
SERIAL_PORT = 'COM14'
BAUDRATE = 115200

# Regex pattern to extract sensor data from USB Serial communication
SENSOR_DATA_REGEX = r"\s*(\d+)\s+(-?\d+)\s+(\d+)\s+(\d+)\s+(\d+)\s+(\d+)\s+([0-9.]+)"

# Generate timestamped filename
timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
csv_filename = f"measurement_log_{timestamp}.csv"

def parse_line(line):
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

def main():
    try:
        with serial.Serial(SERIAL_PORT, BAUDRATE, timeout=1) as ser, \
             open(csv_filename, mode='w', newline='') as csvfile:

            csv_writer = csv.writer(csvfile)
            csv_writer.writerow(["step", "encoder", "SIN_P", "COS_P", "SIN_N", "COS_N", "TEMP"]) # Column header

            time.sleep(2)  # Give Arduino time to reset

            # Flushing out the serial communication
            ser.reset_input_buffer()
            ser.reset_output_buffer()

            print(f"Telling Arduino to start measurement...\nLogging to: {csv_filename}")
            ser.write(b'2\n')

            while True:
                line = ser.readline().decode('utf-8', errors='ignore').strip()
                if not line:
                    continue

                parsed = parse_line(line)
                if parsed:
                    print(f"Step {parsed['step']:>2}: "
                          f"Encoder={parsed['encoder']:>5} | "
                          f"SIN_P={parsed['SIN_P']} | COS_P={parsed['COS_P']} | "
                          f"SIN_N={parsed['SIN_N']} | COS_N={parsed['COS_N']} | "
                          f"Temp={parsed['TEMP']:.2f}°C")

                    # Save data to CSV
                    csv_writer.writerow([
                        parsed['step'],
                        parsed['encoder'],
                        parsed['SIN_P'],
                        parsed['COS_P'],
                        parsed['SIN_N'],
                        parsed['COS_N'],
                        parsed['TEMP']
                    ])

    except KeyboardInterrupt:
        print(f"\nLogging stopped. Data saved to {csv_filename}")
    except serial.SerialException as e:
        print(f"Serial error: {e}")

if __name__ == "__main__":
    main()
