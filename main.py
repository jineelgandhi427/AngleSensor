import serial
import time
import re
import csv
from datetime import datetime

# Define port and baudrate
SERIAL_PORT = 'COM8'
BAUDRATE = 115200

# Set cycle settings here----------------------------------------------------------------------------------------------------------------
RUN_SYSTEM_MIN = 2 # Set the time in mins for how many minutes you want take the readings. !IT CAN ONLY BE GREATER THAN 1 AND IN INTEGER!
#----------------------------------------------------------------------------------------------------------------------------------------

CYCLE_DURATION_MIN = 1.5  # Approximate duration of one cycle.
NUM_OF_CYCLES_TO_COMPLETE = 1 if int(RUN_SYSTEM_MIN/CYCLE_DURATION_MIN) <= 1 else int(RUN_SYSTEM_MIN/CYCLE_DURATION_MIN)

# Regex pattern defines
SENSOR_DATA_REGEX = r"\s*(\d+)\s+(-?\d+)\s+(\d+)\s+(\d+)\s+(\d+)\s+(\d+)\s+([0-9.]+)" # Regex pattern to extract sensor data from USB Serial communication.
CYCLE_ENDED_INDICATOR = "Measurements ended for cycle:"
CYCLE_START_OK = "You have selected option 2 the main program"

# Generate timestamped filename
timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
csv_filename = f"measurement_log_{timestamp}.csv"

main_program_cycle_count = 1 # To store the number of cycles completed.
can_start_new_cycle = True # Flag about when to start the new cycle.
can_close_the_system = False # Flag to inidicate that cycle is completed and ensures the system doesnot stop in between when the time is up.
cycle_start_time = None # To store the time when cycles started.
system_start_time = None # To store the time when system was started.
system_stop_time = None # To store the time when system ended compeletly.

def parse_data(line):
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

def check_if_cycle_ended(line):
    print(f"the msg is: {line}")
    if CYCLE_ENDED_INDICATOR in line:
        return True
    return False

def main():
    global main_program_cycle_count, can_close_the_system, can_start_new_cycle, cycle_start_time, system_start_time, system_stop_time
    try:
        with serial.Serial(SERIAL_PORT, BAUDRATE, timeout=1) as ser, \
             open(csv_filename, mode='w', newline='', buffering=1) as csvfile:

            csv_writer = csv.writer(csvfile)
            csv_writer.writerow(["step", "timestamp", "encoder", "SIN_P", "COS_P", "SIN_N", "COS_N", "TEMP"]) # Column header

            time.sleep(2)  # Give Arduino time to reset

            # Flushing out the serial communication
            ser.reset_input_buffer()
            ser.reset_output_buffer()

            system_start_time = time.time() # here the start time is defined just to avoid errors, it actually doesnot use this.

            print(f"Starting system...\nLogging to: {csv_filename}")
            print(f"The System will run for -> {RUN_SYSTEM_MIN} minutes")
            print(f"The System will execute approximate {NUM_OF_CYCLES_TO_COMPLETE} cycles")

            notedown_system_start_time = True

            while (time.time() - system_start_time <= RUN_SYSTEM_MIN*60) and can_start_new_cycle: # Run the system till mentioned by user. Convert the min to sec -> RUN_SYSTEM_MIN*60
                if can_start_new_cycle:
                    line = ''
                    ser.write(b'2\n') # Sending signal to arduino to start a new cycle.
                    # wait till the arduino has started recevied the cmd and started the cycle.
                    while CYCLE_START_OK not in line:
                        line = ser.readline().decode('utf-8', errors='ignore').strip()
                    print(f"*"*50)
                    print(f"Starting cycle number: {main_program_cycle_count}")
                    print(f"*"*50)
                    can_start_new_cycle = False
                    if notedown_system_start_time: # Note down the time when the ardunio starts the cycles, but only do it once.
                        system_start_time = time.time()
                        notedown_system_start_time = False

                line = ser.readline().decode('utf-8', errors='ignore').strip()
                if not line:
                    continue
                
                timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
                is_cycle_ended = check_if_cycle_ended(line)
                print(is_cycle_ended)
                if not is_cycle_ended:
                    parsed = parse_data(line)
                    if parsed:
                        print(f"Step {parsed['step']:>2}: "
                            f"Timestamp {timestamp}"
                            f"Encoder={parsed['encoder']:>5} | "
                            f"SIN_P={parsed['SIN_P']} | COS_P={parsed['COS_P']} | "
                            f"SIN_N={parsed['SIN_N']} | COS_N={parsed['COS_N']} | "
                            f"Temp={parsed['TEMP']:.2f}°C")

                        # Save data to CSV
                        csv_writer.writerow([
                            parsed['step'],
                            timestamp,
                            parsed['encoder'],
                            parsed['SIN_P'],
                            parsed['COS_P'],
                            parsed['SIN_N'],
                            parsed['COS_N'],
                            parsed['TEMP']
                        ])

                        csvfile.flush() # Forces an immediate write to disk
                else:
                    print(f"*"*50)
                    print(f"Finished cycle number: {main_program_cycle_count}")
                    print(f"*"*50)
                    main_program_cycle_count = main_program_cycle_count + 1
                    can_start_new_cycle = True

    except KeyboardInterrupt:
        print(f"\nLogging stopped. Data saved to {csv_filename}")
    except serial.SerialException as e:
        print(f"Serial error: {e}")

if __name__ == "__main__":
    main()
