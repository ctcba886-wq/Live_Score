import pandas as pd
import time
import shutil
import os
import json

# --- Internal Script Configuration ---
# This section contains private settings, like file paths and polling intervals,
# which are not meant to be changed by the end-user.
SOURCE_EXCEL_PATH = 'source.xlsx'
TEMP_EXCEL_PATH = 'temp_scores.xlsx'
SCORE_SHEET_NAME = 'Scores' # The name of the sheet in the Excel file to read
POLLING_INTERVAL_SECONDS = 300 # 5 minutes
CONFIG_FILE_PATH = 'config.json'
OUTPUT_SCORES_PATH = 'scores.json'

def load_config():
    """
    Loads the event configuration from config.json.
    This file is intended to be edited by tournament staff.
    """
    try:
        with open(CONFIG_FILE_PATH, 'r', encoding='utf-8') as f:
            return json.load(f)
    except FileNotFoundError:
        print(f"FATAL ERROR: Configuration file '{CONFIG_FILE_PATH}' not found.")
        exit()
    except json.JSONDecodeError:
        print(f"FATAL ERROR: Could not parse '{CONFIG_FILE_PATH}'. Please ensure it is a valid JSON.")
        exit()

def process_scoresheet_data():
    """
    Safely copies the Excel file, reads the detailed scoresheet data based on
    a predefined structure, and writes it to the scores.json file.
    """
    # --- 1. Safe File Copy ---
    try:
        shutil.copy(SOURCE_EXCEL_PATH, TEMP_EXCEL_PATH)
    except (IOError, PermissionError) as e:
        print(f"[{time.ctime()}] Could not copy file '{SOURCE_EXCEL_PATH}'. It may be locked. Skipping. Error: {e}")
        return
    except FileNotFoundError:
        print(f"[{time.ctime()}] Source file not found at: '{SOURCE_EXCEL_PATH}'. Please check the path.")
        return

    # --- 2. Read and Process Data from Copied File ---
    try:
        # Note: The data extraction logic below is a placeholder. It assumes a very specific,
        # non-standard Excel layout and will likely need to be adjusted to match the
        # actual scoresheet format. It reads data from fixed cell positions.
        df = pd.read_excel(TEMP_EXCEL_PATH, sheet_name=SCORE_SHEET_NAME, header=None)
        
        boards_data = []
        for i in range(1, 17): # Assuming boards 1-16
            # This logic IS A PLACEHOLDER and needs to be adapted for the real .xlsx file.
            board_data = {
                "boardNumber": i,
                "openRoom": {
                    "contract": df.iloc[i + 2, 2], "scoreNS": df.iloc[i + 2, 3], "scoreEW": df.iloc[i + 2, 4]
                },
                "closedRoom": {
                    "contract": df.iloc[i + 2, 5], "scoreNS": df.iloc[i + 2, 6], "scoreEW": df.iloc[i + 2, 7]
                },
                "diff": { "ns": df.iloc[i + 2, 8], "ew": df.iloc[i + 2, 9] },
                "imp": { "ns": df.iloc[i + 2, 10], "ew": df.iloc[i + 2, 11] }
            }
            boards_data.append(board_data)

        # Placeholder for extracting totals from the sheet
        totals = {
            "impNS": df.iloc[20, 9], "impEW": df.iloc[20, 10],
            "resultVP": df.iloc[21, 9], "totalIMPFinal": df.iloc[22, 9],
            "finalResult": df.iloc[20, 11]
        }

        output_data = {"boards": boards_data, "totals": totals}

        # --- 3. Generate JSON Output ---
        with open(OUTPUT_SCORES_PATH, 'w', encoding='utf-8') as f:
            # Custom JSON encoder to handle pandas/numpy data types that
            # might result from reading the Excel file.
            class NpEncoder(json.JSONEncoder):
                def default(self, obj):
                    if pd.isna(obj): return None
                    if isinstance(obj, (pd.Int64Dtype, pd.IntegerDtype)): return int(obj)
                    if isinstance(obj, (pd.Float64Dtype, pd.Float32Dtype)): return float(obj)
                    if isinstance(obj, (bool, pd.BooleanDtype)): return bool(obj)
                    return super(NpEncoder, self).default(obj)
            
            json.dump(output_data, f, ensure_ascii=False, indent=4, cls=NpEncoder)

        print(f"[{time.ctime()}] Successfully updated {OUTPUT_SCORES_PATH}.")

    except Exception as e:
        print(f"[{time.ctime()}] An error occurred while processing the Excel file: {e}")
    finally:
        if os.path.exists(TEMP_EXCEL_PATH):
            os.remove(TEMP_EXCEL_PATH)

if __name__ == "__main__":
    # Load the user-editable config file to ensure it's valid on startup.
    # The 'config' variable itself is not directly used in the processing loop,
    # as the frontend fetches it, but loading it here acts as a sanity check.
    config = load_config()
    
    print("Starting the live score updater.")
    print(f"Event Name: {config.get('eventName', 'N/A')}")
    print(f"Watching for changes in: {SOURCE_EXCEL_PATH}")
    print(f"Polling every {POLLING_INTERVAL_SECONDS} seconds. Press Ctrl+C to stop.")
    
    # Start the polling loop to update scores.
    # It runs once immediately, then repeats on the defined interval.
    while True:
        try:
            process_scoresheet_data()
            time.sleep(POLLING_INTERVAL_SECONDS)
        except KeyboardInterrupt:
            print("\nStopping the score updater. Goodbye!")
            break
        except Exception as e:
            print(f"An unexpected error occurred in the main loop: {e}")
            time.sleep(POLLING_INTERVAL_SECONDS)