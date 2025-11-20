import pandas as pd
import time
import shutil
import os
import json

# --- Internal Script Configuration ---
# This section contains private settings, like file paths and polling intervals,
# which are not meant to be changed by the end-user.
TEMP_EXCEL_PATH = 'temp_scores.xlsx'
POLLING_INTERVAL_SECONDS = 300 # 5 minutes
CONFIG_FILE_PATH = 'config.json'
OUTPUT_SCORES_PATH = 'scores.json'

# Load sensitive information (like a GitHub token) from environment variables
GITHUB_TOKEN = os.getenv('GITHUB_TOKEN')

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

def process_scoresheet_data(boards_to_play, source_excel_path_arg, score_sheet_name_arg):
    """
    Safely copies the Excel file, reads the detailed scoresheet data based on
    a predefined structure, and writes it to the scores.json file.
    Returns True if the match is complete, False otherwise.
    """
    is_complete = False
    # --- 1. Safe File Copy ---
    try:
        shutil.copy(source_excel_path_arg, TEMP_EXCEL_PATH)
    except (IOError, PermissionError) as e:
        print(f"[{time.ctime()}] Could not copy file '{source_excel_path_arg}'. It may be locked. Skipping. Error: {e}")
        return False
    except FileNotFoundError:
        print(f"[{time.ctime()}] FATAL ERROR: Source Excel file not found at: '{source_excel_path_arg}'. Please ensure the file exists and the path is correct.")
        raise SystemExit(1) # Terminate the script immediately if the source Excel file is not found

    # --- 2. Read and Process Data from Copied File ---
    try:
        df = pd.read_excel(TEMP_EXCEL_PATH, sheet_name=score_sheet_name_arg, header=None)
        
        boards_data = []
        for i in range(1, boards_to_play + 1):
            # Row index is based on the user's specification:
            # Board 1 starts at row 10 (index 9), and each subsequent board is 3 rows down.
            row_index = 9 + (i - 1) * 3
            
            board_data = {
                "boardNumber": int(df.iloc[row_index, 0]),  # Col A
                "openRoom": { 
                    "contract": df.iloc[row_index, 6],  # Col G
                    "scoreNS": df.iloc[row_index, 12], # Col M
                    "scoreEW": df.iloc[row_index, 15]  # Col P
                },
                "closedRoom": { 
                    "contract": df.iloc[row_index, 18], # Col S
                    "scoreNS": df.iloc[row_index, 24], # Col Y
                    "scoreEW": df.iloc[row_index, 27]  # Col AB
                },
                "diff": { 
                    "ns": df.iloc[row_index, 30], # Col AE
                    "ew": df.iloc[row_index, 33]  # Col AH
                },
                "imp": { 
                    "ns": df.iloc[row_index, 36], # Col AK
                    "ew": df.iloc[row_index, 39]  # Col AN
                }
            }
            # Convert pandas NA/NaN to Python None for JSON serialization
            for top_key in ["openRoom", "closedRoom", "diff", "imp"]:
                for sub_key, value in board_data[top_key].items():
                    if pd.isna(value):
                        board_data[top_key][sub_key] = None
            
            boards_data.append(board_data)

        # Totals locations are unknown, so set to None.
        totals = {
            "impNS": None, "impEW": None,
            "resultVP": None, "totalIMPFinal": None,
            "finalResult": None
        }

        output_data = {"boards": boards_data, "totals": totals}

        # --- 3. Check for Match Completion ---
        # Condition: Open and Closed room contract cells for the last board are not empty.
        last_board_row = 9 + (boards_to_play - 1) * 3
        last_board_open_contract = df.iloc[last_board_row, 6]   # Col G
        last_board_closed_contract = df.iloc[last_board_row, 18] # Col S

        if (not pd.isna(last_board_open_contract) and str(last_board_open_contract).strip() != '' and
            not pd.isna(last_board_closed_contract) and str(last_board_closed_contract).strip() != ''):
            is_complete = True

        # --- 4. Generate JSON Output ---
        with open(OUTPUT_SCORES_PATH, 'w', encoding='utf-8') as f:
            json.dump(output_data, f, ensure_ascii=False, indent=4)

        print(f"[{time.ctime()}] Successfully updated {OUTPUT_SCORES_PATH}.")
        return is_complete

    except Exception as e:
        print(f"[{time.ctime()}] An error occurred while processing the Excel file: {e}")
        return False
    finally:
        if os.path.exists(TEMP_EXCEL_PATH):
            os.remove(TEMP_EXCEL_PATH)

def generate_blank_scores_file():
    """Creates a blank scores.json file to clear old data on startup."""
    blank_data = {
        "boards": [],
        "totals": {
            "impNS": None, "impEW": None,
            "resultVP": None, "totalIMPFinal": None,
            "finalResult": None
        }
    }
    try:
        with open(OUTPUT_SCORES_PATH, 'w', encoding='utf-8') as f:
            json.dump(blank_data, f, ensure_ascii=False, indent=4)
        print(f"[{time.ctime()}] Generated a blank {OUTPUT_SCORES_PATH} to clear old scores.")
    except Exception as e:
        print(f"[{time.ctime()}] WARNING: Could not generate a blank scores file. Error: {e}")

if __name__ == "__main__":
    config = load_config()
    boards_per_round = config.get('boards_per_round', 16) # Default to 16 if not found
    event_name = config.get('eventName', 'DefaultEvent')
    current_round = config.get('round', '1')

    # Dynamically construct source Excel path and sheet name
    SOURCE_EXCEL_PATH_DYNAMIC = f"{event_name}.xlsx"
    SCORE_SHEET_NAME_DYNAMIC = f"R{current_round}"
    
    print("Starting the live score updater.")
    if GITHUB_TOKEN:
        print("Successfully loaded GITHUB_TOKEN from environment variables.")
    else:
        print("Warning: GITHUB_TOKEN environment variable not found. Proceeding without it.")
    
    print(f"Event Name: {event_name}")
    print(f"Current Round: {current_round}")
    print(f"Boards per round: {boards_per_round}")
    print(f"Watching Excel file: {SOURCE_EXCEL_PATH_DYNAMIC}, sheet: {SCORE_SHEET_NAME_DYNAMIC}")
    print(f"Polling every {POLLING_INTERVAL_SECONDS} seconds. Press Ctrl+C to stop.")
    print("This script will automatically terminate after 2 hours.")

    # Generate a blank scores file on startup to clear old data
    generate_blank_scores_file()

    start_time = time.time()
    timeout_seconds = 2 * 60 * 60
    
    while True:
        try:
            # Timeout check
            if time.time() - start_time > timeout_seconds:
                print(f"[{time.ctime()}] Script has been running for 2 hours. Terminating automatically.")
                break

            match_is_complete = process_scoresheet_data(boards_per_round, SOURCE_EXCEL_PATH_DYNAMIC, SCORE_SHEET_NAME_DYNAMIC)
            if match_is_complete:
                print(f"[{time.ctime()}] All scores entered for {boards_per_round} boards. Final update complete. Stopping.")
                break
            
            time.sleep(POLLING_INTERVAL_SECONDS)
        except KeyboardInterrupt:
            print("\nStopping the score updater. Goodbye!")
            break
        except Exception as e:
            print(f"An unexpected error occurred in the main loop: {e}")
            time.sleep(POLLING_INTERVAL_SECONDS)