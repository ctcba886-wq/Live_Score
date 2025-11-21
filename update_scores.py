import pandas as pd
import time
import shutil
import os
import json
import base64
import re

try:
    import requests
except ImportError:
    print("FATAL ERROR: The 'requests' library is not installed. Please run: pip install requests")
    exit()

# --- Internal Script Configuration ---
TEMP_EXCEL_PATH = 'temp_scores.xls'
POLLING_INTERVAL_SECONDS = 300  # 5 minutes
CONFIG_FILE_PATH = 'config.json'
OUTPUT_SCORES_PATH = 'scores.json'
REPO_INFO_FILE = 'githubtokenofctcba.txt'  # File containing the repository URL

# ==============================================================================
# GITHUB UPLOAD FUNCTIONS
# ==============================================================================

def get_repo_slug_from_file():
    """Reads the repository slug from the configuration file."""
    try:
        with open(REPO_INFO_FILE, 'r') as f:
            content = f.read()
            repo_match = re.search(r'repo:(.*)', content)

            if not repo_match:
                print(f"[{time.ctime()}] WARNING: Could not find 'repo:' field in {REPO_INFO_FILE}. GitHub upload will be disabled.")
                return None

            repo_url = repo_match.group(1).strip()
            repo_slug_match = re.search(r'github\.com/([^/]+/[^/]+)', repo_url)
            if not repo_slug_match:
                print(f"[{time.ctime()}] WARNING: Could not parse repository owner/name from URL: {repo_url}. GitHub upload will be disabled.")
                return None
            
            repo_slug = repo_slug_match.group(1).replace('.git', '')
            return repo_slug

    except FileNotFoundError:
        print(f"[{time.ctime()}] WARNING: Repo info file '{REPO_INFO_FILE}' not found. GitHub upload will be disabled.")
        return None

def upload_file_to_github(file_path, token, repo_slug, commit_message):
    """Reads a file and uploads it to the specified GitHub repository."""
    if not os.path.exists(file_path):
        print(f"[{time.ctime()}] UPLOAD ERROR: File to upload '{file_path}' not found.")
        return
        
    with open(file_path, 'r', encoding='utf-8') as f:
        file_content = f.read()

    encoded_content = base64.b64encode(file_content.encode('utf-8')).decode('utf-8')
    file_name = os.path.basename(file_path)
    api_url = f"https://api.github.com/repos/{repo_slug}/contents/{file_name}"
    
    headers = {
        "Authorization": f"token {token}",
        "Accept": "application/vnd.github.v3+json"
    }

    sha = None
    try:
        response = requests.get(api_url, headers=headers)
        if response.status_code == 200:
            sha = response.json().get('sha')
        elif response.status_code != 404:
            print(f"[{time.ctime()}] UPLOAD WARNING: Could not get file SHA. Status: {response.status_code}")
    except requests.exceptions.RequestException as e:
        print(f"[{time.ctime()}] UPLOAD WARNING: Network error when checking for file: {e}")

    data = {
        "message": commit_message,
        "content": encoded_content,
        "committer": {"name": "Bridge Score Updater", "email": "bot@example.com"}
    }
    if sha:
        data['sha'] = sha

    try:
        response = requests.put(api_url, headers=headers, data=json.dumps(data))
        response.raise_for_status()
        commit_sha = response.json().get('commit', {}).get('sha', 'N/A')
        print(f"[{time.ctime()}] Successfully uploaded '{file_name}'. Commit: {commit_sha[:7]}")
    except requests.exceptions.HTTPError as e:
        print(f"[{time.ctime()}] UPLOAD ERROR: Failed to upload '{file_name}'. Status: {e.response.status_code}, Response: {e.response.text}")
    except requests.exceptions.RequestException as e:
        print(f"[{time.ctime()}] UPLOAD ERROR: Network error during upload: {e}")

# ==============================================================================
# SCORE PROCESSING FUNCTIONS
# ==============================================================================

def load_config():
    """Loads the event configuration from config.json."""
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
    Safely copies and processes the Excel scoresheet.
    Returns a tuple: (was_updated, is_complete)
    """
    is_complete = False
    try:
        shutil.copy(source_excel_path_arg, TEMP_EXCEL_PATH)
    except (IOError, PermissionError) as e:
        print(f"[{time.ctime()}] Could not copy file '{source_excel_path_arg}'. It may be locked. Skipping. Error: {e}")
        return False, False
    except FileNotFoundError:
        print(f"[{time.ctime()}] FATAL ERROR: Source Excel file not found: '{source_excel_path_arg}'.")
        exit()

    try:
        df = pd.read_excel(TEMP_EXCEL_PATH, sheet_name=score_sheet_name_arg, header=None)
        
        boards_data = []
        for i in range(1, boards_to_play + 1):
            row_index = 9 + (i - 1) * 3
            board_data = {
                "boardNumber": int(df.iloc[row_index, 0]), "openRoom": {"contract": df.iloc[row_index, 6], "scoreNS": df.iloc[row_index, 12], "scoreEW": df.iloc[row_index, 15]},
                "closedRoom": {"contract": df.iloc[row_index, 18], "scoreNS": df.iloc[row_index, 24], "scoreEW": df.iloc[row_index, 27]},
                "diff": {"ns": df.iloc[row_index, 30], "ew": df.iloc[row_index, 33]},"imp": {"ns": df.iloc[row_index, 36], "ew": df.iloc[row_index, 39]}
            }
            for top_key in board_data:
                if isinstance(board_data[top_key], dict):
                    for sub_key, value in board_data[top_key].items():
                        if pd.isna(value): board_data[top_key][sub_key] = None
            boards_data.append(board_data)

        totals = {"impNS": None, "impEW": None, "resultVP": None, "totalIMPFinal": None, "finalResult": None}
        output_data = {"boards": boards_data, "totals": totals}

        last_board_row = 9 + (boards_to_play - 1) * 3
        last_open = df.iloc[last_board_row, 6]
        last_closed = df.iloc[last_board_row, 18]
        if pd.notna(last_open) and str(last_open).strip() and pd.notna(last_closed) and str(last_closed).strip():
            is_complete = True

        with open(OUTPUT_SCORES_PATH, 'w', encoding='utf-8') as f:
            json.dump(output_data, f, ensure_ascii=False, indent=4)
        print(f"[{time.ctime()}] Successfully generated {OUTPUT_SCORES_PATH}.")
        return True, is_complete
    except Exception as e:
        print(f"[{time.ctime()}] An error occurred while processing the Excel file: {e}")
        return False, False
    finally:
        if os.path.exists(TEMP_EXCEL_PATH):
            os.remove(TEMP_EXCEL_PATH)

def generate_blank_scores_file():
    """Creates a blank scores.json file to clear old data."""
    blank_data = {"boards": [], "totals": {"impNS": None, "impEW": None, "resultVP": None, "totalIMPFinal": None, "finalResult": None}}
    with open(OUTPUT_SCORES_PATH, 'w', encoding='utf-8') as f:
        json.dump(blank_data, f, ensure_ascii=False, indent=4)
    print(f"[{time.ctime()}] Generated blank {OUTPUT_SCORES_PATH}.")

# ==============================================================================
# MAIN EXECUTION
# ==============================================================================

if __name__ == "__main__":
    config = load_config()
    boards_per_round = config.get('boards_per_round', 16)
    event_name = config.get('eventName', 'DefaultEvent')
    current_round = config.get('round', '1')

    SOURCE_EXCEL_PATH_DYNAMIC = f"{event_name}.xls"
    SCORE_SHEET_NAME_DYNAMIC = f"R{current_round}"

    GITHUB_TOKEN = os.getenv('GITHUB_TOKEN')
    REPO_SLUG = get_repo_slug_from_file()
    
    print("--- Bridge Live Score Updater ---")
    if GITHUB_TOKEN and REPO_SLUG:
        print(f"GitHub uploads enabled for repository: {REPO_SLUG}")
    else:
        print("GitHub upload disabled: GITHUB_TOKEN environment variable or repo config might be missing.")
    
    print(f"Event: {event_name}, Round: {current_round}, Boards: {boards_per_round}")
    print(f"Watching Excel: {SOURCE_EXCEL_PATH_DYNAMIC}, Sheet: {SCORE_SHEET_NAME_DYNAMIC}")
    print(f"Polling every {POLLING_INTERVAL_SECONDS} seconds. Press Ctrl+C to stop.")
    print("---------------------------------")

    # --- Initial Uploads ---
    if GITHUB_TOKEN and REPO_SLUG:
        print(f"[{time.ctime()}] Performing initial file uploads...")
        upload_file_to_github(
            CONFIG_FILE_PATH, 
            GITHUB_TOKEN, 
            REPO_SLUG, 
            f"Upload event configuration for {event_name}"
        )
        generate_blank_scores_file()
        upload_file_to_github(
            OUTPUT_SCORES_PATH, 
            GITHUB_TOKEN, 
            REPO_SLUG, 
            f"Initial empty scores for Round {current_round}"
        )
    else:
        generate_blank_scores_file()

    start_time = time.time()
    timeout_seconds = 2 * 60 * 60
    
    while True:
        try:
            if time.time() - start_time > timeout_seconds:
                print(f"[{time.ctime()}] Script timed out after 2 hours. Terminating.")
                break

            updated, match_is_complete = process_scoresheet_data(boards_per_round, SOURCE_EXCEL_PATH_DYNAMIC, SCORE_SHEET_NAME_DYNAMIC)
            
            if updated and GITHUB_TOKEN and REPO_SLUG:
                upload_file_to_github(OUTPUT_SCORES_PATH, GITHUB_TOKEN, REPO_SLUG, f"Update scores for Round {current_round}")

            if match_is_complete:
                print(f"[{time.ctime()}] Final scores processed. Stopping.")
                break
            
            time.sleep(POLLING_INTERVAL_SECONDS)
        except KeyboardInterrupt:
            print("\nStopping the score updater. Goodbye!")
            break
        except Exception as e:
            print(f"An unexpected error occurred in the main loop: {e}")
            time.sleep(POLLING_INTERVAL_SECONDS)
