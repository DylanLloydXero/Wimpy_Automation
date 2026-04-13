import os

# Base Directory
BASE_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))

# Data Paths
DATA_DIR = os.path.join(BASE_DIR, "data")
INPUT_DIR = os.path.join(DATA_DIR, "input")
ARCHIVE_DIR = os.path.join(DATA_DIR, "archives")
LOG_DIR = os.path.join(BASE_DIR, "logs")

# Specific File Paths
STAFF_MASTER_FILE = os.path.join(DATA_DIR, "templates", "Staff_Details_Template.xlsx")
LATEST_ROSTER_FILE = os.path.join(INPUT_DIR, "latest_roster.xlsx")
OFF_DAY_REQUESTS_FILE = os.path.join(DATA_DIR, "off_day_requests.json")
STAFF_PINS_FILE = os.path.join(DATA_DIR, "staff_pins.json")
AI_LEARNING_LOG = os.path.join(DATA_DIR, "roster_learning_log.json")

# Ensure directories exist
for d in [DATA_DIR, INPUT_DIR, ARCHIVE_DIR, LOG_DIR, os.path.join(DATA_DIR, "templates")]:
    os.makedirs(d, exist_ok=True)

# Settings
PORT = 8000
HOST = "0.0.0.0"
DEBUG = True
