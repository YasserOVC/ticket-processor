Ticket Processing and Conditional Reporting 🎫


A Python script designed to process raw CSV ticket data, filter for essential fields, sort by scheduled start time, and generate a color-coded Excel report for easy prioritization.
Features Data Filtering✨: Selects only the essential fields: Ticket No, Serial (from Machine), Site Name, Assignee, and Start Time.
Time-Based Sorting: Sorts all records by the Start Time column to organize the workload chronologically.
Conditional Coloring: Applies color-coding to the Start Time column in the output Excel file for quick visual triage:
🟢 Green: Tickets scheduled for Today.
🟡 Yellow: Tickets scheduled for Yesterday.
🔵 Blue: Tickets scheduled for Tomorrow and the Day After Tomorrow.
🔴 Red: All other dates (past or far future).
Setup and Installation🛠️
This project requires Python and a few common data science libraries. It is strongly recommended to use a virtual environment to manage dependencies.
Prerequisites
You must have Python (3.7+) installed on your system.
Step 1: Clone the Repository
git clone git@github.com:YasserOVC/ticket-processor.gitcd ticket-processor
Step 2: Create and Activate the Virtual Environment:
Create the virtual environment
python3 -m venv venv
Activate the environment (Linux/macOS/WSL)
source venv/bin/activate
Activate the environment (Windows PowerShell)
.\venv\Scripts\Activate.ps1
Step 3: Install Dependencie With the virtual environment active, install the required libraries:
pip install pandas openpyxl XlsxWriter
Usage🚀:
Step 1: Prepare the Data Place your input CSV file (e.g., tickets-2025-10-25.csv) into the root directory of the project.
Note: The coloring logic in ticket_processor.py is currently fixed to use 2025-10-25 as "Today" to ensure reproducible results based on your uploaded file.
Step 2: Run the Script
Execute the Python script while your virtual environment is active: python ticket_processor.py.
Output
The script will generate an Excel file named filtered_and_colored_tickets.xlsx in the same directory. This file is ready for immediate review and prioritization based on the color-coded Start Time column.
Project Structure📂
.
├── tickets-2025-10-25.csv  # Example input data
├── ticket_processor.py     # The main Python script
├── venv/                   # Python Virtual Environment folder
└── README.md               # This file
