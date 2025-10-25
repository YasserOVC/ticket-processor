🎫 Ticket Processing and Conditional Reporting
==============================================

A Python script designed to process raw CSV ticket data, filter for essential fields, sort by scheduled start time, and generate a color-coded Excel report for easy prioritization.

✨ Features
----------

*   **Data Filtering:** Selects only the essential fields: **Ticket No**, **Serial** (from Machine), **Site Name**, **Assignee**, and **Start Time**.
    
*   **Time-Based Sorting:** Sorts all records by the Start Time column to organize the workload chronologically.
    
*   **Conditional Coloring:** Applies color-coding to the Start Time column in the output Excel file for quick visual triage:
    
    *   🟢 **Green:** Tickets scheduled for **Today**.
        
    *   🟡 **Yellow:** Tickets scheduled for **Yesterday**.
        
    *   🔵 **Blue:** Tickets scheduled for **Tomorrow** and the **Day After Tomorrow**.
        
    *   🔴 **Red:** All other dates (past or far future).
        

🛠️ Setup and Installation
--------------------------

This project requires Python and a few common data science libraries. It is strongly recommended to use a **virtual environment** to manage dependencies.

### Prerequisites

You must have Python (3.7+) installed on your system.

### Step 1: Clone the Repository

Bash

Plain textANTLR4BashCC#CSSCoffeeScriptCMakeDartDjangoDockerEJSErlangGitGoGraphQLGroovyHTMLJavaJavaScriptJSONJSXKotlinLaTeXLessLuaMakefileMarkdownMATLABMarkupObjective-CPerlPHPPowerShell.propertiesProtocol BuffersPythonRRubySass (Sass)Sass (Scss)SchemeSQLShellSwiftSVGTSXTypeScriptWebAssemblyYAMLXML`   git clone [YOUR_REPOSITORY_URL]  cd [YOUR_REPOSITORY_NAME]   `

### Step 2: Create and Activate the Virtual Environment

Bash

Plain textANTLR4BashCC#CSSCoffeeScriptCMakeDartDjangoDockerEJSErlangGitGoGraphQLGroovyHTMLJavaJavaScriptJSONJSXKotlinLaTeXLessLuaMakefileMarkdownMATLABMarkupObjective-CPerlPHPPowerShell.propertiesProtocol BuffersPythonRRubySass (Sass)Sass (Scss)SchemeSQLShellSwiftSVGTSXTypeScriptWebAssemblyYAMLXML`   # Create the virtual environment  python3 -m venv venv  # Activate the environment (Linux/macOS/WSL)  source venv/bin/activate  # Activate the environment (Windows PowerShell)  # .\venv\Scripts\Activate.ps1   `

### Step 3: Install Dependencies

With the virtual environment active, install the required libraries:

Bash

Plain textANTLR4BashCC#CSSCoffeeScriptCMakeDartDjangoDockerEJSErlangGitGoGraphQLGroovyHTMLJavaJavaScriptJSONJSXKotlinLaTeXLessLuaMakefileMarkdownMATLABMarkupObjective-CPerlPHPPowerShell.propertiesProtocol BuffersPythonRRubySass (Sass)Sass (Scss)SchemeSQLShellSwiftSVGTSXTypeScriptWebAssemblyYAMLXML`   pip install pandas openpyxl XlsxWriter   `

🚀 Usage
--------

### Step 1: Prepare the Data

Place your input CSV file (e.g., tickets-2025-10-25.csv) into the root directory of the project.

**Note:** The script is hardcoded to use tickets-2025-10-25.csv as input to match the date-based coloring logic. You can modify the INPUT\_CSV variable at the bottom of ticket\_processor.py if your file name changes.

### Step 2: Run the Script

Execute the Python script while your virtual environment is active:

Bash

Plain textANTLR4BashCC#CSSCoffeeScriptCMakeDartDjangoDockerEJSErlangGitGoGraphQLGroovyHTMLJavaJavaScriptJSONJSXKotlinLaTeXLessLuaMakefileMarkdownMATLABMarkupObjective-CPerlPHPPowerShell.propertiesProtocol BuffersPythonRRubySass (Sass)Sass (Scss)SchemeSQLShellSwiftSVGTSXTypeScriptWebAssemblyYAMLXML`   python ticket_processor.py   `

### Output

The script will generate an Excel file named **filtered\_and\_colored\_tickets.xlsx** in the same directory. This file is ready for immediate review and prioritization based on the color-coded Start Time column.

📂 File Structure
-----------------

Plain textANTLR4BashCC#CSSCoffeeScriptCMakeDartDjangoDockerEJSErlangGitGoGraphQLGroovyHTMLJavaJavaScriptJSONJSXKotlinLaTeXLessLuaMakefileMarkdownMATLABMarkupObjective-CPerlPHPPowerShell.propertiesProtocol BuffersPythonRRubySass (Sass)Sass (Scss)SchemeSQLShellSwiftSVGTSXTypeScriptWebAssemblyYAMLXML`   .  ├── tickets-2025-10-25.csv  # Your input data  ├── ticket_processor.py     # The main Python script  ├── venv/                   # Python Virtual Environment folder  └── README.md               # This file   `