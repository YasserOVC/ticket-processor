import pandas as pd
from datetime import date, timedelta
import sys

def process_tickets(input_file: str, output_file: str = 'filtered_and_colored_tickets.xlsx'):
    """
    Reads a CSV file, filters columns, sorts by Start Time, and creates
    an Excel file with conditional coloring based on the Start Time date.
    
    Color Codes (Based on file date 2025-10-25):
    - Green: Today (2025-10-25)
    - Yellow: Yesterday (2025-10-24)
    - Blue: Tomorrow (2025-10-26) and Day After Tomorrow (2025-10-27)
    - Red: All other dates

    :param input_file: The name of the input CSV file.
    :param output_file: The name of the output Excel file.
    """
    try:
        # --- Date Definitions (Fixed for reproducible coloring) ---
        # Fixed to 2025-10-25, which is the date associated with your uploaded file.
        TODAY_FIXED = pd.to_datetime(date(2025, 10, 25))
        YESTERDAY_FIXED = TODAY_FIXED - timedelta(days=1)
        TOMORROW_FIXED = TODAY_FIXED + timedelta(days=1)
        DAY_AFTER_TOMORROW_FIXED = TODAY_FIXED + timedelta(days=2)
        
        # 1. Load the CSV file
        df = pd.read_csv(input_file)
        print(f"Successfully loaded data from {input_file}.")

        # Define columns to keep and their final names
        column_map = {
            'Ticket': 'Ticket No',
            'Machine': 'Serial', 
            'Site Name': 'Site Name',
            'Assignee': 'Assignee',
            'Start Time': 'Start Time'
        }

        # 2. Select, rename, and copy the required columns
        df_filtered = df[list(column_map.keys())].copy()
        df_filtered.rename(columns=column_map, inplace=True)

        # 3. Convert 'Start Time' for sorting and comparison
        df_filtered['Start Date'] = pd.to_datetime(df_filtered['Start Time']).dt.normalize()

        # 4. Sort the DataFrame by 'Start Time'
        df_sorted = df_filtered.sort_values(by='Start Date').reset_index(drop=True)

        # Prepare final DataFrame by dropping the temporary date column
        df_output = df_sorted.drop(columns=['Start Date'])

        # --- Excel Writing and Conditional Formatting ---

        # Create a Pandas Excel writer using XlsxWriter as the engine.
        writer = pd.ExcelWriter(output_file, engine='xlsxwriter', datetime_format='yyyy-mm-dd hh:mm:ss')
        df_output.to_excel(writer, sheet_name='Tickets', index=False)

        # Get the xlsxwriter workbook and worksheet objects.
        workbook  = writer.book
        worksheet = writer.sheets['Tickets']

        # Define formats for the coloring
        green_format = workbook.add_format({'bg_color': '#C6EFCE', 'font_color': '#006100'})  # Today
        yellow_format = workbook.add_format({'bg_color': '#FFEB9C', 'font_color': '#9C6500'}) # Yesterday
        blue_format = workbook.add_format({'bg_color': '#ADD8E6', 'font_color': '#000080'})   # Tomorrow/Day After Tomorrow
        red_format = workbook.add_format({'bg_color': '#FFC7CE', 'font_color': '#9C0006'})    # Other

        # Get the column index for 'Start Time' 
        start_time_col_index = df_output.columns.get_loc('Start Time')

        # Apply formatting cell by cell
        for row_num, cell_value in enumerate(df_output['Start Time']):
            # Excel row index is 1-based (row 0 is the header)
            excel_row = row_num + 1

            # Normalize the cell value to date only for comparison
            cell_date = pd.to_datetime(cell_value).normalize()

            # Determine which color format to apply using the correct format variables
            if cell_date == TODAY_FIXED:
                format_to_apply = green_format
            elif cell_date == YESTERDAY_FIXED:
                format_to_apply = yellow_format
            elif cell_date == TOMORROW_FIXED or cell_date == DAY_AFTER_TOMORROW_FIXED:
                format_to_apply = blue_format
            else:
                format_to_apply = red_format

            # Write the cell value with the specific format
            worksheet.write_string(excel_row, start_time_col_index, cell_value, format_to_apply)

        # Close the Pandas Excel writer
        writer.close()
        print(f"Successfully created and formatted Excel file: {output_file}")

    except FileNotFoundError:
        print(f"Error: The input file '{input_file}' was not found.", file=sys.stderr)
    except KeyError as e:
        print(f"Error: One of the required columns was not found in the CSV: {e}", file=sys.stderr)
    except Exception as e:
        print(f"An unexpected error occurred: {e}", file=sys.stderr)

if __name__ == '__main__':
    # Ensure this file name matches your CSV upload
    INPUT_CSV = 'tickets-2025-10-25.csv' 
    OUTPUT_EXCEL = 'filtered_and_colored_tickets.xlsx'

    process_tickets(INPUT_CSV, OUTPUT_EXCEL)