import pandas as pd
import os
from datetime import timedelta
import matplotlib.pyplot as plt
from matplotlib.backends.backend_pdf import PdfPages
import numpy as np
from matplotlib.gridspec import GridSpec
import math
import textwrap
import re
from datetime import datetime

# Configure matplotlib to use a font that supports Chinese characters
import matplotlib
matplotlib.rcParams['font.sans-serif'] = ['SimSun', 'Arial Unicode MS', 'DejaVu Sans', 'Microsoft YaHei']
matplotlib.rcParams['font.family'] = 'sans-serif'
matplotlib.rcParams['axes.unicode_minus'] = False

# Ignore warnings
import warnings
warnings.filterwarnings('ignore')

def get_week_number(date):
    """Get ISO week number from date"""
    if pd.isna(date):
        return None
    return date.isocalendar()[1]

def get_week_start_end(date):
    """Get start and end date of the week for a given date"""
    if pd.isna(date):
        return None, None
    start = date - timedelta(days=date.weekday())
    end = start + timedelta(days=6)
    return start, end

def get_week_description(start_date, end_date):
    """Generate description for weeks that cross months"""
    if start_date is None or end_date is None:
        return "Unknown Week"
    
    if start_date.month != end_date.month:
        return f"Week {get_week_number(start_date)} ({start_date.strftime('%b %d')} - {end_date.strftime('%b %d')})"
    else:
        return f"Week {get_week_number(start_date)} ({start_date.strftime('%b %d')} - {end_date.strftime('%d')})"

def get_simplified_date_range(start_date, end_date):
    """Generate simplified date range without week number"""
    if start_date is None or end_date is None:
        return "Unknown"
    
    if start_date.month != end_date.month:
        return f"{start_date.strftime('%b %d')} - {end_date.strftime('%b %d')}"
    else:
        return f"{start_date.strftime('%b %d')} - {end_date.strftime('%d')}"

def get_month_from_week_description(week_description):
    """Extract month from week description"""
    if "Unknown Week" in week_description:
        return "Unknown"
    
    # Extract month abbreviation
    try:
        month_abbr = week_description.split('(')[1].split(' ')[0]
        return month_abbr
    except:
        return "Unknown"

def get_week_sort_key(week_info):
    """Extract a sortable key from week description for chronological sorting"""
    # First try to use the sample_date if available
    if 'sample_date' in week_info and pd.notna(week_info['sample_date']):
        return (week_info['sample_date'].month, week_info['sample_date'].isocalendar()[1])
    
    # Extract week number
    week_match = re.search(r'Week (\d+)', week_info['week'])
    if not week_match:
        return (9999, 0)  # Default high value for unknown weeks
    
    week_num = int(week_match.group(1))
    
    # Extract month
    month = week_info['month']
    month_order = {
        'Jan': 1, 'Feb': 2, 'Mar': 3, 'Apr': 4, 'May': 5, 'Jun': 6,
        'Jul': 7, 'Aug': 8, 'Sep': 9, 'Oct': 10, 'Nov': 11, 'Dec': 12
    }
    
    month_num = month_order.get(month, 0)
    
    # For weeks that span months, use the first month
    return (month_num, week_num)

def extract_week_number(week_description):
    """Extract just the week number from the week description"""
    match = re.search(r'Week (\d+)', week_description)
    if match:
        return int(match.group(1))
    return 0

def format_num(num):
    """Return integer if whole, otherwise the float."""
    if num == int(num):
        return int(num)
    return num

def format_container_count(rac_count, mbt_count):
    """Format container counts, omitting zero values and handling floats."""
    if rac_count == 0 and mbt_count == 0:
        return ""
    
    parts = []
    if rac_count > 0:
        parts.append(f"RAC {format_num(rac_count)}")
    if mbt_count > 0:
        parts.append(f"MBT {format_num(mbt_count)}")
    
    return " + ".join(parts)

def clean_customer_name(name):
    """Clean customer names by removing special characters"""
    if not isinstance(name, str):
        return name
    
    # Replace Chinese character '台' with ' units'
    if '台' in name:
        name = name.replace('台', ' units')
    
    # Remove any other special characters
    clean_name = ''.join(c for c in name if ord(c) >= 32 and (ord(c) < 127 or ord(c) > 159) and ord(c) != 0x200B and ord(c) != 0xFEFF)
    
    return clean_name

def normalize_contr_id(contr_str):
    """Create a canonical ID for containers, especially combined ones."""
    if not isinstance(contr_str, str):
        return ""
    
    if '拼' in contr_str:
        # Split by the character, sort parts, and join with a standard separator
        parts = sorted([p.strip() for p in contr_str.split('拼')])
        return '/'.join(parts)
    else:
        # For regular containers, just return the cleaned string
        return contr_str.strip()

def calculate_container_counts(group):
    """Custom aggregation function to count containers, handling combined ones."""
    # Count each unique canonical container ID only once
    unique_containers = group.drop_duplicates(subset=['CanonicalContrID'])
    
    rac_count = 0.0
    mbt_count = 0.0
    
    for _, row in unique_containers.iterrows():
        # Check the original 'Contr #' for the combined character
        contr_str = row['Contr #']
        if isinstance(contr_str, str) and '拼' in contr_str:
            rac_count += 0.5
            mbt_count += 0.5
        elif row['ContainerType'] == 'RAC':
            rac_count += 1
        elif row['ContainerType'] == 'MBT':
            mbt_count += 1
            
    return pd.Series({'RAC': rac_count, 'MBT': mbt_count})

def clean_container_number(container_str):
    """Clean container numbers by removing newlines and extra spaces"""
    if not isinstance(container_str, str):
        return ""
    # Remove newlines and extra spaces
    cleaned = re.sub(r'\s+', '', container_str)
    return cleaned

def extract_suburb_from_address(address):
    """Extract suburb from warehouse address"""
    if not isinstance(address, str):
        return "Mel"
    
    if 'Pending' in address:
        return "Mel"
    
    if 'VIC' not in address:
        return "Mel"
    
    # Try to extract suburb
    suburb_match = re.search(r'(?:,\s*)([A-Za-z\s]+)(?:\s+VIC)', address)
    if suburb_match:
        suburb = suburb_match.group(1).strip()
        return suburb
    
    # Try another pattern
    suburb_match = re.search(r'([A-Za-z\s]+)(?:\s+VIC)', address)
    if suburb_match:
        suburb = suburb_match.group(1).strip()
        return suburb
    
    # Try to extract from the beginning of the address
    parts = address.split(',')
    if len(parts) > 1 and 'VIC' in parts[-1]:
        suburb = parts[-2].strip()
        return suburb
    
    return "Mel"  # Default

def map_suburb_to_main_suburb(suburb):
    """Map various suburb names to the main suburbs (Mulgrave, Dandenong, Laverton)"""
    suburb = suburb.lower()
    
    if 'mulgrave' in suburb:
        return 'Mulgrave'
    elif 'dandenong' in suburb:
        return 'Dandenong'
    elif 'laverton' in suburb or 'gilbertson' in suburb:
        return 'Laverton'
    elif 'ravenhall' in suburb or 'freight road' in suburb:
        return 'Laverton'  # Grouping Ravenhall with Laverton
    elif 'truganina' in suburb or 'carmen' in suburb:
        return 'Laverton'  # Grouping Truganina with Laverton
    elif 'cranbourne' in suburb:
        return 'Dandenong'  # Grouping Cranbourne with Dandenong
    elif 'mount waverley' in suburb:
        return 'Mulgrave'  # Grouping Mount Waverley with Mulgrave
    
    return 'Mel'  # Default

def create_paginated_table(data, columns, title, page_title_suffix="", max_rows_per_page=25, merge_group_col=None):
    """Create a paginated table with proper spacing and optional cell merging."""
    # Calculate number of pages needed
    total_rows = len(data)
    total_pages = math.ceil(total_rows / max_rows_per_page)
    
    pages = []
    
    for page_num in range(total_pages):
        # Get data for this page
        start_idx = page_num * max_rows_per_page
        end_idx = min((page_num + 1) * max_rows_per_page, total_rows)
        page_data = data[start_idx:end_idx]
        
        # Create figure with proper margins
        fig = plt.figure(figsize=(11, 8.5))
        
        # Add title with page number if multiple pages
        if total_pages > 1:
            if page_num == 0:
                plt.suptitle(f"{title}", fontsize=16, y=0.98, fontweight='bold')
            else:
                plt.suptitle(f"{title} (Continued{page_title_suffix})", fontsize=16, y=0.98, fontweight='bold')
            
            # Add page number at bottom
            plt.figtext(0.5, 0.02, f"Page {page_num + 1} of {total_pages}", ha='center')
        else:
            plt.suptitle(f"{title}", fontsize=16, y=0.98, fontweight='bold')
        
        # Create table with proper margins
        ax = plt.subplot(111)
        
        # Create the table
        table = ax.table(
            cellText=page_data,
            colLabels=columns,
            cellLoc='center',
            loc='center'
        )
        
        # Style the table
        table.auto_set_font_size(False)
        table.set_fontsize(7) # Reduced font size to fit more columns
        table.scale(1, 1.5)
        table.auto_set_column_width(col=list(range(len(columns))))
        
        # --- New Merging Logic ---
        if merge_group_col is not None and len(page_data) > 1:
            # Find groups of identical values in the merge column
            groups = []
            # Don't merge the final 'Total' row
            data_to_scan = page_data
            if page_data[-1][0] == 'Total':
                data_to_scan = page_data[:-1]

            if len(data_to_scan) > 1:
                current_group_start = 0
                for i in range(1, len(data_to_scan)):
                    val_current = data_to_scan[i][merge_group_col]
                    val_start = data_to_scan[current_group_start][merge_group_col]

                    if val_current != val_start or val_current == '':
                        if i - current_group_start > 1:
                            groups.append((current_group_start, i - 1))
                        current_group_start = i
                
                if len(data_to_scan) - current_group_start > 1:
                    groups.append((current_group_start, len(data_to_scan) - 1))

            # Apply merging effect
            for start, end in groups:
                # Hide text in subsequent cells
                for i in range(start + 1, end + 1):
                    table[i + 1, merge_group_col].get_text().set_text('')

                # Top cell of merge group
                table[start + 1, merge_group_col].visible_edges = 'LTR'
                
                # Middle cells
                for i in range(start + 1, end):
                    table[i + 1, merge_group_col].visible_edges = 'LR'

                # Bottom cell of merge group
                table[end + 1, merge_group_col].visible_edges = 'LBR'


        # Style the header row
        for j in range(len(columns)):
            cell = table[(0, j)]
            cell.set_facecolor('#D3D3D3')
            cell.set_text_props(weight='bold')
        
        # Add alternating row colors for better readability
        for i in range(len(page_data)):
            for j in range(len(columns)):
                cell = table[(i+1, j)]
                if i % 2 == 0:
                    cell.set_facecolor('#F0F0F0')
        
        ax.axis('off')
        
        # Add proper spacing
        plt.tight_layout(rect=[0.05, 0.05, 0.95, 0.95])
        
        pages.append(fig)
    
    return pages

def main():
    try:
        # Get the current directory
        current_dir = os.path.dirname(os.path.abspath(__file__)) or '.'
        
        # Create dynamic filename and title
        now = datetime.now()
        timestamp_str = now.strftime("%Y-%m-%d %H:%M")
        report_title = f"Container Report {timestamp_str}"
        output_filename = f"Container Report {now.strftime('%Y-%m-%d %H-%M-%S')}.pdf"
        
        # Define input and output paths
        input_file = os.path.join(current_dir, '出货汇总表8.25.xlsx')
        output_file = os.path.join(current_dir, 'Container_Summary.pdf')
        log_file = os.path.join(current_dir, 'container_summary_log.txt')
        
        with open(log_file, 'w', encoding='utf-8') as log:
            log.write(f"Current directory: {current_dir}\n")
            log.write(f"Input file: {input_file}\n")
            log.write(f"Output file: {output_file}\n")
            
            # Read the Excel file
            log.write("Reading Excel file...\n")
            rac_df = pd.read_excel(input_file, sheet_name='2025 RAC')
            mbt_df = pd.read_excel(input_file, sheet_name='2025 MBT')
            customs_df = pd.read_excel(input_file, sheet_name='Customs Decl. & Local Delivery')
            
            # Forward-fill 'Contr #' to handle merged cells
            log.write("Forward-filling 'Contr #' to handle merged cells...\n")
            rac_df['Contr #'] = rac_df['Contr #'].ffill()
            mbt_df['Contr #'] = mbt_df['Contr #'].ffill()

            log.write(f"RAC sheet: {len(rac_df)} rows\n")
            log.write(f"MBT sheet: {len(mbt_df)} rows\n")
            log.write(f"Customs sheet: {len(customs_df)} rows\n")
            
            # Ensure ETA column is in datetime format
            rac_df['ETA'] = pd.to_datetime(rac_df['ETA'], errors='coerce')
            mbt_df['ETA'] = pd.to_datetime(mbt_df['ETA'], errors='coerce')
            
            # Filter out rows with NaT in ETA
            rac_df = rac_df.dropna(subset=['ETA'])
            mbt_df = mbt_df.dropna(subset=['ETA'])
            
            # Add container type column
            rac_df['ContainerType'] = 'RAC'
            mbt_df['ContainerType'] = 'MBT'
            
            # Clean customer names
            rac_df['Customer'] = rac_df['Customer'].apply(clean_customer_name)
            mbt_df['Customer'] = mbt_df['Customer'].apply(clean_customer_name)
            
            # Create a canonical container ID to handle combined containers
            rac_df['CanonicalContrID'] = rac_df['Contr #'].apply(normalize_contr_id)
            mbt_df['CanonicalContrID'] = mbt_df['Contr #'].apply(normalize_contr_id)
            
            # Clean container numbers
            rac_df['CleanContr'] = rac_df['Contr #'].apply(clean_container_number)
            mbt_df['CleanContr'] = mbt_df['Contr #'].apply(clean_container_number)
            
            # Extract container numbers (part after slash)
            rac_df['ContainerNum'] = rac_df['CleanContr'].apply(lambda x: x.split('/', 1)[1] if isinstance(x, str) and '/' in x else x)
            mbt_df['ContainerNum'] = mbt_df['CleanContr'].apply(lambda x: x.split('/', 1)[1] if isinstance(x, str) and '/' in x else x)
            
            # Create container to suburb mapping
            container_to_suburb = {}
            
            # Process customs sheet to extract suburbs
            for _, row in customs_df.iterrows():
                if pd.notna(row['Container #']) and pd.notna(row['Warehouse address']):
                    container_num = str(row['Container #']).strip()
                    address = str(row['Warehouse address']).strip()
                    
                    # Extract suburb from address
                    suburb = extract_suburb_from_address(address)
                    
                    # Map to main suburb
                    main_suburb = map_suburb_to_main_suburb(suburb)
                    
                    # Store in mapping
                    container_to_suburb[container_num] = main_suburb
            
            # Update destinations in RAC and MBT sheets
            def update_destination(row):
                if row['Destination'] == 'VIC-Mel':
                    container_num = row['ContainerNum']
                    if container_num in container_to_suburb and container_to_suburb[container_num] != 'Mel':
                        return f"VIC-{container_to_suburb[container_num]}"
                    else:
                        return 'VIC-Mel (No Specific)'
                else:
                    return row['Destination']
            
            rac_df['OriginalDestination'] = rac_df['Destination']
            mbt_df['OriginalDestination'] = mbt_df['Destination']
            
            rac_df['Destination'] = rac_df.apply(update_destination, axis=1)
            mbt_df['Destination'] = mbt_df.apply(update_destination, axis=1)
            
            # Log destination changes
            log.write("\nDestination changes in RAC sheet:\n")
            destination_changes = rac_df[rac_df['Destination'] != rac_df['OriginalDestination']]
            log.write(f"Changed {len(destination_changes)} destinations\n")
            for i, (_, row) in enumerate(destination_changes.head(10).iterrows()):
                log.write(f"  {i+1}. '{row['CleanContr']}': '{row['OriginalDestination']}' → '{row['Destination']}'\n")
            
            log.write("\nDestination changes in MBT sheet:\n")
            destination_changes = mbt_df[mbt_df['Destination'] != mbt_df['OriginalDestination']]
            log.write(f"Changed {len(destination_changes)} destinations\n")
            for i, (_, row) in enumerate(destination_changes.head(10).iterrows()):
                log.write(f"  {i+1}. '{row['CleanContr']}': '{row['OriginalDestination']}' → '{row['Destination']}'\n")
            
            # Combine dataframes
            combined_df = pd.concat([rac_df, mbt_df])
            
            # Add week number and month
            combined_df['Week'] = combined_df['ETA'].apply(get_week_number)
            combined_df['Month'] = combined_df['ETA'].dt.month
            combined_df['MonthNum'] = combined_df['ETA'].dt.month # Added MonthNum
            combined_df['MonthName'] = combined_df['ETA'].dt.strftime('%B')
            combined_df['DayOfWeek'] = combined_df['ETA'].dt.strftime('%A')
            
            # --- START OF NEW SUMMARY TABLE LOGIC ---

            # 1. "Our Containers" Summary
            our_containers_summary = combined_df[combined_df['Customer'] == 'No'].copy()
            
            def get_location_hierarchy(dest):
                if 'VIC' in dest:
                    main = 'VIC'
                    if 'No Specific' in dest:
                        sub = 'Mel (not specific)'
                    elif '-' in dest:
                        sub = dest.split('-', 1)[1]
                    else:
                        sub = 'Mel'
                    return main, sub
                if '-' in dest:
                    parts = dest.split('-', 1)
                    return parts[0], parts[1]
                return dest, dest

            our_containers_summary[['MainLocation', 'SubLocation']] = our_containers_summary['Destination'].apply(lambda x: pd.Series(get_location_hierarchy(x)))
            
            our_summary_agg = our_containers_summary.groupby(['MainLocation', 'SubLocation', 'MonthName', 'MonthNum']).apply(calculate_container_counts).reset_index()

            # 2. "Customer Containers" Summary
            customer_containers_summary = combined_df[combined_df['Customer'] != 'No'].copy()
            
            # Extract state-level location
            def get_state_from_destination(dest):
                if isinstance(dest, str):
                    if '-' in dest:
                        return dest.split('-', 1)[0]
                    return dest # Should not happen, but as a fallback
                return 'Unknown'
            
            customer_containers_summary['State'] = customer_containers_summary['Destination'].apply(get_state_from_destination)

            # Aggregate by State, Customer, and Month
            customer_summary_agg = customer_containers_summary.groupby(
                ['State', 'Customer', 'MonthName', 'MonthNum']
            ).apply(calculate_container_counts).reset_index()


            # --- END OF NEW SUMMARY TABLE LOGIC ---

            # Calculate week start and end dates
            week_dates = combined_df['ETA'].apply(lambda x: get_week_start_end(x) if not pd.isna(x) else (None, None))
            combined_df['WeekStart'] = [date[0] for date in week_dates]
            combined_df['WeekEnd'] = [date[1] for date in week_dates]
            
            # Add week description
            combined_df['WeekDescription'] = combined_df.apply(
                lambda x: get_week_description(x['WeekStart'], x['WeekEnd']), axis=1
            )
            
            # Add simplified date range
            combined_df['SimplifiedDateRange'] = combined_df.apply(
                lambda x: get_simplified_date_range(x['WeekStart'], x['WeekEnd']), axis=1
            )
            
            log.write(f"Processed {len(combined_df)} total records\n")
            
            # Check if required columns exist
            required_columns = ['Contr #', 'Customer', 'Destination', 'ETA']
            missing_columns = [col for col in required_columns if col not in combined_df.columns]
            if missing_columns:
                log.write(f"ERROR: Missing required columns: {missing_columns}\n")
                print(f"ERROR: Missing required columns: {missing_columns}")
                return
            
            # Filter for containers with no customer (shipping to us)
            our_containers = combined_df[combined_df['Customer'] == 'No']
            log.write(f"Found {len(our_containers)} containers with Customer='No'\n")
            
            # Group by week for our containers (Customer='No')
            if not our_containers.empty:
                # First group by week to get unique weeks
                weeks_df = our_containers.groupby('WeekDescription').size().reset_index(name='Count')
                unique_weeks = weeks_df['WeekDescription'].tolist()
                
                # Create a list to store weekly container data
                weekly_container_data = []
                
                # For each week, process the container data
                for week in unique_weeks:
                    week_df = our_containers[our_containers['WeekDescription'] == week]
                    
                    # Get a sample date from this week for sorting
                    sample_date = week_df['ETA'].min()
                    
                    # Extract week number
                    week_num = extract_week_number(week)
                    
                    # Get simplified date range
                    simplified_date_range = week_df['SimplifiedDateRange'].iloc[0]
                    
                    # Group by destination and container type, count unique container numbers
                    destination_summary = week_df.groupby('Destination').apply(calculate_container_counts).reset_index()
                    
                    # Ensure both RAC and MBT columns exist
                    if 'RAC' not in destination_summary.columns:
                        destination_summary['RAC'] = 0
                    if 'MBT' not in destination_summary.columns:
                        destination_summary['MBT'] = 0
                    
                    # Calculate total
                    destination_summary['Total'] = destination_summary['RAC'] + destination_summary['MBT']
                    
                    # Create formatted counts (e.g., "RAC 3" or "MBT 4" or "RAC 3 + MBT 4")
                    destination_summary['FormattedCount'] = destination_summary.apply(
                        lambda x: format_container_count(x['RAC'], x['MBT']), axis=1
                    )
                    
                    # Add to weekly data
                    weekly_container_data.append({
                        'week': week,
                        'week_num': week_num,
                        'date_range': simplified_date_range,
                        'destination_summary': destination_summary,
                        'sample_date': sample_date
                    })
                
                # Sort weekly data chronologically
                weekly_container_data.sort(key=get_week_sort_key)
                
                log.write(f"Generated weekly container summary with {len(unique_weeks)} unique weeks\n")
            else:
                weekly_container_data = []
                log.write("No containers found shipping to us\n")
            
            # Filter for containers with customer (not shipping to us)
            customer_containers = combined_df[combined_df['Customer'] != 'No']
            log.write(f"Found {len(customer_containers)} containers with Customer!='No'\n")
            
            # Group by month for customer containers
            if not customer_containers.empty:
                # Create a list to store all monthly customer data for consolidated table
                all_monthly_data = []
                
                # Group by month
                for month_name, month_df in customer_containers.groupby('MonthName'):
                    month_num = pd.to_datetime(f"2025 {month_name} 1").month
                    
                    # Process each customer in this month
                    for customer, customer_df in month_df.groupby('Customer'):
                        # Group by destination and container type, count unique container numbers
                        destination_summary = customer_df.groupby('Destination').apply(calculate_container_counts).reset_index()
                        
                        # Ensure both RAC and MBT columns exist
                        if 'RAC' not in destination_summary.columns:
                            destination_summary['RAC'] = 0
                        if 'MBT' not in destination_summary.columns:
                            destination_summary['MBT'] = 0
                        
                        # Calculate total
                        destination_summary['Total'] = destination_summary['RAC'] + destination_summary['MBT']
                        
                        # Create formatted counts (e.g., "RAC 3" or "MBT 4" or "RAC 3 + MBT 4")
                        destination_summary['FormattedCount'] = destination_summary.apply(
                            lambda x: format_container_count(x['RAC'], x['MBT']), axis=1
                        )
                        
                        # Add each destination row to the consolidated data
                        for _, row in destination_summary.iterrows():
                            # Only include rows with at least one container
                            if row['Total'] > 0:
                                all_monthly_data.append({
                                    'Month': month_name,
                                    'MonthNum': month_num,
                                    'MonthName': month_name,
                                    'Customer': customer,
                                    'Destination': row['Destination'],
                                    'RAC': row['RAC'],
                                    'MBT': row['MBT'],
                                    'Total': row['Total'],
                                    'FormattedCount': row['FormattedCount']
                                })
                
                # Convert to DataFrame for easier sorting and processing
                monthly_df = pd.DataFrame(all_monthly_data)
                
                # Sort by month, then customer, then destination
                if not monthly_df.empty:
                    monthly_df = monthly_df.sort_values(['MonthNum', 'Customer', 'Destination'])
                
                log.write(f"Generated consolidated monthly customer summary with {len(monthly_df)} rows\n")
            else:
                monthly_df = pd.DataFrame()
                log.write("No containers found shipping to customers\n")
            
            # Create PDF with matplotlib
            log.write(f"Creating PDF report: {output_file}\n")
            
            with PdfPages(output_file) as pdf:
                # Create a figure for the title page
                plt.figure(figsize=(11, 8.5))
                plt.text(0.5, 0.5, report_title, 
                         fontsize=24, ha='center', va='center', fontweight='bold')
                plt.axis('off')
                pdf.savefig()
                plt.close()
                
                # --- RENDER NEW SUMMARY TABLES ---

                # Helper to build and render summary tables
                def create_summary_table_pages(agg_df, group_cols, title, pdf):
                    if agg_df.empty:
                        return
                    
                    month_order = sorted(agg_df[['MonthNum', 'MonthName']].drop_duplicates().values, key=lambda x: x[0])
                    months = [m[1] for m in month_order]
                    row_groups = sorted(agg_df[group_cols].drop_duplicates().values.tolist())

                    header = group_cols + months

                    table_data = []
                    for group_vals in row_groups:
                        row_data = list(group_vals)
                        
                        for month_num, month_name in month_order:
                            mask = (agg_df[group_cols[0]] == group_vals[0])
                            if len(group_cols) > 1:
                                mask &= (agg_df[group_cols[1]] == group_vals[1])
                            mask &= (agg_df['MonthName'] == month_name)
                            cell_data = agg_df[mask]
                            
                            if not cell_data.empty:
                                rac = cell_data['RAC'].iloc[0]
                                mbt = cell_data['MBT'].iloc[0]
                                row_data.append(format_container_count(rac, mbt))
                            else:
                                row_data.append('')
                        
                        table_data.append(row_data)

                    # Add the 'Total' row at the bottom
                    total_row = ['Total'] + [''] * (len(group_cols) - 1)
                    for month_num, month_name in month_order:
                        month_total_rac = agg_df[agg_df['MonthName'] == month_name]['RAC'].sum()
                        month_total_mbt = agg_df[agg_df['MonthName'] == month_name]['MBT'].sum()
                        total_row.append(format_container_count(month_total_rac, month_total_mbt))
                    
                    table_data.append(total_row)
                    
                    pages = create_paginated_table(table_data, header, title, max_rows_per_page=30, merge_group_col=0 if len(group_cols) > 1 else None)
                    for fig in pages:
                        pdf.savefig(fig)
                        plt.close(fig)

                create_summary_table_pages(our_summary_agg, ['MainLocation', 'SubLocation'], "Our Containers Summary", pdf)
                create_summary_table_pages(customer_summary_agg, ['State', 'Customer'], "Customer Containers Summary", pdf)


                # Create a consolidated table for all weekly data
                if weekly_container_data:
                    # First, get all unique destinations across all weeks
                    all_destinations = set()
                    for week_info in weekly_container_data:
                        dest_summary = week_info['destination_summary']
                        for dest in dest_summary['Destination'].tolist():
                            all_destinations.add(dest)
                    
                    all_destinations_list = sorted(list(all_destinations))
                    
                    # Move 'VIC-Mel' and 'VIC-Mel (No Specific)' to the end
                    vic_mel_no_specific = 'VIC-Mel (No Specific)'
                    
                    final_dest_list = []
                    vic_mel_items = []
                    
                    for d in all_destinations_list:
                        if d == 'VIC-Mel' or d == vic_mel_no_specific:
                            vic_mel_items.append(d)
                        else:
                            final_dest_list.append(d)
                    
                    # Add them back at the end, sorted
                    final_dest_list.extend(sorted(vic_mel_items))
                    
                    all_destinations = final_dest_list
                    log.write(f"Found {len(all_destinations)} unique destinations: {all_destinations}\n")
                    
                    # Create table data
                    table_data = []
                    
                    # Create header row with destinations
                    header = ['Week #', 'Date Range']
                    for dest in all_destinations:
                        header.append(dest)
                    header.append('Total')
                    
                    # Add data for each week
                    for week_info in weekly_container_data:
                        row = [week_info['week_num'], week_info['date_range']]
                        
                        # Get destination summary for this week
                        dest_summary = week_info['destination_summary']
                        
                        # Add data for each destination
                        for dest in all_destinations:
                            # Find this destination in the summary
                            dest_row = dest_summary[dest_summary['Destination'] == dest]
                            if not dest_row.empty:
                                row.append(dest_row['FormattedCount'].values[0])
                            else:
                                row.append("")
                        
                        # Calculate and append weekly total
                        weekly_total_rac = dest_summary['RAC'].sum()
                        weekly_total_mbt = dest_summary['MBT'].sum()
                        row.append(format_container_count(weekly_total_rac, weekly_total_mbt))
                        
                        table_data.append(row)
                    
                    # Create paginated weekly tables
                    weekly_pages = create_paginated_table(
                        table_data, 
                        header, 
                        "Weekly Container Summary",
                        max_rows_per_page=20
                    )
                    
                    # Add all pages to PDF
                    for fig in weekly_pages:
                        pdf.savefig(fig)
                        plt.close(fig)
                else:
                    # No weekly data available
                    plt.figure(figsize=(11, 8.5))
                    plt.text(0.5, 0.5, 'No container data available for containers shipping to us', 
                             fontsize=14, ha='center', va='center')
                    plt.axis('off')
                    pdf.savefig()
                    plt.close()
                
                # Create monthly customer container summary tables, one table per month
                if not monthly_df.empty:
                    # Get sorted unique months from the dataframe
                    unique_months = monthly_df.drop_duplicates(subset=['MonthNum'])[['MonthNum', 'MonthName']].sort_values('MonthNum').values

                    for month_num, month_name in unique_months:
                        month_data_df = monthly_df[monthly_df['MonthNum'] == month_num]
                        
                        # Convert DataFrame to list of lists for table, handling customer grouping
                        monthly_table_data = []
                        last_customer = None
                        for _, row in month_data_df.iterrows():
                            customer = row['Customer']
                            if customer == last_customer:
                                customer_display = ""
                            else:
                                customer_display = customer
                            last_customer = customer
                            
                            monthly_table_data.append([
                                customer_display,
                                row['Destination'],
                                row['FormattedCount'],
                                format_num(row['Total'])
                            ])
                        
                        # Create paginated monthly tables for the current month
                        monthly_pages = create_paginated_table(
                            monthly_table_data, 
                            ['Customer', 'Destination', 'Container Count', 'Total'], 
                            f"Monthly Container Summary for {month_name}",
                            f"{month_name} Customer Containers",
                            max_rows_per_page=25
                        )
                        
                        # Add all pages to PDF
                        for fig in monthly_pages:
                            pdf.savefig(fig)
                            plt.close(fig)
                else:
                    # No monthly data available
                    plt.figure(figsize=(11, 8.5))
                    plt.text(0.5, 0.5, 'No container data available for customer containers', 
                             fontsize=14, ha='center', va='center')
                    plt.axis('off')
                    pdf.savefig()
                    plt.close()
            
            log.write("PDF report created successfully\n")
            print(f"Improved Suburb-Specific Container Summary Report generated successfully: {output_file}")
            print(f"Log file created: {log_file}")
            
    except Exception as e:
        print(f"Error: {str(e)}")
        import traceback
        traceback.print_exc()
        
        # Try to write error to log file
        try:
            with open(log_file, 'a', encoding='utf-8') as log:
                log.write(f"\nError: {str(e)}\n")
                log.write(traceback.format_exc())
        except:
            pass

if __name__ == "__main__":
    main()
