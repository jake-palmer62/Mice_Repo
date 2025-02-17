import pandas as pd
import numpy as np
from datetime import datetime, timedelta
import os

def get_mouse_info():
    """Get mouse info from Excel"""
    xl = pd.ExcelFile('Experiment m Data Summary.xlsx')
    wt_df = pd.read_excel(xl, '(WT) Food intake', header=2)
    hom_df = pd.read_excel(xl, '(HOM) Food intake', header=2)
    
    mouse_info = {}
    for df in [wt_df, hom_df]:
        for _, row in df.iterrows():
            mouse_id = str(row['Animal #']).strip()
            if mouse_id.startswith('m'):
                date = row['Date']
                if isinstance(date, str):
                    date = datetime.strptime(date, '%d/%m/%Y')
                elif isinstance(date, datetime):
                    pass
                else:
                    try:
                        date = datetime.fromordinal(datetime(1900, 1, 1).toordinal() + int(date) - 2)
                    except:
                        continue
                
                mouse_info[mouse_id] = {
                    'start_date': date,
                    'strain': row['Strain'],
                    'treatment': row['Treatment'],
                    'cage': None,  # Will be filled in when processing feed files
                    'meals': {
                        'D1': {'18:00-19:00': [], '23:30-00:30': [], '05:00-06:00': []},
                        'D7': {'18:00-19:00': [], '23:30-00:30': [], '05:00-06:00': []},
                        'D14': {'18:00-19:00': [], '23:30-00:30': [], '05:00-06:00': []},
                        'D20': {'18:00-19:00': [], '23:30-00:30': [], '05:00-06:00': []}
                    }
                }
    return mouse_info

def process_feed_file(file_path, mouse_info):
    """Process a single feed file and extract meal data"""
    print(f"\nProcessing file: {file_path}")
    
    with open(file_path, 'r', encoding='utf-8', errors='ignore') as f:
        lines = f.readlines()
    
    # Get metadata and cage mappings
    experiment_start = None
    cage_to_mouse = {}
    current_cage = None
    
    for line in lines:
        line = line.strip()
        if 'EXPERIMENT START:' in line:
            experiment_start = datetime.strptime(line.split(':', 1)[1].strip(), '%d/%m/%Y %H:%M:%S')
            print(f"Experiment start: {experiment_start}")
        elif 'GROUP/CAGE:' in line:
            current_cage = line.split(':')[1].strip()
        elif 'SUBJECT ID:' in line and current_cage:
            mouse_id = line.split(':')[1].strip().replace('Mouse ', '').strip()
            if mouse_id and mouse_id != 'Empty':
                cage_to_mouse[current_cage] = mouse_id
                # Store cage number for this mouse
                if mouse_id in mouse_info:
                    mouse_info[mouse_id]['cage'] = current_cage
                print(f"Found mouse {mouse_id} in cage {current_cage}")
    
    # Map cage columns
    data_start = False
    column_mapping = {}
    
    for line in lines:
        if ':DATA' in line:
            data_start = True
            continue
        if data_start and 'INTERVAL' in line:
            parts = line.strip().split(',')
            for i, part in enumerate(parts):
                for cage in cage_to_mouse:
                    if f"CAGE {cage}" in part:
                        column_mapping[cage] = {
                            'time_col': i-1,
                            'value_col': i
                        }
            break
    
    # Process data
    for line in lines:
        if data_start and not 'INTERVAL' in line and not '===' in line and line.strip():
            parts = [p.strip() for p in line.split(',')]
            if len(parts) < 9:
                continue
            
            for cage, mouse_id in cage_to_mouse.items():
                if cage in column_mapping:
                    cols = column_mapping[cage]
                    try:
                        timestamp = datetime.strptime(parts[cols['time_col']], '%d/%m/%Y %H:%M:%S')
                        value = float(parts[cols['value_col']].replace('E-', 'e-'))
                        
                        if value <= 0.02:  # Skip small/negative values
                            continue
                            
                        # Get time window
                        hour = timestamp.hour
                        minute = timestamp.minute
                        window = None
                        
                        if hour == 18:
                            window = '18:00-19:00'
                        elif (hour == 23 and minute >= 30) or (hour == 0 and minute <= 30):
                            window = '23:30-00:30'
                        elif hour == 5:
                            window = '05:00-06:00'
                            
                        if window and mouse_id in mouse_info:
                            # Determine which experimental day
                            days_diff = (timestamp.date() - mouse_info[mouse_id]['start_date'].date()).days + 1
                            
                            if days_diff in [1, 7, 14, 20]:
                                day_key = f'D{days_diff}'
                                mouse_info[mouse_id]['meals'][day_key][window].append(value)
                                
                    except Exception as e:
                        continue
    
    return mouse_info

def create_summary_sheets(mouse_info):
    """Create summary DataFrames for each day"""
    day_data = {day: [] for day in ['D1', 'D7', 'D14', 'D20']}
    
    for mouse_id, info in mouse_info.items():
        strain_prefix = 'WT' if info['strain'] == 'C57BL/6' else 'GHSR'
        treatment = f"{strain_prefix}-{info['treatment']}"
        
        for day in ['D1', 'D7', 'D14', 'D20']:
            windows = info['meals'][day]
            day_data[day].append({
                'Animal #': mouse_id[1:],  # Remove 'm' prefix
                'Cage': info['cage'],
                'Treatment': treatment,
                'Date': info['start_date'].strftime('%d/%m/%y'),
                '18:00-19:00': round(sum(windows['18:00-19:00']), 2),
                '23:30-00:30': round(sum(windows['23:30-00:30']), 2),
                '05:00-06:00': round(sum(windows['05:00-06:00']), 2),
                '18:00-19:00_Total': sum(windows['18:00-19:00']),
                '23:30-00:30_Total': sum(windows['23:30-00:30']),
                '05:00-06:00_Total': sum(windows['05:00-06:00'])
            })
    
    # Convert to DataFrames
    summary_dfs = {}
    for day, data in day_data.items():
        if data:  # Only create sheet if there's data
            df = pd.DataFrame(data)
            
            # Add means and SEMs
            means = df.groupby('Treatment')[['18:00-19:00', '23:30-00:30', '05:00-06:00', 
                                           '18:00-19:00_Total', '23:30-00:30_Total', '05:00-06:00_Total']].mean().round(2)
            sems = df.groupby('Treatment')[['18:00-19:00', '23:30-00:30', '05:00-06:00',
                                          '18:00-19:00_Total', '23:30-00:30_Total', '05:00-06:00_Total']].sem().round(2)
            
            summary_dfs[day] = {
                'data': df,
                'means': means,
                'sems': sems
            }
    
    return summary_dfs

def main():
    # Get mouse info
    mouse_info = get_mouse_info()
    
    # Get all feed files and sort them properly
    feed_files = [f for f in os.listdir('.') if f.endswith('.CSV') and 'FEED1' in f]
    
    # Custom sort function to handle the file naming pattern
    def get_sort_key(filename):
        # Extract run number and subversion
        parts = filename.split('Run')
        if len(parts) > 1:
            run_part = parts[1].split('.')[0]  # Get the part between 'Run' and first '.'
            # Handle cases like '5' vs '5.0'
            if '.' in run_part:
                run_num, sub_num = run_part.split('.')
            else:
                run_num, sub_num = run_part, '0'
            return (int(run_num), float(sub_num))
        return (0, 0)
    
    # Sort files
    feed_files = sorted(feed_files, key=get_sort_key)
    print("\nProcessing files in order:")
    for file in feed_files:
        print(f"  {file}")
        mouse_info = process_feed_file(file, mouse_info)
    
    # Print raw data for verification
    print("\nMeal summaries:")
    for mouse_id, info in mouse_info.items():
        has_data = False
        data_str = [f"\n{mouse_id} (Cage {info['cage']}):" if info['cage'] else f"\n{mouse_id}:"]
        
        for day in ['D1', 'D7', 'D14', 'D20']:
            meals = info['meals'][day]
            if any(len(m) > 0 for m in meals.values()):
                has_data = True
                data_str.append(f"\n{day}:")
                for window, values in meals.items():
                    if values:
                        total = sum(values)
                        if total > 0:
                            data_str.append(f"  {window}: {total:.2f}g ({len(values)} events)")
        
        if has_data or info['cage']:  # Show mice even if they have no data
            print('\n'.join(data_str))
    
    # Create summary sheets
    summary_dfs = create_summary_sheets(mouse_info)
    
    # Save to Excel
    with pd.ExcelWriter('meal_analysis.xlsx') as writer:
        for day, dfs in summary_dfs.items():
            # Write main data
            start_row = 0
            dfs['data'].to_excel(writer, sheet_name=day, startrow=start_row, index=False)
            
            # Write means
            start_row = len(dfs['data']) + 2
            dfs['means'].to_excel(writer, sheet_name=day, startrow=start_row)
            
            # Write SEMs
            start_row = start_row + len(dfs['means']) + 2
            dfs['sems'].to_excel(writer, sheet_name=day, startrow=start_row)
    
    print("\nAnalysis complete. Results saved to meal_analysis.xlsx")

if __name__ == "__main__":
    main()