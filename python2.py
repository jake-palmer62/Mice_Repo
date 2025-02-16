import pandas as pd
import numpy as np
from datetime import datetime, timedelta
import os

def is_same_experimental_day(time1, time2):
    """Check if two timestamps belong to the same experimental day"""
    # If second time is between midnight and 6am, consider it part of previous day
    if 0 <= time2.hour < 6:
        time2 = time2 - timedelta(days=1)
    return time1.date() == time2.date()

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
                    'meals': {
                        'D1': {'18:00-19:00': [], '23:30-00:30': [], '05:00-06:00': []},
                        'D7': {'18:00-19:00': [], '23:30-00:30': [], '05:00-06:00': []},
                        'D14': {'18:00-19:00': [], '23:30-00:30': [], '05:00-06:00': []},
                        'D20': {'18:00-19:00': [], '23:30-00:30': [], '05:00-06:00': []}
                    }
                }
    return mouse_info

def get_meal_window(timestamp):
    """Get which meal window a timestamp belongs to"""
    hour = timestamp.hour
    minute = timestamp.minute
    
    if hour == 18:
        return '18:00-19:00'
    elif (hour == 23 and minute >= 30) or (hour == 0 and minute <= 30):
        return '23:30-00:30'
    elif hour == 5:
        return '05:00-06:00'
    return None

def process_feed_file(file_path, mouse_info):
    """Process a single feed file and extract meal data"""
    print(f"\nProcessing file: {file_path}")
    
    with open(file_path, 'r', encoding='utf-8', errors='ignore') as f:
        lines = f.readlines()
    
    # Get metadata
    experiment_start = None
    cage_to_mouse = {}
    for line in lines:
        line = line.strip()
        if 'EXPERIMENT START:' in line:
            experiment_start = datetime.strptime(line.split(':', 1)[1].strip(), '%d/%m/%Y %H:%M:%S')
            print(f"File start: {experiment_start}")
        elif 'GROUP/CAGE:' in line:
            current_cage = line.split(':')[1].strip()
        elif 'SUBJECT ID:' in line and 'current_cage' in locals():
            mouse_id = line.split(':')[1].strip().replace('Mouse ', '').strip()
            if mouse_id in mouse_info:
                cage_to_mouse[current_cage] = mouse_id
                print(f"Found {mouse_id} in cage {current_cage}")
    
    # Find cage columns
    header = None
    cage_columns = {}
    data_start = False
    
    for line in lines:
        if ':DATA' in line:
            data_start = True
            continue
        if data_start and 'INTERVAL' in line:
            header = line.strip().split(',')
            for i, col in enumerate(header):
                for cage in cage_to_mouse:
                    if f"CAGE {cage}" in col:
                        cage_columns[cage] = {
                            'time_col': i-1,
                            'value_col': i
                        }
            break
    
    # Process data
    for mouse_id in cage_to_mouse.values():
        print(f"\nProcessing data for mouse {mouse_id}")
        mouse_data = mouse_info[mouse_id]
        start_date = mouse_data['start_date']
        
        # Calculate target dates
        target_dates = {
            'D1': start_date,
            'D7': start_date + timedelta(days=6),
            'D14': start_date + timedelta(days=13),
            'D20': start_date + timedelta(days=19)
        }
        
        for line in lines:
            if not data_start or '===' in line or 'INTERVAL' in line or not line:
                continue
                
            parts = [p.strip() for p in line.split(',')]
            if len(parts) < 9:
                continue
                
            # Find mouse's cage
            mouse_cage = None
            for cage, mouse in cage_to_mouse.items():
                if mouse == mouse_id:
                    mouse_cage = cage
                    break
                    
            if not mouse_cage:
                continue
                
            cols = cage_columns[mouse_cage]
            
            try:
                timestamp = datetime.strptime(parts[cols['time_col']], '%d/%m/%Y %H:%M:%S')
                value = float(parts[cols['value_col']].replace('E-', 'e-'))
                
                if value <= 0.02:  # Skip insignificant values
                    continue
                    
                # Check if this is a feeding window
                window = get_meal_window(timestamp)
                if not window:
                    continue
                    
                # Match to experimental day
                for day, target_date in target_dates.items():
                    if is_same_experimental_day(target_date, timestamp):
                        mouse_data['meals'][day][window].append(value)
                        print(f"Added {value:.2f}g to {day} {window}")
                        
            except Exception as e:
                continue
    
    return mouse_info

def create_summary_sheets(mouse_info):
    """Create summary DataFrames for each day"""
    day_data = {day: [] for day in ['D1', 'D7', 'D14', 'D20']}
    
    for mouse_id, info in mouse_info.items():
        if info['treatment'] == 'Meal-Fed':
            strain_prefix = 'WT' if info['strain'] == 'C57BL/6' else 'GHSR'
            treatment = f"{strain_prefix}-{info['treatment']}"
            
            for day in ['D1', 'D7', 'D14', 'D20']:
                meals = info['meals'][day]
                meal1 = round(sum(meals['18:00-19:00']), 2)
                meal2 = round(sum(meals['23:30-00:30']), 2)
                meal3 = round(sum(meals['05:00-06:00']), 2)
                
                if meal1 > 0 or meal2 > 0 or meal3 > 0:  # Only add if there's data
                    day_data[day].append({
                        'Animal #': mouse_id[1:],  # Remove 'm' prefix
                        'Treatment': treatment,
                        'Date': info['start_date'].strftime('%d/%m/%y'),
                        'Meal 1': meal1,
                        'Meal 2': meal2,
                        'Meal 3': meal3
                    })
    
    # Convert to DataFrames
    summary_dfs = {}
    for day, data in day_data.items():
        if data:  # Only create sheet if there's data
            df = pd.DataFrame(data)
            
            # Calculate means and SEMs by treatment
            means = df.groupby('Treatment')[['Meal 1', 'Meal 2', 'Meal 3']].mean().round(2)
            sems = df.groupby('Treatment')[['Meal 1', 'Meal 2', 'Meal 3']].sem().round(2)
            
            summary_dfs[day] = {
                'data': df,
                'means': means,
                'sems': sems
            }
    
    return summary_dfs

def main():
    # Get mouse info
    mouse_info = get_mouse_info()
    
    # Process feed files
    feed_files = [f for f in os.listdir('.') if f.endswith('.CSV') and 'FEED1' in f]
    for file in feed_files:
        mouse_info = process_feed_file(file, mouse_info)
    
    # Create summary sheets
    summary_dfs = create_summary_sheets(mouse_info)
    
    # Print raw data for verification
    print("\nRaw meal data:")
    for mouse_id, info in mouse_info.items():
        if info['treatment'] == 'Meal-Fed':
            print(f"\n{mouse_id}:")
            for day in ['D1', 'D7', 'D14', 'D20']:
                meals = info['meals'][day]
                meal1 = sum(meals['18:00-19:00'])
                meal2 = sum(meals['23:30-00:30'])
                meal3 = sum(meals['05:00-06:00'])
                if meal1 > 0 or meal2 > 0 or meal3 > 0:
                    print(f"\n{day}:")
                    if meal1 > 0: print(f"Meal 1 (18:00-19:00): {meal1:.2f}g")
                    if meal2 > 0: print(f"Meal 2 (23:30-00:30): {meal2:.2f}g")
                    if meal3 > 0: print(f"Meal 3 (05:00-06:00): {meal3:.2f}g")
    
    # Save to Excel with a default sheet if no data
    with pd.ExcelWriter('meal_analysis.xlsx') as writer:
        if not summary_dfs:
            pd.DataFrame().to_excel(writer, sheet_name='No Data')
        else:
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