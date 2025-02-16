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
                
                target_dates = {
                    'D1': date.date(),
                    'D7': (date + timedelta(days=6)).date(),
                    'D14': (date + timedelta(days=13)).date(),
                    'D20': (date + timedelta(days=19)).date()
                }
                
                mouse_info[mouse_id] = {
                    'start_date': date,
                    'target_dates': target_dates,
                    'strain': row['Strain'],
                    'treatment': row['Treatment'],
                    'meal_data': {
                        'D1': {'18:00-19:00': 0, '23:30-00:30': 0, '05:00-06:00': 0},
                        'D7': {'18:00-19:00': 0, '23:30-00:30': 0, '05:00-06:00': 0},
                        'D14': {'18:00-19:00': 0, '23:30-00:30': 0, '05:00-06:00': 0},
                        'D20': {'18:00-19:00': 0, '23:30-00:30': 0, '05:00-06:00': 0}
                    }
                }
                print(f"\nMouse {mouse_id}:")
                print(f"  Start date: {date.strftime('%Y-%m-%d')}")
                for day, target_date in target_dates.items():
                    print(f"  {day}: {target_date}")
    return mouse_info

def get_time_window(time):
    """Get the feeding time window for a given timestamp"""
    hour = time.hour
    minute = time.minute
    
    if hour == 18 and 0 <= minute < 60:
        return '18:00-19:00'
    elif (hour == 23 and 30 <= minute < 60) or (hour == 0 and 0 <= minute <= 30):
        return '23:30-00:30'
    elif hour == 5 and 0 <= minute < 60:
        return '05:00-06:00'
    return None

def process_feed_file(file_path, mouse_info):
    """Process a single feed file and extract meal data"""
    print(f"\nProcessing file: {file_path}")
    
    with open(file_path, 'r', encoding='utf-8', errors='ignore') as f:
        lines = f.readlines()
        
        # First pass: Get metadata
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
                if mouse_id and mouse_id != 'Empty' and mouse_id in mouse_info:
                    cage_to_mouse[current_cage] = mouse_id
                    print(f"Found mouse {mouse_id} in cage {current_cage}")
        
        print("\nProcessing data for mice:", list(cage_to_mouse.values()))
        
        # Second pass: Process data
        data_start = False
        for line in lines:
            line = line.strip()
            
            if ':DATA' in line:
                data_start = True
                continue
                
            if not data_start or '===' in line or 'INTERVAL' in line or not line:
                continue
                
            parts = [p.strip() for p in line.split(',')]
            if len(parts) < 9:
                continue
                
            # Process each cage's data
            for i in range(1, len(parts)-2, 3):
                if i+2 >= len(parts):
                    break
                    
                time_str = parts[i]
                cage = parts[i+1]
                value_str = parts[i+2]
                
                if cage in cage_to_mouse:
                    mouse_id = cage_to_mouse[cage]
                    mouse_data = mouse_info[mouse_id]
                    
                    try:
                        timestamp = datetime.strptime(time_str, '%d/%m/%Y %H:%M:%S')
                        current_date = timestamp.date()
                        time_window = get_time_window(timestamp)
                        
                        if time_window:  # If in a feeding window
                            try:
                                value = float(value_str.replace('E-', 'e-'))
                                if value > 0.02:  # Only count significant feeding events
                                    # Match to experimental day
                                    for day, target_date in mouse_data['target_dates'].items():
                                        if current_date == target_date:
                                            print(f"\nFound matching date for mouse {mouse_id}:")
                                            print(f"  Current date: {current_date}")
                                            print(f"  Target date: {target_date}")
                                            print(f"  Day: {day}")
                                            print(f"  Window: {time_window}")
                                            print(f"  Value: {value:.2f}g")
                                            
                                            # Only record if it's a valid time window for the treatment
                                            if mouse_data['treatment'] != 'Meal-Fed' or time_window:
                                                mouse_data['meal_data'][day][time_window] += value
                                                
                            except ValueError as e:
                                print(f"Error parsing value: {e}")
                                continue
                    except ValueError as e:
                        print(f"Error parsing timestamp: {e}")
                        continue
    
    return mouse_info

def create_meal_sheets(mouse_info):
    """Create DataFrames for each day's meal data"""
    day_dfs = {}
    
    for day in ['D1', 'D7', 'D14', 'D20']:
        print(f"\nCreating sheet for {day}")
        data = []
        
        for mouse_id, info in mouse_info.items():
            if info['treatment'] == 'Meal-Fed':  # Only include meal-fed mice
                meals = info['meal_data'][day]
                
                # Print data being added
                print(f"\nMouse {mouse_id}:")
                print(f"  Meal 1 (18:00-19:00): {meals['18:00-19:00']:.2f}g")
                print(f"  Meal 2 (23:30-00:30): {meals['23:30-00:30']:.2f}g")
                print(f"  Meal 3 (05:00-06:00): {meals['05:00-06:00']:.2f}g")
                
                strain_prefix = 'WT' if info['strain'] == 'C57BL/6' else 'GHSR'
                data.append({
                    'Animal #': mouse_id[1:],  # Remove 'm' prefix
                    'Treatment': f"{strain_prefix}-{info['treatment']}",
                    'Date': info['start_date'].strftime('%d/%m/%y'),
                    'Meal 1': meals['18:00-19:00'],
                    'Meal 2': meals['23:30-00:30'],
                    'Meal 3': meals['05:00-06:00']
                })
        
        df = pd.DataFrame(data)
        if not df.empty:
            # Calculate means and SEMs
            means = df.groupby('Treatment')[['Meal 1', 'Meal 2', 'Meal 3']].mean()
            sems = df.groupby('Treatment')[['Meal 1', 'Meal 2', 'Meal 3']].sem()
            
            df = df.sort_values('Treatment')
            day_dfs[day] = {
                'data': df,
                'means': means,
                'sems': sems
            }
    
    return day_dfs

def main():
    # Get mouse info
    mouse_info = get_mouse_info()
    
    # Process feed files
    feed_files = [f for f in os.listdir('.') if f.endswith('.CSV') and 'FEED1' in f]
    
    for file in feed_files:
        mouse_info = process_feed_file(file, mouse_info)
    
    # Create summary sheets
    day_dfs = create_meal_sheets(mouse_info)
    
    # Save to Excel
    with pd.ExcelWriter('meal_analysis.xlsx') as writer:
        for day, dfs in day_dfs.items():
            # Write data
            start_row = 0
            dfs['data'].to_excel(writer, sheet_name=day, startrow=start_row, index=False)
            
            # Write means
            start_row = len(dfs['data']) + 2
            dfs['means'].to_excel(writer, sheet_name=day, startrow=start_row)
            
            # Write SEMs
            start_row = start_row + len(dfs['means']) + 2
            dfs['sems'].to_excel(writer, sheet_name=day, startrow=start_row)
    
    print("\nAnalysis complete. Check meal_analysis.xlsx for results.")

if __name__ == "__main__":
    main()