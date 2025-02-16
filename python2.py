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
                mouse_info[mouse_id] = {
                    'start_date': date,
                    'strain': row['Strain'],
                    'treatment': row['Treatment']
                }
    return mouse_info

def process_feed_file(file_path, target_date):
    """Process a single FEED1 file for a specific date"""
    experiment_start = None
    cage_to_mouse = {}
    daily_data = {}
    
    with open(file_path, 'r', encoding='utf-8', errors='ignore') as f:
        lines = f.readlines()
        
        # Get metadata
        for line in lines:
            line = line.strip()
            if 'EXPERIMENT START:' in line:
                experiment_start = datetime.strptime(line.split(':', 1)[1].strip(), '%d/%m/%Y %H:%M:%S')
                if experiment_start.date() != target_date.date():
                    return None, {}
            elif 'GROUP/CAGE:' in line:
                current_cage = line.split(':')[1].strip()
            elif 'SUBJECT ID:' in line:
                mouse_id = line.split(':')[1].strip().replace('Mouse ', '').strip()
                if mouse_id and mouse_id != 'Empty':
                    cage_to_mouse[current_cage] = mouse_id
                    daily_data[mouse_id] = {
                        '18:00-19:00': [],
                        '23:30-00:30': [],
                        '05:00-06:00': []
                    }
        
        # Process data section
        data_start = False
        for line in lines:
            if ':DATA' in line:
                data_start = True
                continue
            if data_start and ',' in line and '===' not in line and 'INTERVAL' not in line:
                parts = [p.strip() for p in line.split(',')]
                if len(parts) >= 9:
                    try:
                        time = datetime.strptime(parts[1], '%d/%m/%Y %H:%M:%S')
                        hour = time.hour
                        minute = time.minute
                        
                        # Process each cage
                        for i in range(2, 9, 3):
                            cage = parts[i]
                            if cage in cage_to_mouse:
                                mouse_id = cage_to_mouse[cage]
                                value = float(parts[i+1]) if parts[i+1].strip() else 0
                                
                                # Check time windows
                                if (hour == 18 and 0 <= minute < 60):
                                    daily_data[mouse_id]['18:00-19:00'].append(value)
                                elif (hour == 23 and 30 <= minute < 60) or (hour == 0 and 0 <= minute <= 30):
                                    daily_data[mouse_id]['23:30-00:30'].append(value)
                                elif (hour == 5 and 0 <= minute < 60):
                                    daily_data[mouse_id]['05:00-06:00'].append(value)
                    except:
                        continue
    
    return experiment_start, daily_data

def main():
    # Get mouse info
    mouse_info = get_mouse_info()
    
    # Initialize results
    results = []
    
    # Process each mouse
    for mouse_id, info in mouse_info.items():
        target_date = info['start_date']
        
        # Create base record
        record = {
            'Animal #': mouse_id,
            'Treatment': f"{info['strain']}-{info['treatment']}",
            'Date': target_date.strftime('%d/%m/%y'),
            '18:00-19:00': 0,
            '18:00-19:00_Total': 0,
            '23:30-00:30': 0,
            '23:30-00:30_Total': 0,
            '05:00-06:00': 0,
            '05:00-06:00_Total': 0
        }
        
        # Find and process relevant feed files
        feed_files = [f for f in os.listdir('.') if f.endswith('.CSV') and 'FEED1' in f]
        for file in feed_files:
            start_date, data = process_feed_file(file, target_date)
            if start_date and mouse_id in data:
                mouse_data = data[mouse_id]
                for window in ['18:00-19:00', '23:30-00:30', '05:00-06:00']:
                    values = mouse_data[window]
                    if values:
                        # Calculate average and total for each time window
                        record[window] = np.mean(values)
                        record[f"{window}_Total"] = sum(values)
        
        results.append(record)
    
    # Create DataFrame and save
    df = pd.DataFrame(results)
    df.to_csv('mouse_feeding_data.csv', index=False)
    print("\nProcessed data:")
    print(df)

if __name__ == "__main__":
    main()