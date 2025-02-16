import pandas as pd
import numpy as np
import os

def extract_cage_info():
    """Extract cage numbers from CSV files"""
    csv_files = [f for f in os.listdir('.') if f.endswith('.CSV') and 'FEED1' in f]
    cage_data = []
    
    print("Processing CSV files for cage numbers...")
    for file in csv_files:
        with open(file, 'r') as f:
            content = f.readlines()[:30]
            current_cage = None
            
            for line in content:
                if 'GROUP/CAGE:' in line:
                    current_cage = line.split(':')[1].strip()
                elif 'SUBJECT ID:' in line:
                    mouse_id = line.split(':')[1].strip()
                    if mouse_id and mouse_id != 'Empty' and current_cage:
                        mouse_id = mouse_id.replace('Mouse ', '').strip()
                        if mouse_id:
                            cage_data.append({
                                'Mouse_ID': mouse_id,
                                'Cage_Number': current_cage
                            })
    
    return pd.DataFrame(cage_data).drop_duplicates()

def extract_mouse_data(excel_file):
    """Extract mouse data from Data Summary (Both) sheet"""
    print("\nReading Excel data...")
    
    # Read the Excel file
    df = pd.read_excel(excel_file, sheet_name='Data Summary (Both)', header=None)
    
    # Find the row containing "Animal #" and "Strain" and "Treatment"
    header_row = None
    for idx, row in df.iterrows():
        if 'Animal #' in row.values:
            header_row = idx
            break
    
    if header_row is None:
        raise ValueError("Could not find header row with 'Animal #'")
    
    # Read data with correct header
    df = pd.read_excel(excel_file, sheet_name='Data Summary (Both)', header=header_row)
    
    # Clean column names
    df.columns = [str(col).strip() for col in df.columns]
    
    # Select and rename relevant columns
    mouse_data = df[['Animal #', 'Strain', 'Treatment']].copy()
    mouse_data.columns = ['Mouse_ID', 'Strain', 'Treatment']
    
    # Clean the data
    mouse_data = mouse_data.dropna(subset=['Mouse_ID'])
    mouse_data = mouse_data[mouse_data['Mouse_ID'].astype(str).str.match(r'^m\d+$')]
    
    # Remove any duplicate mouse IDs
    mouse_data = mouse_data.drop_duplicates(subset=['Mouse_ID'])
    
    return mouse_data

def main():
    try:
        # Get cage info
        cage_df = extract_cage_info()
        print("\nCage data:")
        print(cage_df)
        
        # Get mouse data from Excel
        mouse_data = extract_mouse_data('Experiment m Data Summary.xlsx')
        print("\nMouse data from Excel:")
        print(mouse_data)
        
        # Merge data
        final_data = mouse_data.merge(cage_df, on='Mouse_ID', how='left')
        
        # Save result
        final_data.to_csv('mouse_classification.csv', index=False)
        print("\nFinal classification data:")
        print(final_data)
        print("\nClassification saved to mouse_classification.csv")
        
        # Print missing cage numbers
        missing_cages = final_data[final_data['Cage_Number'].isna()]['Mouse_ID'].tolist()
        if missing_cages:
            print("\nMice missing cage numbers:")
            print(missing_cages)
            
    except Exception as e:
        print(f"Error: {e}")

if __name__ == "__main__":
    main()