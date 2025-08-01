import pandas as pd
import os

# Load the Excel file
excel_file = 'data/1977 rosters formatted.xlsx'

# Check if file exists
if not os.path.exists(excel_file):
    print(f"Error: File {excel_file} not found!")
    exit()

# Read all sheet names
xlsx = pd.ExcelFile(excel_file)
sheet_names = xlsx.sheet_names

print(f"Found {len(sheet_names)} sheets: {sheet_names}")

# List to store all DataFrames
all_teams = []

# Read each sheet, skipping row 1 and using row 2 as headers
for sheet_name in sheet_names:
    print(f"Processing sheet: {sheet_name}")
    
    # Read the sheet, skip row 0 (first row), use row 1 (second row) as headers
    df = pd.read_excel(excel_file, sheet_name=sheet_name, skiprows=1, header=0)
    
    # Add to the list
    all_teams.append(df)
    
    print(f"  - {sheet_name}: {len(df)} players")

# Combine all DataFrames
combined_df = pd.concat(all_teams, ignore_index=True)

print(f"\nCombined data shape: {combined_df.shape}")
print(f"Total players across all teams: {len(combined_df)}")

# Show the first few rows
print("\nFirst 5 rows of combined data:")
print(combined_df.head())

# Show column names
print(f"\nColumns in combined data:")
print(combined_df.columns.tolist())

# Save to new Excel file
output_file = 'data/1977 all ratings.xlsx'
combined_df.to_excel(output_file, index=False, sheet_name='All Teams')

print(f"\nCombined data saved to: {output_file}")

# Show summary by team
print("\nPlayers per team:")
team_counts = combined_df['Team'].value_counts()
print(team_counts) 