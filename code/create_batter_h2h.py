import pandas as pd

# Load the data
df = pd.read_csv('data/1995plays.csv')

# Filter to only include regular season games (exclude lcs, worldseries, allstar)
df_regular_season = df[~df['gametype'].isin(['lcs', 'worldseries', 'allstar'])]

print(f"Original DataFrame shape: {df.shape}")
print(f"Regular season DataFrame shape: {df_regular_season.shape}")

# Columns to sum
columns_to_sum = ['pa', 'ab', 'single', 'double', 'triple', 'hr', 'sh', 'sf', 'hbp', 'walk', 'iw', 'k', 'xi']

# Group by batter and pitcher, then sum the specified columns
batter_h2h = df_regular_season.groupby(['batter', 'pitcher'])[columns_to_sum].sum().reset_index()

# Display the results
print(f"\nBatter-Pitcher head-to-head combinations: {len(batter_h2h)}")
print("\nFirst 10 rows of the head-to-head data:")
print(batter_h2h.head(10))

# Show some statistics
print(f"\nSummary statistics:")
print(f"Total unique batter-pitcher combinations: {len(batter_h2h)}")
print(f"Total plate appearances across all matchups: {batter_h2h['pa'].sum()}")
print(f"Total at-bats across all matchups: {batter_h2h['ab'].sum()}")

# Save to CSV
output_file = 'data/1995 batter h2h.csv'
batter_h2h.to_csv(output_file, index=False)
print(f"\nHead-to-head data saved to: {output_file}")

# Show some interesting matchups (top 10 by plate appearances)
print("\nTop 10 matchups by plate appearances:")
top_matchups = batter_h2h.nlargest(10, 'pa')
print(top_matchups[['batter', 'pitcher', 'pa', 'ab', 'hr', 'k']]) 