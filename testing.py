import pandas as pd

# Use read_csv with sep='\t' because it's a tab-separated text file
file_path = 'data/awards_df.xls' 
df = pd.read_csv(file_path, sep='\t')

print("awards_df columns:", df.columns.tolist())

# Do the same for the second file
file_path_proposals = 'data/proposals_df.xls'
df_proposals = pd.read_csv(file_path_proposals, sep='\t')

print("proposals_df columns:", df_proposals.columns.tolist())