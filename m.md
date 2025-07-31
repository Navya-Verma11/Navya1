import pandas as pd

# Ensure 'Estimated Hours' is in hours, and completion date is parsed
df['Completion Date'] = pd.to_datetime(df['Completion Date'], errors='coerce')
df['Estimated Hours'] = df['Estimated Hours'] / 60  # Convert minutes to hours if needed

# Filter only completed rows
completed_df = df[df['Completed'].str.lower() == 'yes'].copy()

# Flag YTD and July
completed_df['Is_YTD'] = completed_df['Completion Date'].between('2025-01-01', '2025-07-31')
completed_df['Is_July'] = completed_df['Completion Date'].between('2025-07-01', '2025-07-31')

# --- Group by Manager to aggregate employee hours ---
manager_grouped = completed_df.groupby(['Manager', 'Manager Email']).agg(
    Team_Total_Hours=('Estimated Hours', 'sum'),
    Team_YTD_Hours=('Estimated Hours', lambda x: x[completed_df.loc[x.index, 'Is_YTD']].sum()),
    Team_July_Hours=('Estimated Hours', lambda x: x[completed_df.loc[x.index, 'Is_July']].sum())
).reset_index()

# --- Extract manager metadata from employee rows ---
metadata_cols = [
    'Manager', 'Manager Email', 'Job Family', 'Job Title', 'Employee Status',
    'Leader Level 04', 'Leader Level 05', 'Leader Level 06', 'Business Line 04'
]

# Drop employee-specific columns and duplicates
manager_metadata = df[metadata_cols].drop_duplicates(subset=['Manager', 'Manager Email'])

# --- Merge metadata into aggregated manager view ---
manager_summary = manager_grouped.merge(manager_metadata, on=['Manager', 'Manager Email'], how='left')

# --- Rank managers based on total hours ---
manager_summary['Rank'] = manager_summary['Team_Total_Hours'].rank(method='dense', ascending=False).astype(int)
manager_summary = manager_summary.sort_values('Rank')

# Export to Excel
manager_summary.to_excel('managerial_view_from_employee_summary.xlsx', index=False)

print("✅ Managerial view based on employee data exported successfully!")
