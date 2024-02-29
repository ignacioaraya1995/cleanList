import pandas as pd

print("Step 1: Reading the Excel file for properties...")
initial_properties = pd.read_excel('data.xlsx')
print(f"Properties loaded. Total count: {len(initial_properties)}")

print("Step 1.1: Reading the excluded tags from an Excel file...")
excluded_tags_df = pd.read_excel('tags.xlsx')
excluded_tags = excluded_tags_df['TAGS'].str.lower().tolist()
print("Excluded tags loaded.")

initial_count = len(initial_properties)

# Apply filtering and processing while tracking the count
print("Filtering out rows based on TAGS column containing specified values...")
filtered_tags_count = initial_count - len(initial_properties[~initial_properties['TAGS'].apply(lambda x: any(tag.strip().lower() in [y.strip().lower() for y in str(x).split(',')] for tag in excluded_tags))])
df = initial_properties[~initial_properties['TAGS'].apply(lambda x: any(tag.strip().lower() in [y.strip().lower() for y in str(x).split(',')] for tag in excluded_tags))]
print("Filtering based on TAGS completed.")

print("Removing rows where PROPERTY STATUS is 'Dead Lead'...")
dead_lead_count = len(df) - len(df[df['PROPERTY STATUS'] != 'Dead Lead'])
df = df[df['PROPERTY STATUS'] != 'Dead Lead']
print("Filtering based on PROPERTY STATUS completed.")

print("Sorting by 'SCORE' in descending order...")
df = df.sort_values(by='SCORE', ascending=False)
print("Sorting completed.")

print("Removing duplicates based on 'MAILING ADDRESS' and 'MAILING ZIP'...")
duplicates_count = len(df) - len(df.drop_duplicates(subset=['MAILING ADDRESS', 'MAILING ZIP'], keep='first'))
df = df.drop_duplicates(subset=['MAILING ADDRESS', 'MAILING ZIP'], keep='first')
print("Duplicate removal completed.")

print("Removing properties with BUYBOX SCORE equal to 0...")
buybox_zero_count = len(df) - len(df[df['BUYBOX SCORE'] != 0])
df = df[df['BUYBOX SCORE'] != 0]
print("Properties with BUYBOX SCORE equal to 0 removed.")

print("Removing properties with blank MAILING ADDRESS...")
blank_address_count = len(df) - len(df[df['MAILING ADDRESS'].notna() & df['MAILING ADDRESS'] != ''])
df = df[df['MAILING ADDRESS'].notna() & df['MAILING ADDRESS'] != '']
print("Properties with blank MAILING ADDRESS removed.")

final_count = len(df)
print("Saving the modified DataFrame back to an Excel file...")
df.to_excel('filtered_properties.xlsx', index=False)
print("Data saved to 'filtered_properties.xlsx'.")

# Print summary table
print("\nSummary of Properties Removed:")
print(f"{'Initial Properties Count:':<40} {initial_count}")
print(f"{'Removed by TAGS:':<40} {filtered_tags_count}")
print(f"{'Removed by Dead Lead Status:':<40} {dead_lead_count}")
print(f"{'Removed as Duplicates:':<40} {duplicates_count}")
print(f"{'Removed with BUYBOX SCORE = 0:':<40} {buybox_zero_count}")
print(f"{'Removed with Blank MAILING ADDRESS:':<40} {blank_address_count}")
print(f"{'Final Properties Count:':<40} {final_count}")
