import pandas as pd

print("Step 1: Reading the Excel file for properties...")
df = pd.read_excel('data.xlsx')
print("Properties loaded.")

print("Step 1.1: Reading the excluded tags from an Excel file...")
excluded_tags_df = pd.read_excel('tags.xlsx')
excluded_tags = excluded_tags_df['TAGS'].str.lower().tolist()
print("Excluded tags loaded.")

print("Filtering out rows based on TAGS column containing specified values...")
# Normalize and trim each tag before comparison for both the tag and the tags in the row
df = df[~df['TAGS'].apply(lambda x: any(tag.strip().lower() in [y.strip().lower() for y in str(x).split(',')] for tag in excluded_tags))]
print("Filtering based on TAGS completed.")

print("Removing rows where PROPERTY STATUS is 'Dead Lead'...")
df = df[df['PROPERTY STATUS'] != 'Dead Lead']
print("Filtering based on PROPERTY STATUS completed.")

print("Sorting by 'SCORE' in descending order...")
df = df.sort_values(by='SCORE', ascending=False)
print("Sorting completed.")

print("Removing duplicates based on 'MAILING ADDRESS' and 'MAILING ZIP'...")
df = df.drop_duplicates(subset=['MAILING ADDRESS', 'MAILING ZIP'], keep='first')
print("Duplicate removal completed.")

print("Removing properties with BUYBOX SCORE equal to 0...")
df = df[df['BUYBOX SCORE'] != 0]
print("Properties with BUYBOX SCORE equal to 0 removed.")

print("Removing properties with blank MAILING ADDRESS...")
df = df[df['MAILING ADDRESS'].notna() & df['MAILING ADDRESS'] != '']
print("Properties with blank MAILING ADDRESS removed.")

print("Saving the modified DataFrame back to an Excel file...")
df.to_excel('filtered_properties.xlsx', index=False)
print("Data saved to 'filtered_properties.xlsx'.")

print(f"Rows (properties) left after processing: {len(df)}")
