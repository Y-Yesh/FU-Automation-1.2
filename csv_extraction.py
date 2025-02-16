import pandas as pd
import json

from definitions import definitions

columns_to_extract = [
    'Technical Foundation', 'Personal Readiness', 'External Awareness',
    'Process Integration', 'Department Integration',
    'Implementation Impact', 'Individual Score', 'Integrated Score',
    'Total Score'
]

columns_to_extract_personnel = [
    'Company','Country','Department','Technical Foundation', 'Personal Readiness', 'External Awareness',
    'Process Integration', 'Department Integration',
    'Implementation Impact', 'Individual Score', 'Integrated Score',
    'Total Score'
]

# Load the CSV file (update the path as needed)
df = pd.read_csv(r"C:\Users\GOOD\Downloads\data_final.csv")


# -------------------------------------- Clean and format the data --------------------------------

# Trim spaces from column names
df.columns = df.columns.str.strip()

# Removing space before the comma in the 'Participant' column / Names
df['Participant'] = df['Participant'].str.replace(r'\s+,', ',', regex=True)


# -------------------------------------- Filtering DataFrame --------------------------------

df_filtered = df.dropna(subset=['Department', 'Position'])

df_company = df[df['Company'] == 'Company']

df_country = df[df['Company'] == 'Country']


# --------------------- Finding Most Frequent Company mentioned in the 'Company' column ------------------------

most_frequent_company = df_filtered['Company'].value_counts().idxmax()
most_frequent_count = df_filtered['Company'].value_counts().max()



# ---------------------- Dict Creation --------------------------

participant_dict = df_filtered.set_index('Participant')[columns_to_extract_personnel].to_dict(orient='index')

company_dict = df_company.set_index('Participant')[columns_to_extract].to_dict(orient='index')

country_dict = df_country.set_index('Participant')[columns_to_extract].to_dict(orient='index')



# ------------------------- Displaying -----------------------------------------

print(json.dumps(participant_dict,indent=4))


print(f'No of Participants -------------> {len(participant_dict.keys())}')

print(f"The most frequently listed company is '{most_frequent_company}' with {most_frequent_count} occurrences.")


print(json.dumps(company_dict,indent=4))

print(json.dumps(country_dict,indent=4))

