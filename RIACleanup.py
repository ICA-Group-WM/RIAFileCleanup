import pandas as pd
from fuzzywuzzy import process, fuzz

# Read the two excel files
df1 = pd.read_excel('RIA LPL Accts.xlsx')
df2 = pd.read_excel('Schwab.xlsx')

# Convert Zip Codes to only have the part before '-'
df1['Mailing Zip Code'] = df1['Mailing Zip Code'].str.split('-').str[0]


# Modify name handling for df2 to handle names without commas
def split_name(name):
    if ',' in name:
        return name.split(',', 1)
    else:
        return "", name

df2['Last Name'], df2['First Name'] = zip(*df2['Primary Account Holder'].map(split_name))
df2['Email'] = df2['Account Email Address']
df2['Mailing Address Line 1'] = df2['Address Line 1']
df2['Mailing City'] = df2['City']
df2['Mailing State'] = df2['State']
df2['Mailing Zip Code'] = df2['Zip']
df2['Mailing Zip Code'] = df2['Mailing Zip Code'].str.split('-').str[0]

# If there's no First Name or Last Name in df1, use 'Client' for 'First Name' and make 'Last Name' blank
df1['First Name'] = df1.apply(lambda x: x['Client'] if pd.isnull(x['First Name']) else x['First Name'], axis=1)
df1['Last Name'] = df1.apply(lambda x: "" if pd.isnull(x['Last Name']) else x['Last Name'], axis=1)
# Add "Mailing Address Line 2" to df2 if it doesn't exist
if 'Mailing Address Line 2' not in df2.columns:
    df2['Mailing Address Line 2'] = ""
# Convert all string columns to lowercase for both df1 and df2
string_columns = ['First Name', 'Last Name', 'Email', 'Mailing Address Line 1', 'Mailing Address Line 2', 'Mailing City', 'Mailing State', 'Mailing Zip Code']
for col in string_columns:
    df1[col] = df1[col].str.lower()
    df2[col] = df2[col].str.lower()
    
names_dropped = []
# Compare records and drop duplicates based on your criteria
indices_to_drop = []
for idx, row in df1.iterrows():
    matches = process.extract(row['First Name'], df2['First Name'], limit=df2.shape[0])
    matching_indices = [match[2] for match in matches if match[1] > 95]

    for index in matching_indices:
        match_row = df2.iloc[index]
        if row['Last Name'] == match_row['Last Name']:
            if row['Email'] == match_row['Email']:
                names_dropped.append((row['First Name'], row['Last Name']))
                indices_to_drop.append(index)
            elif pd.isna(row['Email']) and pd.isna(match_row['Email']):
                if (row['Mailing Address Line 1'] == match_row['Mailing Address Line 1'] and
                    row['Mailing City'] == match_row['Mailing City'] and
                    row['Mailing State'] == match_row['Mailing State'] and
                    row['Mailing Zip Code'] == match_row['Mailing Zip Code']):
                    names_dropped.append((row['First Name'], row['Last Name']))
                    indices_to_drop.append(index)
# Now drop the indices and reset the index
df2.drop(indices_to_drop, inplace=True)
df2.reset_index(drop=True, inplace=True)

print(f"Total similar duplicates dropped based on First Name, Last Name, Email, and Address: {len(names_dropped)}")
print("Names dropped:", names_dropped)

# Combine the dataframes
df_all = pd.concat([df1[['First Name', 'Last Name', 'Email', 'Mailing Address Line 1', 'Mailing Address Line 2', 'Mailing City', 'Mailing State', 'Mailing Zip Code']], df2[['First Name', 'Last Name', 'Email', 'Mailing Address Line 1', 'Mailing Address Line 2', 'Mailing City', 'Mailing State', 'Mailing Zip Code']]], ignore_index=True)

# Removing entries with missing emails for the "Emails" tab
df_with_email = df_all[df_all['Email'].notna()]
df_email_unique = df_with_email.drop_duplicates(subset='Email', keep='first')
print(f"Removed {len(df_with_email) - len(df_email_unique)} email duplicates and kept the first occurrences.")

# Checking for mailing address duplicates for the "Mailed" tab
df_no_email = df_all[df_all['Email'].isna()]
# This function creates a single string for address from a row for comparison
def get_full_address(row):
    return ' '.join([str(row['Mailing Address Line 1']), 
                     str(row['Mailing Address Line 2']), 
                     str(row['Mailing City']), 
                     str(row['Mailing State']), 
                     str(row['Mailing Zip Code'])])

# Get a list of all addresses
addresses = df_no_email.apply(get_full_address, axis=1)

# This set keeps track of indices we've already matched
checked_indices = set()

# We'll store the indices of addresses to label as "Household" here
household_indices = []

for idx, address1 in enumerate(addresses):
    if idx not in checked_indices:  # Skip already checked addresses
        # Find similar addresses; we use a list comprehension
        similar_addresses = [j for j, address2 in enumerate(addresses) 
                             if idx != j 
                             and fuzz.ratio(address1, address2) > 95]
        
        if similar_addresses:  # If there are similar addresses
            household_indices.append(idx)  # The current address becomes "Household"
            checked_indices.update([idx] + similar_addresses)  # Mark all these addresses as checked

# Update names for rows with similar addresses to "Household"
df_no_email.loc[df_no_email.index[household_indices], 'First Name'] = 'Household'
df_no_email.loc[df_no_email.index[household_indices], 'Last Name'] = ''

# Find rows with duplicate mailing addresses and set their names to "Household"
# We reuse the df_no_email variable here since we've already modified it for similar addresses
duplicate_address_mask = df_no_email.duplicated(subset=['Mailing Address Line 1', 'Mailing Address Line 2', 'Mailing City', 'Mailing State', 'Mailing Zip Code'], keep=False)
df_no_email.loc[duplicate_address_mask, 'First Name'] = 'Household'
df_no_email.loc[duplicate_address_mask, 'Last Name'] = ''

# Finally, we filter out the unique addresses to generate df_mail_unique
df_mail_unique = df_no_email.drop_duplicates(subset=['Mailing Address Line 1', 'Mailing Address Line 2', 'Mailing City', 'Mailing State', 'Mailing Zip Code'], keep='first')

# Capitalize names
df_mail_unique['First Name'] = df_mail_unique['First Name'].str.title()
df_mail_unique['Last Name'] = df_mail_unique['Last Name'].str.title()
df_email_unique['First Name'] = df_email_unique['First Name'].str.title()
df_email_unique['Last Name'] = df_email_unique['Last Name'].str.title()

# Capitalize address components
columns_to_title = ['Mailing Address Line 1', 'Mailing Address Line 2', 'Mailing State', 'Mailing City']
for col in columns_to_title:
    df_mail_unique[col] = df_mail_unique[col].str.title()
    df_email_unique[col] = df_email_unique[col].str.title()

df_mail_unique['Mailing State'] = df_mail_unique['Mailing State'].str.upper()
# Adjust output format for the both tabs
df_mail_unique['Name'] = df_mail_unique['First Name'] + ' ' + df_mail_unique['Last Name']
df_email_unique['Name'] = df_email_unique['First Name'] + ' ' + df_email_unique['Last Name']
# Create 'Address' column for the Mailed tab
def assemble_address(row):
    address_parts = [row['Mailing Address Line 1'], 
                     row['Mailing Address Line 2'], 
                     row['Mailing City'], 
                     row['Mailing State'], 
                     row['Mailing Zip Code']]
    return ', '.join([part for part in address_parts if part and not pd.isna(part)])

df_mail_unique['Address'] = df_mail_unique.apply(assemble_address, axis=1)

# Save to Excel with two tabs
with pd.ExcelWriter('CleanedAccounts.xlsx') as writer:
    df_email_unique[['Name', 'Email']].to_excel(writer, sheet_name='Emails', index=False)
    df_mail_unique[['Name', 'Address']].to_excel(writer, sheet_name='Mailing', index=False)
