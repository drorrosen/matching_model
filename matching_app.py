import streamlit as st
import pandas as pd
from Functions import *


st.image("logo.png", width=400)

st.title("MatchPoint: Ultimate Address Intelligence")

# Start a form block
st.sidebar.title("Upload Files")
with st.sidebar.form(key='file_upload_form'):
    # Create file uploaders
    uploaded_file_1 = st.file_uploader("Upload Internal Dataset", type=['xlsx', 'csv'])
    uploaded_file_2 = st.file_uploader("Upload Target Dataset", type=['xlsx', 'csv'])

    # Create a submit button
    submit_button = st.form_submit_button(label='Submit')

# Use the uploaded files
if submit_button:
    if uploaded_file_1 is not None and uploaded_file_2 is not None:
        # read uploaded_file_1 first and check if xlsx or csv
        if uploaded_file_1.name.endswith('.xlsx'):
            dataset_1 = pd.read_excel(uploaded_file_1)
        elif uploaded_file_1.name.endswith('.csv'):
            dataset_1 = pd.read_csv(uploaded_file_1)
        else:
            st.error("Invalid file format. Please upload an Excel file.")

        # read uploaded_file_2 and check if xlsx or csv
        if uploaded_file_2.name.endswith('.xlsx'):
            dataset_2 = pd.read_excel(uploaded_file_2)
        elif uploaded_file_2.name.endswith('.csv'):
            dataset_2 = pd.read_csv(uploaded_file_2)
        else:
            st.error("Invalid file format. Please upload an Excel file.")


            # Clean up the column names for both datasets
        dataset_1 = clean_column_names(dataset_1)
        dataset_2 = clean_column_names(dataset_2)


        #make sure everything is lowered case
        dataset_1['Address'] = dataset_1['Address'].str.lower()
        dataset_2['Address'] = dataset_2['Address'].str.lower()

        # Merge 'Address' and 'Suburb' in Dataset 2 for full address comparison
        dataset_2['Address'] = dataset_2['Address'].str.cat(dataset_2['Suburb'], sep=', ')

        # Normalize addresses by converting to lowercase and removing commas and spaces for consistent matching
        dataset_1['Address'] = dataset_1['Address'].str.lower().str.replace(',', '').str.strip()
        dataset_2['Address'] = dataset_2['Address'].str.lower().str.replace(',', '').str.strip()

        # Prepare names by combining 'First Name' and 'Last Name' in Dataset 1 and normalizing
        dataset_1['Full Name'] = dataset_1.apply(lambda row: f"{row['First Name']} {row['Last Name']}".strip().lower(), axis=1)

        # Ensure the 'Owner's Name' in Dataset 2 is also normalized for comparison
        dataset_2["Owner's Name"] = dataset_2["Owner's Name"].str.lower().str.strip()

        # Standardize addresses
        dataset_1['Address'] = dataset_1['Address'].apply(standardize_address)
        dataset_2['Address'] = dataset_2['Address'].apply(standardize_address)
        matches_df = find_best_matches(dataset_2, dataset_1)



    # Example of applying the updated function to all rows in dataset_2 and updating the mobile column
        #results = dataset_2.apply(lambda x: combined_matching(x['Address'], x["Owner's Name"], dataset_1), axis=1)
        #address_scores, name_scores = precompute_scores(dataset_1)



        # Assuming dataset_2 already has a 'Mobile' column you want to update
        dataset_2[['Best Match Address', 'Best Match Name', 'Combined Score', 'Mobile']] = matches_df[['Best Match Address', 'Best Match Name', 'Combined Score', 'Mobile']]
        #dataset_2['Combined Score'] = dataset_2['Address Score'] * dataset_2['Name Score']

        dataset_2['Combined Score'] = normalize_combined_score(dataset_2['Combined Score'])
        dataset_2['Confidence'] = dataset_2['Combined Score'].apply(confidence)
        #dataset_2 = dataset_2.drop(['Address Score', 'Name Score'], axis=1)



        # Sort the DataFrame by 'Combined Score' in descending order
        df_sorted = dataset_2.sort_values(by='Combined Score', ascending=False)

        # Create a mask to identify duplicates, excluding the first occurrence
        is_duplicate = df_sorted.duplicated(subset=['Best Match Address'], keep='first')

        # For the identified duplicates, replace all other columns with NaN
        df_sorted.loc[is_duplicate, ['Best Match Address', 'Mobile', 'Best Match Name']] = np.nan
        df_sorted.loc[is_duplicate, ['Combined Score']] = 0
        df_sorted.loc[is_duplicate, ['Confidence']] = 'No Match'


        st.write(df_sorted)

        st.write('Number of addresses with high confidence:', df_sorted[df_sorted['Confidence'] == 'High'].shape[0])
        st.write('Number of addresses with medium confidence:', df_sorted[df_sorted['Confidence'] == 'Medium'].shape[0])
        st.write('Number of addresses with low confidence:', df_sorted[df_sorted['Confidence'] == 'Low'].shape[0])


        # Determine the index of the 'Confidence' column (+1 because Excel columns are 1-indexed)
        confidence_col_index = df_sorted.columns.get_loc('Confidence') + 1

        # Create the styled Excel file
        excel_data = create_styled_excel(df_sorted, confidence_col_index)

        # Provide the download button in Streamlit
        st.download_button(
            label="Download Excel file",
            data=excel_data.read(),  # Use `.read()` to get the bytes
            file_name="styled_dataframe.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )








