import pandas as pd
import re
import datetime
import streamlit as st

st.title("Manufactuer and MNP matcher and statistics")

def clean_special_and_split_words(text):
    if not isinstance(text, str):
        text = str(text)
        
    text = text.replace('-', ' ').replace('+', ' ').replace("ö", "o")
    text_cleaned = re.sub(r'[^a-zA-Z0-9.\s]', '', text)
    text_cleaned = re.sub(r'(?<=\w)\.(?=\w)', ' ', text_cleaned)
    text_cleaned = text_cleaned.upper()
    words = text_cleaned.split()
    
    return words


def clean_special_and_split_letter(text):
    if not isinstance(text, str):
        text = str(text)
        
    text = text.replace("ö", "o").replace("TBD-", "")
    text_cleaned = re.sub(r'[^a-zA-Z0-9]', '', text).upper()
    
    return text_cleaned


def compare_cells(row):
    
    a = row['Manufacturer_implementation_clean']
    b = row['Manufacturer_client_clean']
    if any(word in a for word in b):
        return 1
    else:
        return 0
    

def match_percentage(row):
    
    match_percentage = 0
    try:
        a = str(row['Mnp_implementation_clean'])
        b = str(row['Mnp_client_clean'])
        
        if a == b:
            match_percentage = 100
        
        match_count = 0
        min_len = min(len(a), len(b))
        for i in range(min_len):
            if a[i] == b[i]:
                match_count += 1
        
        match_percentage = round((match_count / max(len(a), len(b))) * 100, 1)
        
    except :
        match_percentage = 0
        
    return match_percentage

    
def manufacturer_matching(df_merged):

    df_merged = df_merged[["Item_code", "Manufacturer_client", "Manufacturer_implementation"]]
    df_merged['Item_code'] = df_merged['Item_code'].astype(str)
    df_merged = df_merged.sort_values(by="Item_code")
    df_merged.dropna(inplace=True)
    df_merged['match_manufacturer'] = 0
    
    df_merged.loc[:, 'Manufacturer_client'] = df_merged['Manufacturer_client']
    df_merged.loc[:, 'Manufacturer_client_clean'] = df_merged['Manufacturer_client'].apply(clean_special_and_split_words)

    df_merged.loc[:, 'Manufacturer_implementation'] = df_merged['Manufacturer_implementation']
    df_merged.loc[:, 'Manufacturer_implementation_clean'] = df_merged['Manufacturer_implementation'].apply(clean_special_and_split_words)

    df_merged['match_manufacturer'] = df_merged.apply(compare_cells, axis=1)
    
    return df_merged


def mnp_matching(df_merged):
    
    df_merged = df_merged[["Item_code", "Mnp_client", "Mnp_implementation"]]
    df_merged['Item_code'] = df_merged['Item_code'].astype(str)
    df_merged = df_merged.sort_values(by="Item_code")
    df_merged.dropna(inplace=True)
    df_merged['match_mnp'] = 0
    
    df_merged.loc[:, 'Mnp_client'] = df_merged['Mnp_client']
    df_merged.loc[:, 'Mnp_client_clean'] = df_merged['Mnp_client'].apply(clean_special_and_split_letter)

    df_merged.loc[:, 'Mnp_implementation'] = df_merged['Mnp_implementation']
    df_merged.loc[:, 'Mnp_implementation_clean'] = df_merged['Mnp_implementation'].apply(clean_special_and_split_letter)
    
    df_merged['match_mnp'] = df_merged.apply(match_percentage, axis=1)
    
    return df_merged


def manufacturer_improvment(df_client, df_implementation):
    
    total_manufacturer_client = df_client[df_client["Item_code"] != None ].shape[0]
    total_manufacturer_implementation = df_implementation[df_implementation["Item_code"] != None ].shape[0]

    known_manufacturer_client = df_client[df_client["Manufacturer_client"] != "TBD - To Be Determined" ].shape[0]
    known_manufacturer_implementation = df_implementation[df_implementation["Manufacturer_implementation"] != "TBD - To Be Determined" ].shape[0]

    known_manufacturer_client = round((known_manufacturer_client / total_manufacturer_client) * 100, 0)
    known_manufacturer_atthemoment = round((known_manufacturer_implementation / total_manufacturer_client) * 100, 0)
    known_manufacturer_implementation = round((known_manufacturer_implementation / total_manufacturer_implementation) * 100, 0)

    return known_manufacturer_client, known_manufacturer_atthemoment, known_manufacturer_implementation


def mnp_improvment(df_client, df_implementation):
    
    total_mnp = df_client.shape[0]

    known_mnp_client = df_client[df_client["Mnp_client"] != "TBD - To Be Determined" ].shape[0]
    known_mnp_implementation = df_implementation[df_implementation["Mnp_implementation"] != "TBD - To Be Determined" ].shape[0]

    known_mnp_client = round((known_mnp_client / total_mnp) * 100, 0)
    known_mnp_implementation = round((known_mnp_implementation / total_mnp) * 100, 0)

    return known_mnp_client, known_mnp_implementation


def improvment(df_match_manuf, df_match_mnp):
    
    manufacturer_improved = (df_match_manuf[df_match_manuf["match_manufacturer"] == 0].shape[0])
    mnp_improved = (df_match_mnp[df_match_mnp["match_mnp"] != 100].shape[0])

    return manufacturer_improved, mnp_improved



def main(df_client, df_implementation, file_name):
    
    if df_client is not None and df_implementation is not None:
        df_merged = pd.merge(df_client, df_implementation, on="Item_code", how="outer")
    
    df_match_manuf = manufacturer_matching(df_merged)
    df_match_mnp = mnp_matching(df_merged)
    
    df_match = pd.merge(df_match_manuf, df_match_mnp, on="Item_code", how="outer")
    df_match.drop(["Mnp_implementation_clean", "Mnp_client_clean", "Manufacturer_client_clean", "Manufacturer_implementation_clean"], axis=1, inplace=True)   
    
    known_manufacturer_client, known_manufacturer_atthemoment, known_manufacturer_implementation = manufacturer_improvment(df_client, df_implementation)
    known_mnp_client, known_mnp_implementation = mnp_improvment(df_client, df_implementation)
    manufacturer_improved, mnp_improved = improvment(df_match_manuf, df_match_mnp)
    
    today_date = datetime.date.today()
    output_file = f'{file_name}_match_{today_date}.xlsx'
    df_match.to_excel(output_file, sheet_name='Results', index=False)
    
    st.markdown(f"""
    - **{known_manufacturer_client}%** of manufacturers where known.
    - **{known_manufacturer_atthemoment}%** of the total manufacturers have been provided up to today.
    - **{known_manufacturer_implementation}%** of implemented manufacturers have been provided.
    - **{known_mnp_client}%** of MNPs are known where applicable, with **{known_mnp_implementation}%** currently implemented.
    - We have improved **{manufacturer_improved}** manufacturers and **{mnp_improved}** MNPs.
    
    Your excel report has been generated!
    """)


with st.sidebar:
    st.header('Files Upload')
    
    uploaded_client_file = st.file_uploader("Drag and drop client file here, or click to browse", type=["xlsx"], key = "client_file" )    
    uploaded_implementation_file = st.file_uploader("Drag and drop client sync extract here, or click to browse", type=["xlsx"], key = "implementation_file")
    df_client = None
    df_implementation = None 

    if uploaded_client_file is not None:
        st.write("Client file uploaded successfully!")
        df_client = pd.read_excel(uploaded_client_file, engine='openpyxl')
        columns_client = df_client.columns
        
        columns_client_selected = []
        st.write("Select clients columns:")
        for column in columns_client:
            if st.checkbox(column, key=f"checkbox_{column}"):
                columns_client_selected.append(column)
        
        if columns_client_selected:
            st.markdown("### Selected Columns")
            st.write(columns_client_selected)
        else:
            st.write("No columns selected.")
        
        df_client = df_client[columns_client_selected]
        df_client.rename(columns={columns_client_selected[0]: "Item_code", columns_client_selected[1]: "Manufacturer_client", columns_client_selected[2]: "Mnp_client"}, inplace=True)
        
        
    if uploaded_implementation_file is not None:
        st.write("Implementation file uploaded successfully!")
        df_implementation = pd.read_excel(uploaded_implementation_file, engine='openpyxl')
        df_implementation = df_implementation[["Item", "Manufacturer (Item) (Item)", "Mfr. Part # (Item) (Item)"]]
        df_implementation.rename(columns={"Item": "Item_code", "Manufacturer (Item) (Item)": "Manufacturer_implementation", "Mfr. Part # (Item) (Item)": "Mnp_implementation"}, inplace=True)

file_name = st.text_input("What is the name of the site?")
if all(v is not None for v in [df_client, df_implementation]):
    if st.button("Run"):
        main(df_client, df_implementation, file_name)
