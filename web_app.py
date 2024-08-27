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
        
        if match_percentage >= 75:
            match_percentage = 100
        else : 
            match_percentage = 0
        
    except :
        match_percentage = 0
        
    return match_percentage


def legacy_to_manufacturer_tbd(df_implementation):
    
    mask = df_implementation["Manufacturer_implementation"] == "TBD - To Be Determined"
    if mask.any():
        df_implementation.loc[mask, "Manufacturer_implementation"] = df_implementation.loc[mask, "Legacy_manufacturer"]
    
    return df_implementation


def red_tag_counter(df_implementation):
    
    df_implementation["Red_tag"] = 0
    df_implementation.loc[df_implementation["Description"].str.contains("RED TAG", na=False, case=False), "Red_tag"] = 1
    df_implementation.loc[df_implementation["Mnp_implementation"].str.contains("tbd", na=False, case=False), "Red_tag"] = 1
    df_implementation.loc[df_implementation["Mnp_implementation"].str.contains("TBD", na=False, case=False), "Red_tag"] = 1
    df_implementation.loc[df_implementation["Manufacturer_implementation"] == "TBD - To Be Determined", "Red_tag"] = 1
    
    return df_implementation[df_implementation["Red_tag"] == 1].shape[0]


    
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


def main(df_client, df_implementation, file_name, items_created):
     
    original_items = df_client.shape[0]
    
    items_processed = df_implementation.shape[0]
    
    df_implementation = legacy_to_manufacturer_tbd(df_implementation)
    
    red_tags = red_tag_counter(df_implementation)
    
    df_implementation = df_implementation[df_implementation["Manufacturer_implementation"] != "TBD - To Be Determined"]
    df_implementation.dropna(inplace=True)
    
    manufacturer_client = df_client[df_client["Manufacturer_client"] != "TBD - To Be Determined"].shape[0]
    mnp_client = df_client[~df_client["Mnp_client"].str.contains("tbd", na=False, case=False)].shape[0]
    
    if df_client is not None and df_implementation is not None:
        df_merged = pd.merge(df_client, df_implementation, on="Item_code", how="outer")
    
    df_match_manuf = manufacturer_matching(df_merged)
    df_match_mnp = mnp_matching(df_merged)
    
    df_match = pd.merge(df_match_manuf, df_match_mnp, on="Item_code", how="outer")
    df_match.drop(["Mnp_implementation_clean", "Mnp_client_clean", "Manufacturer_client_clean", "Manufacturer_implementation_clean"], axis=1, inplace=True)   
    
    amount_of_manuf_match = df_match[df_match["match_manufacturer"] == 1].shape[0]
    amount_of_mnp_match = df_match[df_match["match_mnp"] == 100].shape[0]
    
    today_date = datetime.date.today()
    output_file = f'{file_name}_match_{today_date}.xlsx'
    df_match.to_excel(output_file, sheet_name='Results', index=False)
    
    st.markdown(f"""
    - **{original_items}** items from original data
        - {manufacturer_client} manufacturers in original datapack
        - {mnp_client} Mnp in original datapack
    - **{items_processed}** items have been processed
    - **{original_items - items_processed}** items remaining
    - **{items_created}** items have been created
    - **{red_tags}** red tags
    - **{amount_of_manuf_match / (items_processed + items_created) }** 
    - **{amount_of_mnp_match / (items_processed + items_created) }** 

    Your excel report has been generated!
    """)


with st.sidebar:
    st.header('Files Upload')
    
    uploaded_client_file = st.file_uploader("Drag and drop initial file here, or click to browse", type=["xlsx"], key = "client_file" )    
    uploaded_implementation_file = st.file_uploader("Drag and drop client SYNC extract here, or click to browse", type=["xlsx"], key = "implementation_file")
    df_client = None
    df_implementation = None 

    if uploaded_client_file is not None:
        st.write("Client file uploaded successfully!")
        df_client = pd.read_excel(uploaded_client_file, engine='openpyxl')
        
        columns_client = df_client.columns
        columns_client_selected = []
        st.write("Select clients columns:")
        st.write("You need to select Item number, Manufacturer and MNP in this order.")
        for column in columns_client:
            if st.checkbox(column, key=f"checkbox_{column}"):
                columns_client_selected.append(column)

        if len(columns_client_selected) == 3 :
            df_client = df_client[columns_client_selected]
            df_client.rename(columns={columns_client_selected[0]: "Item_code", columns_client_selected[1]: "Manufacturer_client", columns_client_selected[2]: "Mnp_client"}, inplace=True)
        
        
    if uploaded_implementation_file is not None:
        st.write("Implementation file uploaded successfully!")
        df_implementation = pd.read_excel(uploaded_implementation_file, engine='openpyxl')
        df_implementation = df_implementation[["Item", 
                                               "Manufacturer (Item) (Item)", 
                                               "Legacy Manufacturer (Item) (Item)", 
                                               "Mfr. Part # (Item) (Item)", 
                                               "Description (Item) (Item)", 
                                               "Modified By (Item) (Item)",
                                               "Created By"]]
        df_implementation.rename(columns={"Item": "Item_code", 
                                          "Manufacturer (Item) (Item)": "Manufacturer_implementation", 
                                          "Legacy Manufacturer (Item) (Item)" : "Legacy_manufacturer", 
                                          "Mfr. Part # (Item) (Item)": "Mnp_implementation", 
                                          "Description (Item) (Item)" : "Description", 
                                          "Modified By (Item) (Item)" : "Modified_by",
                                          "Created By" : "Created_by"}, inplace=True)

        implementers = df_implementation.Modified_by.unique()
        implementers_on_site = []
        st.write("Select implementer on site:")
        for implementer in implementers:
            if st.checkbox(implementer, key=f"checkbox_{implementer}"):
                implementers_on_site.append(implementer)
            
        items_created = df_implementation[df_implementation["Created_by"].isin(implementers_on_site)].shape[0]    

        if implementers_on_site:
            df_implementation = df_implementation[df_implementation["Modified_by"].isin(implementers_on_site)]

if all(v is not None for v in [df_client, df_implementation]) and len(columns_client_selected) == 3 :
    file_name = st.text_input("What is the name of the site?")
    if file_name is not None :
        if st.button("Run"):
            main(df_client, df_implementation, file_name, items_created)
            
