import streamlit as st
import pandas as pd
from io import BytesIO
import zipfile

# Upload File
def upload_file():
    uploaded_files = st.file_uploader("Upload Excel/CSV files", type=["xlsx", "csv"], accept_multiple_files=True)
    return uploaded_files

# Sheet Selection Box (only if .xlsx files exist)
def drop_box(uploaded_files):
    excel_files = [file for file in uploaded_files if file.name.endswith(".xlsx")]
    if excel_files:
        try:
            first_file = excel_files[0]
            sheet_names = pd.ExcelFile(first_file, engine='openpyxl').sheet_names
            selected_sheet = st.selectbox('Select sheet name to Concate', sheet_names)
            return selected_sheet
        except Exception as e:
            st.error(f"Error reading sheet names: {e}")
    return None

# concat Files Button Logic
def concat_button(uploaded_files, selected_sheet,header_row):
    if st.button('Concat Files') and uploaded_files:
        concat_df = pd.DataFrame()
        for file in uploaded_files:
            try:
                if file.name.endswith('.xlsx'):
                    if selected_sheet:
                        df = pd.read_excel(file, sheet_name=selected_sheet,header=header_row)
                    else:
                        st.warning(f'Skipping Excel file {file.name} due to no selected sheet.')
                        continue
                elif file.name.endswith('.csv'):
                    df = pd.read_csv(file, low_memory=False)
                else:
                    st.error(f'Unsupported file type: {file.name}')
                    continue

                # Optional: Add file name as a column
                # df['Source file'] = file.name
                concat_df = pd.concat([concat_df, df], ignore_index=True)
            except Exception as e:
                st.error(f'Error reading {file.name}: {e}')
        
        if not concat_df.empty:
            st.success('Files concatenated successfully')
            return concat_df
        else:
            st.warning('No data found to concate')
    return None

# Download Button
def download_merg(concat_df):
    if concat_df is not None:
        output = BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            concat_df.to_excel(writer, index=False, sheet_name="Concatenated")       
        output.seek(0)
        st.download_button(
            label="Download Concatenated Excel",
            data=output,
            file_name="concat_output.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

def add_css():
    with open('style.css') as f:
        st.markdown(f'<style>{f.read()}</style>', unsafe_allow_html=True)
        
def upload_file_split():
    uploaded_data = st.file_uploader('Upload an Excel/CSV file...', type=['xlsx','csv'])
    return uploaded_data

def drop_box_sheet(uploaded_data):
    if uploaded_data.name.endswith('.xlsx'):
        sheets = pd.ExcelFile(uploaded_data, engine='openpyxl').sheet_names
        selected_sheet = st.selectbox('Select sheet to split', sheets)
        return selected_sheet
    return None

def drop_box_col(uploaded_data, selected_sheet,header_row):
    if uploaded_data:
        if uploaded_data.name.endswith('.xlsx'):
            df = pd.read_excel(uploaded_data, sheet_name=selected_sheet,header=header_row)
            cols = df.columns
        elif uploaded_data.name.endswith('.csv'):
            df  = pd.read_csv(uploaded_data, header=header_row)
            cols =df.columns
        else:
            st.error('Unsupported file format')

        selected_col = st.selectbox('Select column to split', cols)
        return selected_col, df

def select_header():
    header_option = st.selectbox('Select the header row',options=['1st row', '2nd row', '3rd row'])
    if header_option == '1st row':
        header_row = 0 
    elif header_option == '2nd row':
        header_row = 1
    else:
        header_row = 2

    return header_row

    
def split_file_by_column(df, selected_col):
    zip_buffer = BytesIO()
    with zipfile.ZipFile(zip_buffer, "w") as zip_file:
        unique_values = df[selected_col].dropna().unique()
        for val in unique_values:
            filtered_df = df[df[selected_col] == val]
            output = BytesIO()
            filtered_df.to_excel(output, index=False)
            output.seek(0)
            file_name = f"{selected_col}_{val}.xlsx"
            zip_file.writestr(file_name, output.read())
    zip_buffer.seek(0)
    return zip_buffer

# Main App
def main():
    
    st.set_page_config(page_title="Ease to Excel",
                        page_icon='🟩',
                        layout="wide",
                        initial_sidebar_state="expanded")
   
    add_css()
    tab1, tab2 = st.tabs(['File Concatenator', 'File Splitter'])
    
    with tab1:
         
        col1, col2, col3 = st.columns([3,1,2])
        with col1:
            st.markdown('## Excel/CSV File Concatenator')
            uploaded_files = upload_file()
            # selected_sheet = None
            # header_row = None
            if uploaded_files:
                # Only show sheet dropdown if any Excel file is uploaded
                if any(file.name.endswith('.xlsx') for file in uploaded_files):
                    selected_sheet = drop_box(uploaded_files)
                header_row = select_header()
                # Pass selected_sheet (can be None for CSV-only files)
                concat_df = concat_button(uploaded_files, selected_sheet,header_row)
                download_merg(concat_df)

        with col3:
            st.info('''**Instructions:**  
                        1. Upload your files and choose the sheet name (for Excel only) based on which you want to merge the data.  
                        2. Ensure that all files have the same sheet name and columns name.  
                        3. Files with different sheet names will be skipped during the merging process.    
                        4. Avoid uploading large files, a single file over 10 MB may take time to load.  
                        5. If your file size exceeds 10MB, save that as a CSV format before uploading, for smooth transformation.''')
            
            st.markdown('### Concatenated File Details')
            if uploaded_files and 'concat_df' in locals() and concat_df is not None:
                st.write(f'Rows in concatenated file: {concat_df.shape[0]}')
                st.write(f'Columns in concatenated file: {concat_df.shape[1]}')

    with tab2:
        col1, col2, col3 = st.columns([3,1,2])
        with col1:
            st.markdown('## Excel/CSV file spliter')
            uploaded_data = upload_file_split()
            selected_sheet = None
            if uploaded_data:
                if uploaded_data.name.endswith('.xlsx'):
                    selected_sheet = drop_box_sheet(uploaded_data)
                header_row = select_header()
                selected_col, df = drop_box_col(uploaded_data, selected_sheet,header_row)

                if st.button("Split and Download"):
                    zip_file = split_file_by_column(df, selected_col)
                    st.success("Files split successfully!")
                    st.download_button(
                    label="Download Split Files (ZIP)",
                    data=zip_file,
                    file_name="split_files.zip",
                    mime="application/zip")
        with col3:
            st.info('''**Instructions:**  
                        1. Upload your file and choose the sheet name (for Excel only) and column name based on which you want to split the data.      
                        2. Avoid uploading large files, a file over 10 MB may take time to load.  
                        3. If your file size exceeds 10MB, save that as a CSV format before uploading, for smooth transformation.''')
            st.markdown('### Uploaded File Details')
            if uploaded_data:
                st.write(f'Rows in uploaded file: {df.shape[0]}')
                st.write(f'Columns in uploaded file: {df.shape[1]}')


# Run app
if __name__ == "__main__":
    main()
