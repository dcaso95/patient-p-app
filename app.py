import streamlit as st
import pandas as pd
import re
import msoffcrypto
import io

def find_missing_dates(df_patient):
    """
    Finds missing continuous dates for any modality that contains 'INGRESO NO QUIR.' in its name.
    """
    ingreso_no_quir_df = df_patient[df_patient['Producto'].astype(str).str.contains('INGRESO NO QUIR.', case=False, na=False)].copy()
    
    if ingreso_no_quir_df.empty:
        return []
    
    # Get the unique dates for this patient and modality
    present_dates = pd.to_datetime(ingreso_no_quir_df['F. Actividad'].unique(), errors='ignore')

    # Remove any NaT values
    present_dates = present_dates[pd.notna(present_dates)]

    if len(present_dates) < 2:
        return []

    # Find the start and end dates
    start_date = present_dates.min()
    end_date = present_dates.max()

    # Create a continuous date range
    continuous_range = pd.date_range(start=start_date, end=end_date, freq='D')

    # Find the missing dates by comparing the two sets
    missing_dates = [date.strftime('%d-%m-%y') for date in continuous_range if date not in present_dates]
    
    return missing_dates

def process_data(df, person_under_control):
    """
    Processes the DataFrame to extract the required information.
    """
    if df.empty:
        st.warning("The selected sheet is empty. Please check the file.")
        return

    # Check for required columns
    required_columns = ['Liquidación', 'Paciente', 'F. Actividad', 'Producto', 'I. Liquidado', 'source_file']
    if not all(col in df.columns for col in required_columns):
        st.error(f"One or more required columns are missing. Please check the file headers. Missing columns: {list(set(required_columns) - set(df.columns))}")
        return

    # Filter the DataFrame for the specific person under control
    df_filtered = df[df['Liquidación'].astype(str).str.contains(str(person_under_control), na=False)]

    if df_filtered.empty:
        st.warning(f"No data found for the person under control: {person_under_control}")
        return

    # Get the list of all patients
    patients = df_filtered['Paciente'].unique()

    for patient in patients:
        st.subheader(f"Results for: {patient}")
        patient_data = df_filtered[df_filtered['Paciente'] == patient].copy()
        
        # Ensure 'F. Actividad' is a date type
        patient_data['F. Actividad'] = pd.to_datetime(patient_data['F. Actividad'], errors='ignore')

        # Ensure 'I. Liquidado' is numeric
        patient_data['I. Liquidado'] = pd.to_numeric(patient_data['I. Liquidado'].astype(str).str.replace(',', '.'), errors='ignore')
        patient_data.dropna(subset=['I. Liquidado'], inplace=True)
        
        # Sort the patient data by date, which correctly orders by year > month > day
        patient_data.sort_values(by='F. Actividad', inplace=True)

        # Calculate the total money for the patient
        total_money_patient = patient_data['I. Liquidado'].sum()
        st.write(f"**Total money for {patient}**: €{total_money_patient:,.2f}")
        
        # Find and display missing dates for 'INGRESO NO QUIR'
        missing_days = find_missing_dates(patient_data)
        if missing_days:
            st.warning(f"🚨 **Missing 'INGRESO NO QUIR' days:** {', '.join(missing_days)}")
        else:
            st.success("✅ No missing 'INGRESO NO QUIR' days found.")
            
        # Group by 'Producto' (modality)
        modalities = patient_data['Producto'].unique()
        
        for modality in modalities:
            st.markdown(f"**Modality: {modality}**")
            modality_data = patient_data[patient_data['Producto'] == modality].copy()
            
            # Sort the modality data by date, which correctly orders by year > month > day
            modality_data.sort_values(by='F. Actividad', inplace=True)
            
            # Money per day (summing values for the same date)
            money_per_day = modality_data.groupby('F. Actividad')['I. Liquidado'].sum()
            
            # Find the most common payment value
            most_common_value = None
            if not money_per_day.empty:
                value_counts = money_per_day.value_counts()
                if not value_counts.empty:
                    most_common_value = value_counts.idxmax()
            
            # Identify weird dates (outlier payments and dates from multiple files)
            weird_dates_payments = pd.Series(dtype='float64')
            if most_common_value is not None:
                weird_dates_payments = money_per_day[money_per_day != most_common_value]
            
            cross_file_dates_group = modality_data.groupby('F. Actividad')['source_file'].nunique()
            cross_file_dates = cross_file_dates_group[cross_file_dates_group > 1]
            
            # Combine all weird dates and sort them chronologically
            all_weird_dates_dt = weird_dates_payments.index.union(cross_file_dates.index).sort_values()
            
            if not all_weird_dates_dt.empty:
                st.markdown(f"**⚠️ Unusual Payments and Cross-File Dates for {modality}:**")
                for date_dt in all_weird_dates_dt:
                    date_str = date_dt.strftime('%d-%m-%y')
                    money = money_per_day.get(date_dt, 'N/A')
                    source_files_count = cross_file_dates.get(date_dt, 0)
                    
                    if source_files_count > 1:
                        st.write(f"  - **{date_str}**: €{money:,.2f} (from multiple files)")
                    else:
                        st.write(f"  - **{date_str}**: €{money:,.2f} (unusual payment)")

            # Number of days
            days_count = modality_data['F. Actividad'].nunique()
            st.write(f"• **Days in modality**: {days_count}")

            # List the days and money, correctly sorted
            dates_with_money = [f"{date.strftime('%d-%m-%y')} (€{money:,.2f})" for date, money in money_per_day.items()]
            st.write(f"• **Dates and Money**: {', '.join(dates_with_money)}")
            
            # Total money for this modality
            total_money_modality = modality_data['I. Liquidado'].sum()
            st.write(f"• **Total money for this modality**: €{total_money_modality:,.2f}")
            st.markdown("---")

def main():
    # Set the page configuration to use a wide layout
    st.set_page_config(layout="wide")

    # This is the parameter you can change to adjust the sidebar width
    sidebar_width = 500  # Change this value to adjust the width in pixels

    # Use custom CSS to make the sidebar wider with the user-specified value
    st.markdown(
        f"""
        <style>
        [data-testid="stSidebar"] {{
            width: {sidebar_width}px;
        }}
        </style>
        """,
        unsafe_allow_html=True,
    )

    st.title('Patient Payment Data Analysis App')
    st.markdown("Upload your Excel files to analyze patient data")
    
    # Sidebar with instructions in Spanish
    st.sidebar.markdown("""
    ## Instrucciones de Uso
    
    1. **Sube tus archivos Excel (.xlsx o .xls):** Puedes seleccionar uno o varios archivos a la vez.
    
    2. **Ingresa la contraseña:** Si tus archivos están protegidos, introduce la contraseña para descifrarlos. Si no tienen contraseña, déjalo en blanco.
    
    3. **Haz clic en "Clear Data":** Si necesitas borrar los datos y empezar de nuevo, usa este botón.
    
    4. **Analiza los resultados:** La aplicación procesará los datos y mostrará un informe por cada paciente, incluyendo el total de dinero, los días que faltan y los pagos inusuales.
    
    
    Con cariño,
    Diego ❤️
    """)

    uploaded_files = st.file_uploader("Choose Excel files", type=['xlsx', 'xls'], accept_multiple_files=True)
    password = st.text_input("Enter password (if files are encrypted)", type="password")

    if st.button("Clear Data"):
        st.experimental_rerun()

    if uploaded_files:
        all_dfs = []
        st.info(f"Attached files: {', '.join([file.name for file in uploaded_files])}")

        for uploaded_file in uploaded_files:
            file_extension = uploaded_file.name.split('.')[-1].lower()
            df = None

            try:
                uploaded_bytes = io.BytesIO(uploaded_file.getvalue())

                if file_extension == 'xlsx':
                    if password:
                        try:
                            # Decrypt the file
                            office_file = msoffcrypto.OfficeFile(uploaded_bytes)
                            office_file.load_key(password=password)
                            decrypted_stream = io.BytesIO()
                            office_file.decrypt(decrypted_stream)
                            decrypted_stream.seek(0)
                            
                            df = pd.read_excel(decrypted_stream, sheet_name=1, engine='openpyxl')
                        except msoffcrypto.exceptions.DecryptionError:
                            st.error(f"Incorrect password for file: {uploaded_file.name}. Please try again.")
                            continue
                        except Exception as e:
                            st.warning(f"Could not decrypt {uploaded_file.name} with password. Attempting to open as a standard XLSX. Error: {e}")
                            uploaded_bytes.seek(0)
                            df = pd.read_excel(uploaded_bytes, sheet_name=1, engine='openpyxl')
                    else:
                        uploaded_bytes.seek(0)
                        df = pd.read_excel(uploaded_bytes, sheet_name=1, engine='openpyxl')
                
                elif file_extension == 'xls':
                    if password:
                        st.warning(f"Password-protected .xls files are not supported. Please save {uploaded_file.name} as a non-password-protected .xlsx and re-upload.")
                        continue
                    uploaded_bytes.seek(0)
                    df = pd.read_excel(uploaded_bytes, sheet_name=1, engine='xlrd')
                
                if df is not None:
                    # Add a new column to track the source file
                    df['source_file'] = uploaded_file.name
                    all_dfs.append(df)
            
            except Exception as e:
                st.error(f"An error occurred while processing {uploaded_file.name}: {e}. Please check the file format.")
        
        if all_dfs:
            combined_df = pd.concat(all_dfs, ignore_index=True)
            
            # Extract the name of the person under control from the first file
            person_under_control = None
            if 'Liquidación' in combined_df.columns and not combined_df['Liquidación'].empty:
                match = re.search(r'-(.*?)$', str(combined_df['Liquidación'].iloc[0]))
                if match:
                    person_under_control = match.group(1).strip()
            
            st.header(f"Report for: {person_under_control}")
            process_data(combined_df, person_under_control)
        else:
            st.warning("No valid Excel files were processed.")

if __name__ == "__main__":
    main()