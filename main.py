import streamlit as st
import pandas as pd
#import win32com.client
from datetime import datetime
import xlsxwriter


#Function to upload CSV file
@st.cache_data
def load_csv_data(file, data=None):
    #st.write(f"Loading CSV file {file.name}")
    #if file is not None:
    #    st.write(file)
    #Print the data of file
    if file is not None:
        new_data = pd.read_csv(file)
    else:
        st.write("No file uploaded")
        if data is None:
            return pd.DataFrame()
        else:
            data
    #if st.button('Print Data'):
    #    st.write("Printing Data in Tabular Format")
        #st.write(st.cache(pd.read_csv)(file))
    #    st.write(data)
    # Join the both dataframes on the Time(in CT) column
    if data is None:
        data = new_data
    else:
        data = pd.concat([data, new_data], axis=0)
    return data


def getMailsFromFolder(folder_name = None, subfolder_name = None):
    outlook = win32com.client.Dispatch("Outlook.Application")
    namespace = outlook.GetNamespace("MAPI")
    inbox = namespace.GetDefaultFolder(6) # "6" refers to the inbox
    folder = None
    for f in inbox.Folders:
        if f.Name == folder_name:
            folder = f
            break
    if folder is None:
        st.write(f'folder not found: {folder_name} in Inbox, Checking in default folder Inbox')
        messages = inbox.Items
    else:
        if subfolder_name is None:
            messages = folder.Items
        else:
            # Checking subfolder_name
            subfolder = None
            for sf in folder.Folders:
                if sf.Name == subfolder_name:
                    subfolder = sf
                    break
            messages = subfolder.Items
    # Last x days mails in inbox
    messages.Sort("[ReceivedTime]", True)
    #messages = messages.Restrict("[ReceivedTime] >= '09/09/2024'")
    st.write(messages)


def main():
    st.title("CSV Visualization")
    st.write("Welcome to the Dashboard Developed by Sunil")
    #Sidebar
    #st.sidebar.header('API Stats Dashboard')
    #st.sidebar.text('This is a dashboard Page')
    #st.sidebar.balloons()
    #st.sidebar.markdown('---')

    st.write("Please upload the data to visualize")
    csv_data = st.file_uploader('Upload CSV', type=['csv'])
    if csv_data is not None:
        csv_data = load_csv_data(csv_data)
    
    if 'csv_data' not in locals() or csv_data is None:
        return
    #st.write(csv_data)
    #Replace missing values with 0
    csv_data = csv_data.fillna(0)
    # plot the graph
    if csv_data is not None and not csv_data.empty:
        time_field = st.selectbox('Select Time Field', csv_data.columns)
        if time_field not in csv_data.columns:
            st.write(f"The dataset must contain a {time_field} column.")
        else:
            selected_fields = st.multiselect('Select Columns to plot', [col for col in csv_data.columns if col not in [time_field]])
            if selected_fields:
                start_time, end_time = st.select_slider('Select Start Time', options=csv_data[time_field].values, value=(csv_data[time_field].values[0],csv_data[time_field].values[-1] ))
                # Filter the data
                csv_data = csv_data[(csv_data[time_field] >= start_time) & (csv_data[time_field] <= end_time)]
                plt_data = pd.DataFrame(csv_data[[time_field] + selected_fields])
                plt_data.set_index(time_field, inplace=True, drop=True, append=False, verify_integrity=False)
                #st.write(plt_data)
                st.line_chart(plt_data, use_container_width=True, height=500, width=30)
            else:
                st.write("Please select at least one column to plot.")
    else:
        st.write("Please upload the data to plot the graph")




if __name__ == "__main__":
    # Run Steamlit App
    st.set_page_config(page_title='API Stats Dashboard', page_icon=':bar_chart:', layout='wide', initial_sidebar_state='auto')
    #st.set_option('deprecation.showfileUploaderEncoding', False)
    st.set_option('deprecation.showPyplotGlobalUse', False)
    main()
