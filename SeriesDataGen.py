import streamlit as st
import pandas as pd
import numpy as np

# Streamlit app title
st.title("Hourly Data Generator from Excel Series File")

# File uploader
uploaded_file = st.file_uploader("Upload your Excel file for data generation:", type=["xlsx", "xls"])

if uploaded_file:
    st.success("File uploaded successfully!")

    # Load the file
    raw_df = pd.read_excel(uploaded_file, header=None)

    try:
        # Clean the column: convert to string, strip spaces, and lowercase
        raw_df[1] = raw_df[1].astype(str).str.strip().str.lower()


        # Dynamically find rows for Start, End, and Frequency
        start_rows = raw_df[1][raw_df[1].str.contains("start", na=False)].index
        end_rows = raw_df[1][raw_df[1].str.contains("end", na=False)].index
        freq_rows = raw_df[1][raw_df[1].str.contains("frequency", na=False)].index

        # Validate rows exist
        if len(start_rows) == 0 or len(end_rows) == 0 or len(freq_rows) == 0:
            st.error("Could not find 'Start', 'End', or 'Frequency' in the uploaded file. Please check the file structure.")
            st.stop()

        # Extract values
        start_row = start_rows[0]
        end_row = end_rows[0]
        freq_row = freq_rows[0]

        start_date = pd.to_datetime(raw_df.iloc[start_row, 2], errors="coerce")  # Start Date
        end_date = pd.to_datetime(raw_df.iloc[end_row, 2], errors="coerce")      # End Date
        frequency = raw_df.iloc[freq_row, 2]                                     # Frequency

        # Validate start and end dates
        if pd.isnull(start_date) or pd.isnull(end_date):
            st.error("Start or End date is invalid. Please check your file formatting.")
            st.stop()

        # Validate frequency
        if frequency.lower() != "hourly":
            st.error("Only 'hourly' frequency is supported.")
            st.stop()

        # Locate "Series" and extract relevant data
        series_start_rows = raw_df[2][raw_df[2].astype(str).str.contains("series", case=False, na=False)].index

        if len(series_start_rows) == 0:
            st.error("Could not find 'Series' in the uploaded file. Please check the file structure.")
            st.stop()

        series_start_row = series_start_rows[0] + 1
        series_df = raw_df.iloc[series_start_row:, [2, 3, 4]].dropna()
        series_df.columns = ["Series", "Low", "High"]

        # Generate hourly timestamps
        timestamps = pd.date_range(start=start_date, end=end_date, freq="H")

        # Initialize output DataFrame
        output_df = pd.DataFrame({"timestamp": timestamps})

        # Generate random data for each series
        for _, row in series_df.iterrows():
            series_name = row["Series"]
            low, high = int(row["Low"]), int(row["High"])
            output_df[series_name] = np.random.uniform(low, high, len(timestamps))

        # Display generated data
        st.subheader("Generated Hourly Data:")
        st.dataframe(output_df)

        # Download button
        csv = output_df.to_csv(index=False).encode("utf-8")
        st.download_button(
            label="Download Generated Data as CSV",
            data=csv,
            file_name="generated_hourly_data.csv",
            mime="text/csv",
        )

    except Exception as e:
        st.error(f"An error occurred: {e}")
