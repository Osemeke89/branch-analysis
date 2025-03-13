import datetime
import random

import altair as alt
import numpy as np
import pandas as pd
import streamlit as st

# Show app title and description.
st.set_page_config(page_title="Support tickets", page_icon="🎫")
st.title("🎫 Branch Analysis")
st.write(
    """
    This app allows users to upload an Excel file and instantly analyze loan data, providing insights on total disbursements, healthy vs. unhealthy loans, and loan cycle performance—all in a simple and interactive dashboard.
    """
)

def analyze_loan_data(uploaded_file):
    # Read the uploaded Excel file
    excel_file = pd.ExcelFile(uploaded_file)
    sheet_names = excel_file.sheet_names

    # Load necessary sheets
    mcr_sheet = pd.read_excel(uploaded_file, sheet_name='MCR without Last Audit selectio')
    eom_sheet = pd.read_excel(uploaded_file, sheet_name='EOM Disbursement')

    # Define loan categories
    loan_categories = [
        'Micro Loan', 'Micro Plus', 'TAKA LOAN', 'SME Loan',
        'Building Improv Loan Category B', 'Building Improv Loan Category A',
        'SME PLUS Loan', 'OPPORTUNITY LOAN'
    ]

    # Initialize dictionary for storing results
    results = {}

    for category in loan_categories:
        filtered_data = mcr_sheet[mcr_sheet['CATEGORY DESC'] == category]

        results[category] = {
            'Total Disbursement': filtered_data.shape[0],
            'Healthy Loans': filtered_data[filtered_data['PD INDICATOR'] == 'N'].shape[0],
            'Unhealthy Loans': filtered_data[filtered_data['PD INDICATOR'] == 'Y'].shape[0],
            '1st Loan Cycle on PAR': filtered_data[(filtered_data['PD INDICATOR'] == 'Y') & (filtered_data['LOAN CYCLE'] == 1)].shape[0]
        }

    results_df = pd.DataFrame(results).T
    results_df.loc['Total'] = results_df.sum()

    # EOM Disbursement Analysis
    eom_results = {
        'Total Disbursement': eom_sheet.shape[0],
        'Healthy Loans': eom_sheet[eom_sheet['PD INDICATOR'] == 'N'].shape[0],
        'Unhealthy Loans': eom_sheet[eom_sheet['PD INDICATOR'] == 'Y'].shape[0],
        '1st Loan Cycle on PAR': eom_sheet[(eom_sheet['PD INDICATOR'] == 'Y') & (eom_sheet['LOAN CYCLE'] == 1)].shape[0]
    }

    results_df.loc['EOM Disbursement'] = eom_results

    return results_df

# Streamlit UI
st.title("Loan Data Analysis Web App")
uploaded_file = st.file_uploader("Upload an Excel file", type=['xlsx'])

if uploaded_file:
    results_df = analyze_loan_data(uploaded_file)
    st.write("### Analysis Result")
    st.dataframe(results_df)
    st.download_button("Download Analysis as CSV", data=results_df.to_csv(), file_name="loan_analysis.csv", mime="text/csv")
