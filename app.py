import streamlit as st
import pandas as pd
import io

# Define functions for data processing
def re_aggregate_sales_profitability(df):
    return df.groupby(['InvNo', 'BillTo']).agg({
        'AmountAfterTax': 'sum',
        'PurchaseAmount': 'sum'
    }).reset_index()

def aggregate_invoice_list(df):
    aggregated_df = df.groupby('PONo').agg({
        'DocNo': lambda x: ', '.join(map(str, x.unique()))
    }).reset_index()
    aggregated_df.columns = ['PONo', 'InvoiceNos_List']
    return aggregated_df

def calculate_profit(invoice_list, sales_profitability_df):
    filtered_df = sales_profitability_df[sales_profitability_df['InvNo'].isin(invoice_list.split(', '))]
    if not filtered_df.empty:
        return filtered_df['AmountAfterTax'].sum() - filtered_df['PurchaseAmount'].sum()
    return None

def fetch_billto(invoice_list, sales_profitability_df):
    filtered_df = sales_profitability_df[sales_profitability_df['InvNo'].isin(invoice_list.split(', '))]
    return filtered_df['BillTo'].unique()[0] if not filtered_df.empty else None

# Streamlit UI
st.title("Emaar PO Profitability Generator")

# Upload Excel files
st.sidebar.header("Upload Excel Files")
sales_register_file = st.sidebar.file_uploader("Upload Sales Register Excel File", type=["xls", "xlsx"])
sales_profitability_file = st.sidebar.file_uploader("Upload Sales Profitability Excel File", type=["xls", "xlsx"])

if sales_register_file is not None and sales_profitability_file is not None:
    # Data processing
    sales_register_df = pd.read_excel(sales_register_file, skiprows=3)
    sales_profitability_df = pd.read_excel(sales_profitability_file, skiprows=3)

    re_aggregated_df = re_aggregate_sales_profitability(sales_profitability_df)
    aggregated_invoice_list_df = aggregate_invoice_list(sales_register_df)

    aggregated_invoice_list_df['Profit'] = aggregated_invoice_list_df['InvoiceNos_List'].apply(
        lambda x: calculate_profit(x, re_aggregated_df))
    aggregated_invoice_list_df['BillTo'] = aggregated_invoice_list_df['InvoiceNos_List'].apply(
        lambda x: fetch_billto(x, re_aggregated_df))

    # Output as Excel
    st.subheader("Resulting Data")
    st.write(aggregated_invoice_list_df)

    output_filename = "output_result.xlsx"

    # Custom function for downloading Excel
    def download_excel():
        output_buffer = io.BytesIO()
        with pd.ExcelWriter(output_buffer, engine="xlsxwriter") as writer:
            aggregated_invoice_list_df.to_excel(writer, sheet_name="Sheet1", index=False)
        output_buffer.seek(0)
        return output_buffer

    excel_data = download_excel()
    st.sidebar.download_button(
        label="Download Result as Excel",
        data=excel_data,
        key=output_filename,
        file_name=output_filename,
    )