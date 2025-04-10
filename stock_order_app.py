import streamlit as st
import pandas as pd
from io import BytesIO
from openpyxl import Workbook



# Set page title and favicon
st.set_page_config(page_title="Stock Order App", page_icon="📦")

# Optional logo
st.image("data/logo.png", width=150)
st.title("Stock Order Generator")
# App title


# Prompt for user name
user_name = st.text_input("Please enter your full name to begin:")

if user_name:
    # Extract first name
    first_name = user_name.strip().split()[0].capitalize()
    st.success(f"Welcome, {first_name} 👋")

    # Load static files
    try:
        base_df = pd.read_excel("data/Brian stock base.xlsx")
        catalogue_df = pd.read_excel("data/CATALOGUE.xlsx")
    except Exception as e:
        st.error(f"Error loading data files: {e}")
        st.stop()

    # Upload dynamic SOH file
    soh_file = st.file_uploader("Upload SOH File", type=["xlsx"])

    # Restock calculation logic
    def calculate_order_qty(df):
        df['Available Qty'].fillna(0, inplace=True)
        df['Order Qty'] = df.apply(
            lambda row: max(row['Base Qty'] - row['Available Qty'], 0)
            if row['Available Qty'] <= row['Trigger QTY'] else 0,
            axis=1
        )
        return df[df['Order Qty'] > 0]

    # Excel export function with auto-width
    def to_excel_bytes(df):
        output = BytesIO()
        wb = Workbook()
        ws = wb.active
        ws.title = "StockOrder"

        headers = ["Quantity", "ItemCode", "ItemName", "ItemType", "Weight", "WeightUoM"]
        ws.append(headers)

        for _, row in df.iterrows():
            ws.append([
                row["Order Qty"],
                row["Item Code"],
                row["Item Name"],
                "", "", ""
            ])

        # Auto-size columns
        for col in ws.columns:
            max_length = 0
            column_letter = col[0].column_letter
            for cell in col:
                if cell.value:
                    max_length = max(max_length, len(str(cell.value)))
            ws.column_dimensions[column_letter].width = max_length + 2

        wb.save(output)
        output.seek(0)
        return output

    # Main logic after SOH is uploaded
    if soh_file:
        try:
            soh_df = pd.read_excel(soh_file)
            soh_df.columns = soh_df.columns.str.strip()  # handle trailing spaces
            merged = pd.merge(base_df, soh_df[['Item Code', 'Available Qty']], on='Item Code', how='left')
            restock_df = calculate_order_qty(merged)

            # Split into catalogued and manual pick
            in_catalogue = restock_df[restock_df['Item Code'].isin(catalogue_df['ItemCode'])]
            not_in_catalogue = restock_df[~restock_df['Item Code'].isin(catalogue_df['ItemCode'])]

            st.success("Files generated! Download below:")

            st.download_button(
                label=f"📥 Download {first_name} LM stock order.xlsx",
                data=to_excel_bytes(in_catalogue),
                file_name=f"{first_name} LM stock order.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

            st.download_button(
                label=f"📥 Download {first_name} manual pick stock.xlsx",
                data=to_excel_bytes(not_in_catalogue),
                file_name=f"{first_name} manual pick stock.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
        except Exception as e:
            st.error(f"Something went wrong: {e}")
else:
    st.info("Enter your name to begin.")
