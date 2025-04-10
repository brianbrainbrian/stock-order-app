import streamlit as st
import pandas as pd
from io import BytesIO
from openpyxl import Workbook
from pathlib import Path
import base64

# Set page title and favicon
st.set_page_config(page_title="Stock Order App", page_icon="📦")

# Convert image to base64 for embedding
def get_base64_image(image_path):
    with open(image_path, "rb") as img_file:
        return base64.b64encode(img_file.read()).decode()

# Load and embed logo if it exists
logo_path = "data/logo.png"
if Path(logo_path).exists():
    encoded_logo = get_base64_image(logo_path)
    st.markdown(
        f"""
        <style>
        .logo-container {{
            text-align: center;
            margin-bottom: 10px;
        }}
        @media (prefers-color-scheme: dark) {{
            .logo-container img {{
                filter: invert(1);
            }}
        }}
        </style>
        <div class="logo-container">
            <img src="data:image/png;base64,{encoded_logo}" width="200">
            <h1 style='text-align: center;'>Stock Order Generator</h1>
        </div>
        """,
        unsafe_allow_html=True
    )

# Initialize session state
if "user_name" not in st.session_state:
    st.session_state.user_name = ""

if "soh_file" not in st.session_state:
    st.session_state.soh_file = None

# Ask for user's name (if not set)
if st.session_state.user_name == "":
    name_input = st.text_input("Please enter your full name to begin:")
    if name_input:
        st.session_state.user_name = name_input

if st.session_state.user_name:
    first_name = st.session_state.user_name.strip().split()[0].capitalize()
    st.success(f"Welcome, {first_name} 👋")

    # Load static files
    try:
        base_df = pd.read_excel("data/Brian stock base.xlsx")
        catalogue_df = pd.read_excel("data/CATALOGUE.xlsx")
    except Exception as e:
        st.error(f"Error loading data files: {e}")
        st.stop()

    # Upload SOH file
    if st.session_state.soh_file is None:
        soh_file = st.file_uploader("Upload SOH File", type=["xlsx"])
        if soh_file:
            st.session_state.soh_file = soh_file

    # Proceed if file has been uploaded
    if st.session_state.soh_file:
        try:
            soh_df = pd.read_excel(st.session_state.soh_file)
            soh_df.columns = soh_df.columns.str.strip()

            merged = pd.merge(base_df, soh_df[['Item Code', 'Available Qty']], on='Item Code', how='left')

            def calculate_order_qty(df):
                df['Available Qty'].fillna(0, inplace=True)
                df['Order Qty'] = df.apply(
                    lambda row: max(row['Base Qty'] - row['Available Qty'], 0)
                    if row['Available Qty'] <= row['Trigger QTY'] else 0,
                    axis=1
                )
                return df[df['Order Qty'] > 0]

            restock_df = calculate_order_qty(merged)

            in_catalogue = restock_df[restock_df['Item Code'].isin(catalogue_df['ItemCode'])]
            not_in_catalogue = restock_df[~restock_df['Item Code'].isin(catalogue_df['ItemCode'])]

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

    # Divider and reset button
    st.markdown("---")
    if st.button("🔁 Start Again"):
        st.session_state.user_name = ""
        st.session_state.soh_file = None
        st.experimental_rerun()

else:
    st.info("Enter your name to begin.")
