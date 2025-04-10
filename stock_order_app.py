# Step 1: Request user's full name
user_name = st.text_input("Please enter your full name to begin:")

if user_name:
    # Extract first name for filename use
    first_name = user_name.strip().split()[0].capitalize()
    st.success(f"Welcome, {first_name} 👋")

    # Load static files
    base_df = pd.read_excel("data/Brian stock base.xlsx")
    catalogue_df = pd.read_excel("data/CATALOGUE.xlsx")

    # Upload SOH file
    soh_file = st.file_uploader("Upload SOH File", type=["xlsx"])

    def calculate_order_qty(df):
        df['Available Qty'].fillna(0, inplace=True)
        df['Order Qty'] = df.apply(
            lambda row: max(row['Base Qty'] - row['Available Qty'], 0) if row['Available Qty'] <= row['Trigger QTY'] else 0,
            axis=1
        )
        return df[df['Order Qty'] > 0]

    def to_excel_bytes(df):
        output = BytesIO()
        wb = Workbook()
        ws = wb.active
        ws.title = "StockOrder"
        ws.append(["Quantity", "ItemCode", "ItemName", "ItemType", "Weight", "WeightUoM"])
        for _, row in df.iterrows():
            ws.append([
                row["Order Qty"],
                row["Item Code"],
                row["Item Name"],
                "", "", ""
            ])
        wb.save(output)
        output.seek(0)
        return output

    if soh_file:
        soh_df = pd.read_excel(soh_file)
        merged = pd.merge(base_df, soh_df[['Item Code', 'Available Qty']], on='Item Code', how='left')
        restock_df = calculate_order_qty(merged)

        in_catalogue = restock_df[restock_df['Item Code'].isin(catalogue_df['ItemCode'])]
        not_in_catalogue = restock_df[~restock_df['Item Code'].isin(catalogue_df['ItemCode'])]

        st.success("Files generated! Download below:")

        st.download_button(
            f"📥 Download {first_name} LM stock order.xlsx",
            to_excel_bytes(in_catalogue),
            file_name=f"{first_name} LM stock order.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

        st.download_button(
            f"📥 Download {first_name} manual pick stock.xlsx",
            to_excel_bytes(not_in_catalogue),
            file_name=f"{first_name} manual pick stock.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
else:
    st.info("Enter your name to begin.")
