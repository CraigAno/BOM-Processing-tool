# =============================================================================
# BOM Processing Streamlit App
# =============================================================================
# Purpose       :
#   Upload BOM / Sales Excel files, expand them by workcenter
#   using operational time data, multiply by quantity,
#   and output ONLY valid asset rows.
#
# Key Rules:
#   - Rows where Setup, Machine AND Labour are all 0/empty are removed
#   - Rows with at least one valid time are kept
#   - Multiplication by Qty is applied
#   - Remaining blanks replaced with 0
#
# Input Excel Columns (0-based index):
#   - Sales Number      : Column A (0)
#   - Ship Date         : Column C (2)
#   - Part Number (SKU) : Column D (3)
#   - Quantity          : Column F (5)
#
# Database CSV:
#   - Top_Material column identifies the Part Number
#   - Workcenter columns formatted as:
#       "61 Setup", "61 Machine", "61 Labour",
#       "62 Setup", "62 Machine", "62 Labour", etc.
#
# Output Excel Columns:
#   - Order
#   - SKU
#   - Qty
#   - Planned Ship Date
#   - Asset (Workcenter Number)
#   - Total Setup Time
#   - Total Machine Time
#   - Total Labour Time
#
# Processing Rules:
#   - Each input row may expand into multiple output rows (one per asset)
#   - Times are multiplied by quantity
#   - Parts missing from the database output a single row with zero times
#   - Parts with no operational times also output a zero-time row
#   - Same part on different dates or quantities produces separate rows
#
# Created On    : January 26, 2026
# Author        : Westeel ME Dept
# =============================================================================

import streamlit as st
import pandas as pd
import re
from io import BytesIO

# =============================================================================
# Configuration
# =============================================================================
DB_PATH = "Operational_Time_Totals_By_Top_Material_SML_Updated.csv"
# =============================================================================
# BOM Processing Logic
# =============================================================================
def process_bom(input_df: pd.DataFrame) -> pd.DataFrame:
    db_df = pd.read_csv(DB_PATH)
    db_df["Top_Material"] = db_df["Top_Material"].astype(str)

    # Detect workcenters dynamically
    workcenter_map = {}
    for col in db_df.columns:
        match = re.match(r"(\d+)\s+(Setup|Machine|Labour)", col)
        if match:
            wc, time_type = match.groups()
            workcenter_map.setdefault(wc, {})[time_type] = col

    output_rows = []

    for _, row in input_df.iterrows():
        order = row.iloc[0]
        ship_date = pd.to_datetime(row.iloc[2]).date()
        sku = str(row.iloc[3])
        qty = float(row.iloc[5])

        db_match = db_df[db_df["Top_Material"] == sku]
        if db_match.empty:
            continue  # skip missing parts

        db_row = db_match.iloc[0]

        for wc, cols in workcenter_map.items():
            def safe(v):
                try:
                    return float(v)
                except:
                    return 0.0

            # Multiply by Qty here
            setup = safe(db_row.get(cols.get("Setup"))) * qty
            machine = safe(db_row.get(cols.get("Machine"))) * qty
            labour = safe(db_row.get(cols.get("Labour"))) * qty

            # Only keep rows where at least one time > 0
            if setup == 0 and machine == 0 and labour == 0:
                continue

            output_rows.append({
                "Order": order,
                "SKU": sku,
                "Qty": qty,
                "Planned Ship Date": ship_date,
                "Asset": wc,
                "Total Setup Time": setup,
                "Total Machine Time": machine,
                "Total Labour Time": labour
            })

    return pd.DataFrame(output_rows)

# =============================================================================
# Final Cleanup
# =============================================================================
def clean_output(df: pd.DataFrame) -> pd.DataFrame:
    time_cols = ["Total Setup Time", "Total Machine Time", "Total Labour Time"]

    # Ensure numeric
    for col in time_cols:
        df[col] = pd.to_numeric(df[col], errors="coerce")

    # Drop rows where ALL time columns are 0
    df = df[~(df[time_cols].sum(axis=1) == 0)]

    # Replace remaining NaN with 0
    df[time_cols] = df[time_cols].fillna(0)

    return df.reset_index(drop=True)

# =============================================================================
# Streamlit UI
# =============================================================================
st.set_page_config(
    page_title="BOM Processing Tool",
    layout="wide",
    initial_sidebar_state="expanded"
)

st.title("📦 BOM Processing Tool")
st.markdown("""
Upload your BOM / Sales Excel file.  
Only valid asset rows with time values (after multiplying by Qty) will appear in the output.
""")

uploaded_file = st.file_uploader(
    "Upload your BOM Excel file",
    type=["xlsx", "xls"]
)

if uploaded_file:
    try:
        input_df = pd.read_excel(uploaded_file, dtype=str)
        st.success("✅ File uploaded successfully!")

        st.subheader("Preview of Uploaded Data")
        st.dataframe(input_df.head(10), use_container_width=True)

        # -----------------------------
        # Processing
        # -----------------------------
        st.info("Processing file...")
        raw_df = process_bom(input_df)

        # Clean output to remove empty assets & fill 0
        processed_df = clean_output(raw_df)

        st.success("✅ Processing complete!")

        st.subheader("Preview of Processed Data (Cleaned & Qty Applied)")
        st.dataframe(processed_df.head(20), use_container_width=True)

        # -----------------------------
        # Download
        # -----------------------------
        def to_excel(df: pd.DataFrame) -> bytes:
            output = BytesIO()
            with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
                df.to_excel(writer, index=False, sheet_name="Processed_BOM")
                worksheet = writer.sheets["Processed_BOM"]

                # Auto-fit columns based on data
                for i, col in enumerate(df.columns):
                    max_len = max(
                        df[col].astype(str).map(len).max(),
                        len(col)
                    ) + 2  # padding
                    worksheet.set_column(i, i, max_len)

            return output.getvalue()

        st.download_button(
            label="📥 Download Processed Excel",
            data=to_excel(processed_df),
            file_name="processed_bom.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

    except Exception as e:
        st.error(f"❌ Failed to process uploaded file: {e}")

else:
    st.info("Please upload an Excel file to get started.")

# =============================================================================
# Footer
# =============================================================================
st.markdown(
    """
    <style>
    .footer {
        position: fixed;
        left: 0;
        bottom: 0;
        width: 100%;
        background-color: #f1f1f1;
        color: #333;
        text-align: center;
        padding: 20px;
        font-size: 15px;
        z-index: 9999;
    }
    </style>
    <div class="footer">
        © 2026 Westeel M.E. | Contact Westeel M.E.Dept for questions or support.
    </div>
    """,
    unsafe_allow_html=True
)
