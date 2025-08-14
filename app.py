import streamlit as st
import pandas as pd

st.set_page_config(page_title="Virtual CI Assistant", layout="wide")

# Title
st.title("📊 Virtual CI Assistant - OAE & OEE Analysis")

# File uploader
uploaded_file = st.file_uploader("Upload your Excel file", type=["xlsx", "xls"])

if uploaded_file:
    try:
        # Read Excel
        df = pd.read_excel(uploaded_file)

        # Normalize column names (lowercase + strip spaces)
        df.columns = df.columns.str.strip().str.lower()

        # Required columns
        required_columns = ["department", "date", "loss minutes", "reason"]

        # Check if all required columns are present
        if not all(col.lower() in df.columns for col in required_columns):
            st.error(f"Your Excel must contain these columns: {', '.join(required_columns)}")
        else:
            # Clean data
            df_clean = df.copy()
            df_clean["department"] = df_clean["department"].astype(str).str.strip()
            df_clean["reason"] = df_clean["reason"].astype(str).str.strip()
            df_clean["loss minutes"] = pd.to_numeric(df_clean["loss minutes"], errors="coerce").fillna(0)
            df_clean["date"] = pd.to_datetime(df_clean["date"], errors="coerce")

            # Drop rows with invalid dates
            df_clean = df_clean.dropna(subset=["date"])

            # Basic OAE calculation example (you can extend later)
            total_loss = df_clean["loss minutes"].sum()
            planned_time = len(df_clean) * 480  # assuming 480 min/day per record
            oae = ((planned_time - total_loss) / planned_time) * 100 if planned_time > 0 else 0

            st.success("✅ File processed successfully!")
            st.metric(label="OAE %", value=f"{oae:.2f}")

            # Show cleaned data
            st.subheader("Cleaned Data Preview")
            st.dataframe(df_clean)

    except Exception as e:
        st.error(f"Error processing file: {e}")

else:
    st.info("Please upload an Excel file to proceed.")
                    
