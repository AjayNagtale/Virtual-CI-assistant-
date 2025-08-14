        import streamlit as st
import pandas as pd
import plotly.express as px

st.title("📊 Virtual CI Specialist - Test App")

# Step 1: Upload Excel
uploaded_file = st.file_uploader("Upload your Excel file", type=["xlsx", "xls"])

if uploaded_file:
    try:
        df = pd.read_excel(uploaded_file)

        # Check required columns
        required_cols = {"Date", "Department", "Reason", "Loss Minutes"}
        if not required_cols.issubset(df.columns):
            st.error(f"Your Excel must contain these columns: {', '.join(required_cols)}")
        else:
            st.success("✅ File uploaded successfully!")
            st.dataframe(df.head())

            # Example OAE Calculation (placeholder)
            total_loss = df["Loss Minutes"].sum()
            planned_time = 24 * 60  # Example: 24 hours in minutes
            oae = ((planned_time - total_loss) / planned_time) * 100
            st.metric("OAE %", f"{oae:.2f}")

            # Sample chart
            fig = px.bar(df, x="Reason", y="Loss Minutes", title="Loss by Reason")
            st.plotly_chart(fig)

    except Exception as e:
        st.error(f"Error reading the file: {e}")
else:
    st.info("Please upload your Excel file to proceed.")
