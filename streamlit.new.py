import pandas as pd
from rapidfuzz import fuzz as rapid_fuzz
import jellyfish
import streamlit as st
import io

# Ensure you have xlsxwriter installed
# You can do this by running: pip install xlsxwriter

# Define the function for matching products
def match_products(basket_df, master_df, threshold=65):
    matched_basket = []
    for index, row in basket_df.iterrows():
        basket_desc = row["Product Description"]
        best_match_score = 0
        best_match_product = None
        basket_metaphone = jellyfish.metaphone(basket_desc)

        for _, master_row in master_df.iterrows():
            master_desc = master_row["Product Description"]
            master_metaphone = jellyfish.metaphone(master_desc)

            # Calculate different scores
            rapid_fuzz_ratio = rapid_fuzz.ratio(basket_desc, master_desc)
            metaphone_score = 100 * (basket_metaphone == master_metaphone)

            # Taking the max of all scores
            max_score = max(rapid_fuzz_ratio, metaphone_score)

            # Check if the maximum score is above threshold
            if max_score > threshold:
                best_match_score = max_score
                best_match_product = master_row

        if best_match_score >= threshold:  # Threshold for fuzzy matching
            matched_basket.append({**row.to_dict(), **best_match_product.to_dict()})
        else:
            matched_basket.append(row.to_dict())  # Add unmatched item

    return pd.DataFrame(matched_basket)

# Streamlit interface
st.title("Product Matching Tool")

# Slider for threshold selection
threshold = st.slider("Select Matching Threshold", min_value=0, max_value=99, value=80, step=1)

# Text input for sheet name
sheet_name = st.text_input("Enter the sheet to match", "")

# File upload
basket_file = st.file_uploader("Upload Basket Excel File", type=["xlsx"])
master_file = st.file_uploader("Upload Master CSV File (Contract list)", type=["csv"])

# Text input for output file name
output_file_name = st.text_input("Enter the file name for the download (without extension)", "")

# Button to run the matching process
if st.button("Run Matching Process"):
    if basket_file and master_file:
        # Read the uploaded files
        try:
            basket_df = pd.read_excel(basket_file, sheet_name=sheet_name)
        except Exception as e:
            st.error(f"Error reading sheet '{sheet_name}' from the basket file: {e}")
            st.stop()

        master_df = pd.read_csv(master_file, encoding='latin-1')

        # Perform product matching
        matched_basket_df = match_products(basket_df, master_df, threshold)

        # Display the matched DataFrame
        st.write("Matched Products")
        st.dataframe(matched_basket_df)

        # Save the matched DataFrame to an Excel file in memory
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
            matched_basket_df.to_excel(writer, index=False, sheet_name="Matched Output")
        output.seek(0)

        # Download button
        st.download_button(
            label="Download Matched Results",
            data=output,
            file_name=f"{output_file_name}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
    else:
        st.error("Please upload both the basket Excel file and the master CSV file to perform matching.")
