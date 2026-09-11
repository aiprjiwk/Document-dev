import pandas as pd
import streamlit as st

st.write("Test")

df = pd.DataFrame({"A": [1, 2, 3], "B": ["x", "y", "z"]})

def color_rows(row):
    return ['background-color: red'] * len(row) if row['A'] == 2 else [''] * len(row)

styled_df = df.style.apply(color_rows, axis=1)

edited_df = st.data_editor(styled_df, num_rows="dynamic")
st.write(edited_df)
