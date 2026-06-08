import streamlit as st
import pandas as pd
import json

if 'df' not in st.session_state:
    st.session_state.df = pd.DataFrame({'TÊN ĐƠN VỊ': ['A', 'B'], 'Đạt chỉ tiêu': ['Đạt', 'Không đạt'], 'Ghi chú': ['', '']})

edited_df = st.data_editor(st.session_state.df, key='editor')
st.write(edited_df)
