import streamlit as st
import pandas as pd

df = pd.DataFrame({'TÊN ĐƠN VỊ': ['A', 'B'], 'Đạt chỉ tiêu': ['Đạt', 'Không đạt'], 'Ghi chú': ['', '']})
def color_dat(val):
    if val == 'Đạt':
        return 'background-color: #d4edda; color: #155724;'
    elif val == 'Không đạt':
        return 'background-color: #f8d7da; color: #721c24;'
    return ''

st.data_editor(
    df.style.map(color_dat, subset=['Đạt chỉ tiêu']), 
    column_config={
        "Đạt chỉ tiêu": st.column_config.SelectboxColumn(
            options=["Đạt", "Không đạt"]
        )
    }
)
