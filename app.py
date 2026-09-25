import streamlit as st
import pandas as pd

st.set_page_config(page_title="BOM Production Dashboard", layout="wide")

# โหลดข้อมูล Excel
@st.cache_data
def load_data():
    df_onhand = pd.read_excel('TIKONHAND (100926 11.40).xlsx')
    df_onhand.columns = ['LPROD', 'ONHAND']

    df_drop = pd.read_excel("Drop short Movement on assy Sep'26.xlsx", sheet_name='Drop short Movement on8-30S (2)', header=None)
    df_drop_clean = df_drop.iloc[7:].copy()

    parts_df = df_drop_clean[[9, 6, 1, 5]].copy()
    parts_df.columns = ['Part', 'Requirement', 'Model', 'SO_No']
    parts_df = parts_df.dropna(subset=['Part']).copy()
    parts_df['Requirement'] = pd.to_numeric(parts_df['Requirement'], errors='coerce').fillna(0)

    merged = pd.merge(parts_df, df_onhand, left_on='Part', right_on='LPROD', how='left')
    merged['ONHAND'] = merged['ONHAND'].fillna(0)
    merged['Status'] = merged.apply(lambda r: 'Ready' if r['ONHAND'] >= r['Requirement'] else 'Shortage', axis=1)

    model_summary = merged.groupby('Model').agg({
        'Requirement': 'sum',
        'Status': lambda s: 'Shortage' if 'Shortage' in s.values else 'Ready'
    }).reset_index()
    return model_summary

df = load_data()

# Header
st.title("📊 Production Risk & Shortage Command Center")
st.markdown("Monitor multi-level BOM allocations and sequence scheduling in real-time.")

# Metrics
col1, col2, col3 = st.columns(3)
col1.metric("Total Models", len(df))
col2.metric("Shortage Models", len(df[df['Status'] == 'Shortage']))
col3.metric("Ready Models", len(df[df['Status'] == 'Ready']))

st.divider()

# Search bar
search_query = st.text_input("🔍 Search Model Name", "")

if search_query:
    filtered_df = df[df['Model'].str.contains(search_query, case=False, na=False)]
else:
    filtered_df = df

# Display Table
st.subheader("Model Production Status & Schedule")
for index, row in filtered_df.iterrows():
    col_a, col_b, col_c, col_d = st.columns([3, 2, 2, 2])
    col_a.write(f"**{row['Model']}**")
    col_b.write(f"{row['Requirement']:,.0f} EA")
    
    if row['Status'] == 'Shortage':
        col_c.error("Shortage")
    else:
        col_c.success("Ready")
        
    if col_d.button("View BOM Tree ↗", key=f"btn_{index}"):
        st.info(f"Showing BOM Tree hierarchy for: {row['Model']}")