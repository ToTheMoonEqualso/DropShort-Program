import pandas as pd
import streamlit as st
import os
import json
import html
import re
import numpy as np
from datetime import date, timedelta

st.set_page_config(page_title="DropShot Visualizer - Multi-Level Allocation", layout="wide")

MASTER_BOM_PATH = "Usage & Level.xlsx"

def load_html_template():
    if os.path.exists("index.html"):
        with open("index.html", "r", encoding="utf-8") as f:
            return f.read(), True
    return "<h3>Error: ไม่พบไฟล์ index.html ในโฟลเดอร์นี้</h3>", False

@st.cache_data
def load_master_bom():
    if not os.path.exists(MASTER_BOM_PATH):
        st.error(f"⚠️ ไม่พบไฟล์ Master BOM `{MASTER_BOM_PATH}` ในโฟลเดอร์โปรเจกต์ กรุณานำไฟล์มาวางคู่กับ BOMCheck.py")
        return None
    
    xls = pd.ExcelFile(MASTER_BOM_PATH)
    df_master = pd.read_excel(xls, sheet_name=0)
    df_master.columns = df_master.columns.astype(str).str.strip()
    
    df_master['Item'] = df_master['Item'].astype(str).str.strip()
    df_master['Description'] = df_master['Description'].astype(str).str.strip()
    df_master['Level_Str'] = df_master['Level'].astype(str).str.strip()
    df_master['Level_Num'] = df_master['Level_Str'].apply(lambda x: x.count('.') if x != 'nan' else 0)
    
    if 'Usage' in df_master.columns:
        df_master['Usage'] = pd.to_numeric(df_master['Usage'], errors='coerce').fillna(1.0)
    else:
        df_master['Usage'] = 1.0
        
    return df_master

@st.cache_data
def load_and_clean_movement(uploaded_file):
    try:
        xls = pd.ExcelFile(uploaded_file, engine='xlrd')
    except Exception:
        xls = pd.ExcelFile(uploaded_file)
        
    if 'Drop short Movement on 19-3 (2)' in xls.sheet_names:
        df = pd.read_excel(xls, sheet_name='Drop short Movement on 19-3 (2)')
        df.columns = df.columns.astype(str).str.strip()
        df_clean = df.dropna(subset=['Part']).copy()
        df_clean['Part'] = df_clean['Part'].astype(str).str.strip()
        df_clean['Require'] = pd.to_numeric(df_clean['Require'], errors='coerce').fillna(0.0)
        df_clean['On hand'] = pd.to_numeric(df_clean['On hand'], errors='coerce').fillna(0.0)
        df_clean['Shortage'] = pd.to_numeric(df_clean['Shortage'], errors='coerce').fillna(0.0)
        return df_clean

    df_raw = pd.read_excel(xls, sheet_name=0)
    raw_lines = df_raw.iloc[:, 0].astype(str).tolist()
    
    parsed_data = []
    current_model = ""
    current_req = 0.0
    
    for line in raw_lines:
        line_str = line.strip()
        if not line_str or '====' in line_str or 'Page :' in line_str or 'Part Praparation Order' in line_str:
            continue
            
        if re.match(r'^[A-Z0-9]{2}\s+', line_str):
            parts = line_str.split()
            if len(parts) >= 7:
                current_model = parts[1] + " " + parts[2] if len(parts) > 2 else parts[1]
                try:
                    current_req = float(parts[6].replace(',', ''))
                except ValueError:
                    current_req = 0.0
        elif '|' in line_str:
            sub_parts = line_str.split('|')
            if len(sub_parts) >= 2:
                content = sub_parts[-1].strip()
                tokens = content.split()
                if len(tokens) >= 3:
                    part_code = tokens[0]
                    try:
                        shortage_qty = float(tokens[-2].replace(',', ''))
                    except ValueError:
                        shortage_qty = 0.0
                    
                    desc = " ".join(tokens[1:-3]) if len(tokens) > 4 else tokens[1]
                    
                    parsed_data.append({
                        'Model Description': current_model,
                        'Part': part_code,
                        'Description': desc,
                        'Require': current_req,
                        'On hand': 0.0,
                        'Shortage': shortage_qty
                    })
                    
    return pd.DataFrame(parsed_data)

def extract_category(description):
    desc = str(description).strip()
    categories = ['FGD', 'INJ', 'PBA', 'IMP', 'BOI', 'SCB', 'SUC', 'DOM', 'CUW', 'SMT', 'MVA']
    desc_upper = desc.upper()
    for cat in categories:
        if desc_upper.endswith(cat):
            return cat
    tokens = desc.split()
    if tokens:
        last = tokens[-1].upper()
        if len(last) in [3, 4]:
            return last
    return "OTHER"

def build_tree_structure(df_model_bom, stock_map, target_order_qty):
    stack = []
    root_nodes = []
    
    for idx, row in df_model_bom.iterrows():
        lvl = int(row['Level_Num'])
        usage = float(row.get('Usage', 1.0))
        item = str(row['Item']).strip()
        desc = str(row['Description']).strip()
        um = str(row.get('UM', 'EA')).strip()
        cat = extract_category(desc)
        
        node = {
            'Node_ID': f"node_{idx}",
            'Level': lvl,
            'Level_Str': str(row['Level']),
            'Item': item,
            'Description': desc,
            'Category': cat,
            'Usage': usage,
            'UM': um,
            'Stock_Qty': 0.0,
            'Total_Required': 0.0,
            'Shortage_Qty': 0.0,
            'Is_Shortage': False,
            'Has_Child_Shortage': False,
            'Not_Needed': False,
            'Children': []
        }
        
        while stack and stack[-1]['Level'] >= lvl:
            stack.pop()
            
        if stack:
            stack[-1]['Children'].append(node)
        else:
            root_nodes.append(node)
            
        stack.append(node)

    def calculate_requirements(node, parent_needed_qty, parent_is_not_needed):
        lvl = node['Level']
        item = node['Item']
        usage = node['Usage']
        
        if lvl == 0:
            node['Total_Required'] = target_order_qty
            node['Not_Needed'] = False
            node['Is_Shortage'] = False
            node['Shortage_Qty'] = 0.0
            node['Stock_Qty'] = 0.0
            
            qty_to_pass_down = target_order_qty
            child_not_needed = False
        else:
            if parent_is_not_needed or parent_needed_qty <= 0:
                node['Not_Needed'] = True
                node['Total_Required'] = 0.0
                node['Stock_Qty'] = stock_map.get(item, {}).get('On_Hand', 0.0)
                node['Shortage_Qty'] = 0.0
                node['Is_Shortage'] = False
                qty_to_pass_down = 0.0
                child_not_needed = True
            else:
                movement_info = stock_map.get(item, {'On_Hand': 0.0, 'Shortage': 0.0})
                current_stock = movement_info['On_Hand']
                required_qty = usage * parent_needed_qty
                
                node['Total_Required'] = required_qty
                node['Stock_Qty'] = current_stock
                node['Not_Needed'] = False
                
                if current_stock >= required_qty:
                    node['Shortage_Qty'] = 0.0
                    node['Is_Shortage'] = False
                    qty_to_pass_down = 0.0
                    child_not_needed = True
                else:
                    needed_more = required_qty - current_stock
                    node['Shortage_Qty'] = needed_more
                    node['Is_Shortage'] = True
                    qty_to_pass_down = needed_more
                    child_not_needed = False

        for child in node['Children']:
            calculate_requirements(child, qty_to_pass_down, child_not_needed)

    def mark_child_shortages(n):
        has_sub_shortage = False
        for child in n['Children']:
            child_has_shortage = mark_child_shortages(child)
            if child['Is_Shortage'] or child_has_shortage:
                has_sub_shortage = True
        n['Has_Child_Shortage'] = has_sub_shortage
        return has_sub_shortage

    for root in root_nodes:
        calculate_requirements(root, target_order_qty, False)
        mark_child_shortages(root)
        
    return root_nodes

def build_echarts_json(nodes):
    def _convert_node(n):
        item_code = str(n.get("Item", "")).strip()
        description = str(n.get("Description", "")).strip()
        level_num = n.get("Level", 0)
        um = str(n.get("UM", "pc")).strip()
        usage = n.get("Usage", 1.0)
        req_qty = n.get("Total_Required", 0.0)
        cat = n.get("Category", "OTHER")
        not_needed = n.get("Not_Needed", False)
        
        is_shortage = n.get("Is_Shortage", False)
        has_child_shortage = n.get("Has_Child_Shortage", False)
        
        if not_needed:
            status = "NOT_NEEDED"
        elif is_shortage:
            status = "SHORTAGE"
        elif has_child_shortage:
            status = "CHILD_SHORTAGE"
        else:
            status = "OK"
            
        children = [_convert_node(child) for child in n.get("Children", [])]
        usage_str = f"{usage:,.0f}" if usage.is_integer() else f"{usage:,.4f}"
        qty_str = f"{req_qty:,.0f} {um}" if req_qty.is_integer() else f"{req_qty:,.2f} {um}"
        
        return {
            "code": item_code,
            "desc": description,
            "category": cat,
            "level": str(level_num),
            "req_base": usage_str,
            "qty": qty_str,
            "status": status,
            "not_needed": not_needed,
            "shortage_qty": n.get("Shortage_Qty", 0.0),
            "stock_qty": n.get("Stock_Qty", 0.0),
            "children": children
        }

    if isinstance(nodes, list):
        if len(nodes) == 1:
            return _convert_node(nodes[0])
        else:
            return {
                "code": "ROOT",
                "desc": "Main Model Breakdown",
                "category": "ALL",
                "level": "0",
                "req_base": "1",
                "qty": "",
                "status": "CHILD_SHORTAGE" if any(n.get("Is_Shortage") or n.get("Has_Child_Shortage") for n in nodes) else "OK",
                "children": [_convert_node(r) for r in nodes]
            }
    return _convert_node(nodes)

def get_all_parts_df(nodes):
    flat = []
    def _traverse(n_list):
        for n in n_list:
            flat.append({
                'Level (ชั้น)': n['Level'],
                'Item Code': n['Item'],
                'Description': n['Description'],
                'Category': n['Category'],
                'Usage/Unit (ต่อหน่วย)': n['Usage'],
                'Stock ในคลังจริง (On Hand)': n['Stock_Qty'],
                'ต้องใช้ทั้งหมด (Total Req)': n['Total_Required'],
                'สถานะ Drop-shot': 'ขาด (Shortage)' if n['Is_Shortage'] else 'พอ (Sufficient)',
                'จำนวนที่ขาด': n['Shortage_Qty'] if n['Is_Shortage'] else 0.0,
                'หน่วย': n['UM']
            })
            _traverse(n['Children'])
    _traverse(nodes)
    return pd.DataFrame(flat)

def check_model_overall_shortage(df_master_bom, stock_map, model_name, default_target=500.0):
    df_model = df_master_bom[df_master_bom['Model'] == model_name].copy()
    tree_nodes = build_tree_structure(df_model, stock_map, default_target)
    df_all_parts = get_all_parts_df(tree_nodes)
    shortage_parts = df_all_parts[df_all_parts['สถานะ Drop-shot'] == 'ขาด (Shortage)']
    return len(shortage_parts) > 0, len(shortage_parts)

# --- STREAMLIT UI ---
st.title("🎯 DropShot Visualizer — Multi-Level Shortage & Stock Allocation System")

df_master_bom = load_master_bom()

if df_master_bom is not None:
    st.sidebar.header("📁 เมนูหลัก")
    
    use_mock_stock = st.sidebar.checkbox("🧪 โหมดทดสอบ: สมมติค่า Stock On-Hand เอง", value=False)
    
    uploaded_file = None
    if not use_mock_stock:
        uploaded_file = st.sidebar.file_uploader("📂 อัปโหลดไฟล์ Drop short Movement (.xls / .xlsx)", type=['xls', 'xlsx'])
    
    if uploaded_file or use_mock_stock:
        stock_map = {}
        
        if use_mock_stock:
            st.sidebar.subheader("⚙️ ตั้งค่า Mock Stock")
            mock_mode = st.sidebar.radio("รูปแบบการจำลอง:", ["🎲 สุ่มสต็อกหลายระดับอัตโนมัติ (Random)", "✏️ กำหนด On-Hand แต่ละชิ้นส่วนเอง (Manual Table)"])
            
            all_items = df_master_bom['Item'].unique()
            
            if mock_mode == "🎲 สุ่มสต็อกหลายระดับอัตโนมัติ (Random)":
                seed_val = st.sidebar.number_input("Random Seed (สุ่มรูปเดิม):", min_value=1, value=42)
                np.random.seed(seed_val)
                
                for item in all_items:
                    rand_type = np.random.choice(["SUFFICIENT", "SHORTAGE", "ZERO"], p=[0.5, 0.3, 0.2])
                    if rand_type == "SUFFICIENT":
                        stock_map[item] = {'On_Hand': float(np.random.randint(1000, 10000)), 'Shortage': 0.0}
                    elif rand_type == "SHORTAGE":
                        stock_map[item] = {'On_Hand': float(np.random.randint(10, 300)), 'Shortage': 100.0}
                    else:
                        stock_map[item] = {'On_Hand': 0.0, 'Shortage': 500.0}
                        
            else:
                st.info("💡 สามารถกดดับเบิ้ลคลิกเพื่อแก้ไขตัวเลขในคอลัมน์ **On_Hand_Stock** ทางด้านล่างได้ทันทีครับ")
                if 'mock_stock_df' not in st.session_state:
                    sample_df = pd.DataFrame({'Part': all_items, 'On_Hand_Stock': 500.0})
                    st.session_state['mock_stock_df'] = sample_df
                    
                edited_df = st.data_editor(
                    st.session_state['mock_stock_df'],
                    column_config={
                        "Part": "รหัสชิ้นส่วน (Item)",
                        "On_Hand_Stock": st.column_config.NumberColumn("ยอด On-Hand สมมติ", min_value=0.0, step=10.0)
                    },
                    use_container_width=True,
                    height=200
                )
                
                for _, r in edited_df.iterrows():
                    stock_map[str(r['Part']).strip()] = {'On_Hand': float(r['On_Hand_Stock']), 'Shortage': 0.0}

        else:
            df_movement = load_and_clean_movement(uploaded_file)
            if 'Part' in df_movement.columns:
                for _, r in df_movement.iterrows():
                    stock_map[str(r['Part']).strip()] = {
                        'On_Hand': float(r.get('On hand', 0.0)),
                        'Shortage': float(r.get('Shortage', 0.0))
                    }

        df_lvl0 = df_master_bom[df_master_bom['Level_Num'] == 0].copy()
        models_lvl0 = df_lvl0['Model'].unique()
        
        if 'selected_model' not in st.session_state or st.session_state['selected_model'] not in models_lvl0:
            st.session_state['selected_model'] = models_lvl0[0]

        st.sidebar.divider()
        st.sidebar.subheader("🎯 เลือก Model")
        selected_model_input = st.sidebar.selectbox(
            "เลือก Model (Level 0) ที่ต้องการดูรายละเอียด:",
            models_lvl0,
            index=list(models_lvl0).index(st.session_state['selected_model'])
        )
        st.session_state['selected_model'] = selected_model_input
        
        target_qty_input = st.sidebar.number_input("เป้าหมายการผลิต (Target Order Qty):", min_value=1.0, value=500.0, step=50.0)

        tab1, tab2, tab3 = st.tabs([
            "📋 รายการ Model ทั้งหมด (Level 0 Overview)", 
            "🧠 รายละเอียดและโครงสร้าง Drop-down (Model Hierarchy View)",
            "📊 สรุปรวมชิ้นส่วนซ้ำทุก Model (Consolidated View)"
        ])
        
        # --- TAB 1: Level 0 Overview ---
        with tab1:
            st.subheader("📋 รายการ Model และวิเคราะห์ชิ้นส่วนช็อตตามลำดับการผลิต")
            
            # --- ปุ่มสุ่มวันที่สำหรับทดสอบตัดสต็อกสะสมตามลำดับวัน ---
            col_date_a, col_date_b = st.columns([3, 1])
            with col_date_b:
                if st.button("🎲 สุ่มกระจายวันที่ต้องการผลิต"):
                    np.random.seed(None)
                    st.session_state['model_plan_dates'] = {
                        m: date.today() + timedelta(days=int(np.random.randint(0, 5))) 
                        for m in models_lvl0
                    }
                    st.rerun()

            if 'model_plan_dates' not in st.session_state:
                np.random.seed(42)
                st.session_state['model_plan_dates'] = {
                    m: date.today() + timedelta(days=int(np.random.randint(0, 5))) 
                    for m in models_lvl0
                }

            # เรียงคิวการคำนวณตัดสต็อกสะสมตามลำดับ "วันที่ต้องการผลิต" (FIFO Sequential Depletion)
            sorted_models_by_date = sorted(models_lvl0, key=lambda x: st.session_state['model_plan_dates'].get(x, date.today()))
            
            running_stock_map = {k: v.copy() for k, v in stock_map.items()}
            model_status_list = []
            detailed_shortage_records = []
            
            # วนลูปตัดสต็อกสะสมจริงตามลำดับคิวผลิต
            for seq_idx, m in enumerate(sorted_models_by_date, start=1):
                m_date = st.session_state['model_plan_dates'].get(m, date.today())
                
                # สร้าง Tree และหาชิ้นส่วนที่ช็อตของ Model นี้โดยอิงจากสต็อกที่เหลือสะสม (running_stock_map)
                df_m_bom = df_master_bom[df_master_bom['Model'] == m]
                tree_nodes = build_tree_structure(df_m_bom, running_stock_map, target_qty_input)
                df_all_parts = get_all_parts_df(tree_nodes)
                df_short_parts = df_all_parts[df_all_parts['สถานะ Drop-shot'] == 'ขาด (Shortage)']
                
                # บันทึกข้อมูลชิ้นส่วนที่ช็อตสำหรับตารางวิเคราะห์
                for _, p_row in df_short_parts.iterrows():
                    detailed_shortage_records.append({
                        'ลำดับผลิต': seq_idx,
                        'Model': m,
                        'วันที่ผลิต': m_date,
                        'Part Code': p_row['Item Code'],
                        'ชื่อสิ่งที่ช็อต (Description)': p_row['Description'],
                        'Requirement': p_row['ต้องใช้ทั้งหมด (Total Req)'],
                        'On-Hand': p_row['Stock ในคลังจริง (On Hand)'],
                        'Shorted': p_row['จำนวนที่ขาด'],
                        'หน่วย': p_row['หน่วย']
                    })
                
                model_status_list.append({
                    'Model': m,
                    'Has_Shortage': len(df_short_parts) > 0,
                    'Shortage_Count': len(df_short_parts)
                })
                
                # ตัดสต็อกสะสมจริงตาม BOM สำหรับคิวถัดไป
                for _, r in df_m_bom.iterrows():
                    p_item = str(r['Item']).strip()
                    p_usage = float(r.get('Usage', 1.0))
                    needed = p_usage * target_qty_input
                    if p_item in running_stock_map:
                        running_stock_map[p_item]['On_Hand'] = max(0.0, running_stock_map[p_item]['On_Hand'] - needed)

            df_status_map = pd.DataFrame(model_status_list)
            shortage_models_cnt = df_status_map['Has_Shortage'].sum()
            sufficient_models_cnt = len(models_lvl0) - shortage_models_cnt
            
            # --- Metric Summary Cards ---
            col1, col2, col3 = st.columns(3)
            col1.metric("จำนวน Model ทั้งหมด", f"{len(models_lvl0)} รุ่น")
            col2.metric("Model ที่ขาดชิ้นส่วน (ทุกชั้น)", f"{shortage_models_cnt} รุ่น", delta_color="inverse")
            col3.metric("Model ที่ชิ้นส่วนพอครบทุกชั้น", f"{sufficient_models_cnt} รุ่น")
            
            st.divider()

            # --- สร้าง Sub-tabs ย่อยใน Tab 1 ---
            tab1_sub1, tab1_sub2 = st.tabs([
                "📋 ภาพรวมรายการ Model ทั้งหมด", 
                "🚨 วิเคราะห์ชิ้นส่วนช็อตซ้ำตามคิวผลิต (Sequential Shortage)"
            ])

            # --- Sub-tab 1: ตารางสรุปภาพรวม Model ---
            with tab1_sub1:
                st.info(f"📌 Model ที่เลือกดูในปัจจุบัน: **{st.session_state['selected_model']}** (เปลี่ยน Model ได้จากแถบ Sidebar ด้านซ้าย)")
                
                lvl0_display = df_lvl0[['Model', 'Description', 'Usage']].copy()
                lvl0_display['Requirement'] = lvl0_display['Usage'] * target_qty_input
                lvl0_display = lvl0_display.merge(df_status_map, on='Model', how='left')
                
                lvl0_display['MODEL NAME'] = lvl0_display['Model'] + "\n" + lvl0_display['Description']
                lvl0_display['วันที่ต้องการผลิต'] = lvl0_display['Model'].map(st.session_state['model_plan_dates'])
                lvl0_display['SHORTAGE'] = lvl0_display.apply(
                    lambda r: f"🔴 ขาด ({r['Shortage_Count']} รายการ)" if r['Has_Shortage'] else "🟢 พร้อมผลิต (0 รายการ)", axis=1
                )
                lvl0_display = lvl0_display.sort_values(by='วันที่ต้องการผลิต')
                
                st.markdown("#### 📄 ตารางสรุปรายการ Model ทั้งหมด (คำนวณตัดสต็อกสะสมตามลำดับวันที่ผลิต)")
                
                st.data_editor(
                    lvl0_display[['MODEL NAME', 'วันที่ต้องการผลิต', 'Requirement', 'SHORTAGE']],
                    column_config={
                        "MODEL NAME": st.column_config.TextColumn("ชื่อโมเดล (MODEL NAME)"),
                        "วันที่ต้องการผลิต": st.column_config.DateColumn("วันที่ต้องการผลิต", format="YYYY-MM-DD"),
                        "Requirement": st.column_config.NumberColumn("ความต้องการ (REQUIREMENT)", format="%d EA"),
                        "SHORTAGE": st.column_config.TextColumn("รายการที่ช็อต (SHORTAGE)")
                    },
                    use_container_width=True,
                    hide_index=True,
                    disabled=["MODEL NAME", "วันที่ต้องการผลิต", "Requirement", "SHORTAGE"]
                )

            # --- Sub-tab 2: รายการชิ้นส่วนช็อตซ้ำ เรียงตามคิวผลิต ---
            with tab1_sub2:
                st.markdown("#### 🚨 รายการชิ้นส่วนที่ช็อตซ้ำกันข้าม Model (แยกประมวลผลเฉพาะภายในวันผลิตเดียวกัน)")
                st.caption("ตารางนี้คำนวณตัดสต็อกสะสมเฉพาะ **Model ที่ผลิตในวันเดียวกัน** ( On-Hand ถูกอัปเดตใหม่ทุกวัน ไม่ดึงสต็อกตัดข้ามวัน)")

                if detailed_shortage_records:
                    # แปลงข้อมูลเป็น DataFrame
                    df_all_records = pd.DataFrame(detailed_shortage_records)
                    
                    # 1. จัดกลุ่มประมวลผลหาชิ้นส่วนที่ช็อตซ้ำ "ภายในวันผลิตเดียวกัน" (Same Date Repeated Parts)
                    df_all_records['Same_Date_Part_Count'] = df_all_records.groupby(['วันที่ผลิต', 'Part Code'])['Model'].transform('count')
                    
                    filter_mode = st.radio(
                        "🔍 รูปแบบการกรองรายการช็อต:",
                        ["แสดงเฉพาะชิ้นส่วนที่ช็อตซ้ำกันในวันเดียวกัน (Repeated Shortage on Same Date)", "แสดงชิ้นส่วนที่ช็อตทั้งหมดแยกตามวัน (All Shortage Items by Date)"],
                        horizontal=True
                    )

                    if filter_mode == "แสดงเฉพาะชิ้นส่วนที่ช็อตซ้ำกันในวันเดียวกัน (Repeated Shortage on Same Date)":
                        df_display_final = df_all_records[df_all_records['Same_Date_Part_Count'] > 1].copy()
                        if df_display_final.empty:
                            st.info("💡 ไม่พบชิ้นส่วนช็อตที่ใช้ซ้ำกันข้าม Model ในวันผลิตเดียวกัน")
                    else:
                        df_display_final = df_all_records.copy()

                    if not df_display_final.empty:
                        # 2. จัดกลุ่มแสดงผลแยกตาม "วันที่ผลิต" ก่อน แล้วค่อยกลุ่มตาม "ชิ้นส่วน"
                        grouped_by_date = df_display_final.groupby('วันที่ผลิต')

                        for prod_date, date_group_df in grouped_by_date:
                            st.markdown(f"##### 📅 วันที่ผลิต: **{prod_date.strftime('%Y-%m-%d')}**")
                            
                            grouped_by_part = date_group_df.groupby(['Part Code', 'ชื่อสิ่งที่ช็อต (Description)', 'หน่วย'])

                            for (p_code, p_desc, p_unit), part_group_df in grouped_by_part:
                                total_short = part_group_df['Shorted'].sum()
                                model_cnt = part_group_df['Model'].nunique()
                                
                                # ใช้ st.container(border=True) สร้างการ์ดขอบมน มินิมอล
                                with st.container(border=True):
                                    card_html = f"""
                                    <div style="display: flex; align-items: center; justify-content: space-between; width: 100%; padding: 4px 0;">
                                        <div style="display: flex; align-items: center; gap: 24px; flex-wrap: nowrap; overflow: hidden;">
                                            <div style="font-size: 16px; font-weight: 800; color: #f8fafc; white-space: nowrap; min-width: 170px;">
                                                {p_code}
                                            </div>
                                            <div style="font-size: 14px; font-weight: 700; color: #e2e8f0; white-space: nowrap; overflow: hidden; text-overflow: ellipsis;">
                                                {p_desc}
                                            </div>
                                            <div style="font-size: 12.5px; color: #94a3b8; white-space: nowrap;">
                                                กระทบในวันนี้: <b style="color: #f1f5f9;">{model_cnt} Models</b> | ยอดช็อตรวมวันนี้: 
                                                <span style="color: #dc2626; font-weight: 800; border: 1.5px solid #ef4444; padding: 1px 6px; border-radius: 4px; background: transparent;">{total_short:,.2f} {p_unit}</span>
                                            </div>
                                        </div>
                                        <div style="color: #94a3b8; font-size: 18px; padding-left: 12px;">&#8594;</div>
                                    </div>
                                    """
                                    st.markdown(card_html, unsafe_allow_html=True)
                                    
                                    # ตารางแสดงคิวผลิตย่อยเฉพาะของวันนั้นๆ
                                    with st.expander(f"🔍 ดูรายละเอียดคิวผลิตวันที่ {prod_date.strftime('%Y-%m-%d')} ของชิ้นส่วนนี้"):
                                        st.dataframe(
                                            part_group_df[[
                                                'ลำดับผลิต', 'Model', 'Requirement', 'On-Hand', 'Shorted'
                                            ]].style.format({
                                                'Requirement': '{:,.2f}',
                                                'On-Hand': '{:,.2f}',
                                                'Shorted': '{:,.2f}'
                                            }).map(
                                                lambda x: 'color: #dc2626; font-weight: 800; border: 1.5px solid #ef4444;',
                                                subset=['Shorted']
                                            ),
                                            use_container_width=True,
                                            hide_index=True,
                                            column_config={
                                                "ลำดับผลิต": st.column_config.NumberColumn("คิวที่ (Seq)", format="#%d"),
                                                "Model": st.column_config.TextColumn("Model"),
                                                "Requirement": st.column_config.NumberColumn("Requirement"),
                                                "On-Hand": st.column_config.NumberColumn("On-Hand สต็อกวันนั้น"),
                                                "Shorted": st.column_config.NumberColumn("ยอด Shorted")
                                            }
                                        )
                            st.divider()
                else:
                    st.success("🎉 สมบูรณ์แบบ! ไม่พบรายการชิ้นส่วนที่ช็อตในการวางแผนผลิตรอบนี้")

        # --- TAB 2: Model Hierarchy View ---
        with tab2:
            selected_model = st.session_state['selected_model']
            st.subheader(f"🎯 รายละเอียดและโครงสร้าง Drop-Down: Model **{selected_model}**")
            
            df_model_bom = df_master_bom[df_master_bom['Model'] == selected_model].copy()
            
            tree_nodes = build_tree_structure(df_model_bom, stock_map, target_qty_input)
            df_all_parts = get_all_parts_df(tree_nodes)
            df_shortage_only = df_all_parts[df_all_parts['สถานะ Drop-shot'] == 'ขาด (Shortage)']
            
            col_a, col_b = st.columns(2)
            if len(df_shortage_only) > 0:
                col_a.error(f"⚠️ พบชิ้นส่วนไม่เพียงพอ (Shortage) ทั้งหมด {len(df_shortage_only)} รายการ")
            else:
                col_a.success("🎉 ชิ้นส่วนครบถ้วน พร้อมสำหรับการประกอบชิ้นงาน (Sufficient)")
                
            col_b.info(f"📦 จำนวนรายการชิ้นส่วนทั้งหมดใน Model นี้: {len(df_all_parts)} รายการ")
            
            # --- สร้าง Sub-tabs ย่อยใน Tab 2 ---
            sub_tab1, sub_tab2 = st.tabs([
                "🌳 รายการโครงสร้างทั้งหมด (All Hierarchy Tree)", 
                "🚨 เฉพาะชิ้นส่วนที่ช็อต (Shortage Only)"
            ])

            with sub_tab1:
                html_template, is_success = load_html_template()
                if not is_success:
                    st.error(html_template)
                else:
                    tree_data = build_echarts_json(tree_nodes)
                    shortage_count = len(df_shortage_only)
                    shortage_color = "#dc2626" if shortage_count > 0 else "#16a34a"
                    
                    rendered_html = html_template.replace("{{MODEL_NAME}}", html.escape(str(selected_model)))\
                                                 .replace("{{SHORTAGE_COUNT}}", str(shortage_count))\
                                                 .replace("{{SHORTAGE_COLOR}}", shortage_color)\
                                                 .replace("{{TARGET_QTY}}", f"{target_qty_input:,.0f}")\
                                                 .replace("{{TREE_DATA_JSON}}", json.dumps(tree_data, ensure_ascii=False))
                    
                    st.components.v1.html(rendered_html, height=850, scrolling=True)

            with sub_tab2:
                st.markdown("### 🚨 รายการชิ้นส่วนที่ไม่เพียงพอต่อการผลิต (Shortage List)")
                if len(df_shortage_only) > 0:
                    df_short_display = df_shortage_only[[
                        'Level (ชั้น)', 'Item Code', 'Description', 'Category', 
                        'Usage/Unit (ต่อหน่วย)', 'Stock ในคลังจริง (On Hand)', 
                        'ต้องใช้ทั้งหมด (Total Req)', 'จำนวนที่ขาด', 'หน่วย'
                    ]].copy()
                    df_short_display.rename(columns={'จำนวนที่ขาด': 'Shorted'}, inplace=True)

                    # ตัวหนังสือสีแดง ตัวหนา + กรอบแดง (ไม่มีพื้นหลัง)
                    st.dataframe(
                        df_short_display.style.format({
                            'Usage/Unit (ต่อหน่วย)': '{:,.4f}',
                            'Stock ในคลังจริง (On Hand)': '{:,.2f}',
                            'ต้องใช้ทั้งหมด (Total Req)': '{:,.2f}',
                            'Shorted': '{:,.2f}'
                        }).map(
                            lambda x: 'color: #dc2626; font-weight: 800; border: 2px solid #ef4444;',
                            subset=['Shorted']
                        ),
                        use_container_width=True,
                        hide_index=True
                    )
                else:
                    st.success("🎉 ไม่มีรายการชิ้นส่วนที่ช็อตใน Model นี้")

        # --- TAB 3: Consolidated View ---
        with tab3:
            st.subheader("📊 ตารางรวมรายการชิ้นส่วนที่ใช้ซ้ำกันข้าม Model")
            
            all_rows = []
            for m in models_lvl0:
                df_m = df_master_bom[df_master_bom['Model'] == m]
                tn = build_tree_structure(df_m, stock_map, target_qty_input)
                df_ap = get_all_parts_df(tn)
                df_ap['Model'] = m
                all_rows.append(df_ap)
                
            df_consolidated = pd.concat(all_rows, ignore_index=True)
            
            grouped_summary = df_consolidated.groupby(['Item Code', 'Description', 'หน่วย']).agg(
                Grand_Total=('ต้องใช้ทั้งหมด (Total Req)', 'sum'),
                Model_Count=('Model', 'nunique'),
                Stock_On_Hand=('Stock ในคลังจริง (On Hand)', 'first')
            ).reset_index().sort_values(by='Model_Count', ascending=False)
            
            search_kw = st.text_input("🔍 ค้นหา Item / Description:")
            if search_kw:
                grouped_summary = grouped_summary[
                    grouped_summary['Item Code'].str.contains(search_kw, case=False) | 
                    grouped_summary['Description'].str.contains(search_kw, case=False)
                ]
                
            col_dl1, col_dl2 = st.columns([3, 1])
            with col_dl1:
                st.caption("คลิกที่รายการชิ้นส่วนด้านล่าง เพื่อเปิดดูรายละเอียดความต้องการของแต่ละ Model")
            with col_dl2:
                csv_data = grouped_summary.to_csv(index=False).encode('utf-8-sig')
                st.download_button(
                    label="📥 ดาวน์โหลดตารางสรุป (CSV)",
                    data=csv_data,
                    file_name="BOM_Consolidated_Summary.csv",
                    mime="text/csv"
                )
            
            for idx, row in grouped_summary.iterrows():
                item_code = row['Item Code']
                desc = row['Description']
                total = row['Grand_Total']
                um = row['หน่วย']
                m_count = row['Model_Count']
                stock = row['Stock_On_Hand']
                
                with st.expander(f"📦 **{item_code}** — {desc} | **Stock On Hand: {stock:,.2f} {um}** | **ความต้องการรวมทุกรุ่น: {total:,.2f} {um}** (ใช้ใน {m_count} Model)"):
                    df_detail = df_consolidated[(df_consolidated['Item Code'] == item_code) & (df_consolidated['Description'] == desc)]
                    df_detail_grouped = df_detail.groupby('Model').agg(
                        Usage_Unit=('Usage/Unit (ต่อหน่วย)', 'first'),
                        Total_Required=('ต้องใช้ทั้งหมด (Total Req)', 'sum')
                    ).reset_index()
                    df_detail_grouped.columns = ['Model (รุ่นที่ใช้)', 'Usage / Unit', f'จำนวนที่ต้องใช้ ({um})']
                    
                    st.dataframe(
                        df_detail_grouped.style.format({
                            'Usage / Unit': '{:,.2f}',
                            f'จำนวนที่ต้องใช้ ({um})': '{:,.2f}'
                        }),
                        use_container_width=True
                    )
    else:
        st.info("👈 กรุณาอัปโหลดไฟล์ `Drop short Movement...` หรือติ๊ก **โหมดทดสอบสมมติค่า Stock** ที่ Sidebar ด้านซ้าย")