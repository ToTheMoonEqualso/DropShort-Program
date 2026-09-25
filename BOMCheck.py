import pandas as pd
import json
from datetime import datetime
import re

def generate_full_dashboard():
    df_onhand = pd.read_excel('TIKONHAND (100926 11.40).xlsx')
    df_onhand.columns = ['LPROD', 'ONHAND']
    onhand_dict = dict(zip(df_onhand['LPROD'], df_onhand['ONHAND']))

    sheets_config = {
        'movement': ("Drop short Movement on assy Sep'26.xlsx", 'Drop short Movement on8-30S (2)'),
        'asy': ("Drop short Movement on assy Sep'26.xlsx", 'Drop short Movement on8-30S (2)'),
        'clock': ("Drop short Movement on assy Sep'26.xlsx", 'Drop short Movement on8-30S (2)'),
        'dec': ("Drop short Movement on assy Sep'26.xlsx", 'Drop short Movement on8-30S (2)'),
        'pcbcoil': ("Drop short Movement on assy Sep'26.xlsx", 'Drop short Movement on8-30S (2)')
    }

    combined_parts_df = []
    for cat, (filename, sheet_name) in sheets_config.items():
        try:
            df_drop = pd.read_excel(filename, sheet_name=sheet_name, header=None)
            df_drop_clean = df_drop.iloc[7:].copy()
            parts_df = df_drop_clean[[1, 3, 4, 5, 6, 9, 10]].copy()
            parts_df.columns = ['Model', 'Warehouse', 'ReleaseDate', 'ShopNo', 'Requirement', 'Part', 'PartDesc']
            
            parts_df['Model'] = parts_df['Model'].ffill()
            parts_df['Warehouse'] = parts_df['Warehouse'].ffill()
            parts_df['ReleaseDate'] = pd.to_datetime(parts_df['ReleaseDate'], errors='coerce')
            parts_df['ReleaseDate'] = parts_df['ReleaseDate'].ffill()
            parts_df['ShopNo'] = parts_df['ShopNo'].ffill()
            parts_df['Requirement'] = pd.to_numeric(parts_df['Requirement'], errors='coerce').ffill()
            parts_df = parts_df.dropna(subset=['Model']).copy()
            parts_df['Category'] = cat
            combined_parts_df.append(parts_df)
        except Exception as e:
            print(f"Warning loading {cat}: {e}")

    full_parts_df = pd.concat(combined_parts_df, ignore_index=True) if combined_parts_df else pd.DataFrame()

    current_date = pd.Timestamp(datetime.now().date())
    parts_df_filtered = full_parts_df[full_parts_df['ReleaseDate'] >= current_date].copy() if not full_parts_df.empty else pd.DataFrame()

    df_fgd = pd.read_excel('Usage & Level.xlsx', sheet_name='FGD')
    df_notest = pd.read_excel('Usage & Level.xlsx', sheet_name='NOTEST')

    def parse_bom_sheet(df, is_notest=False):
        bom_dict = {}
        current_group_key = None
        current_components = []

        for _, row in df.iterrows():
            model_desc = str(row['Model Description']).strip() if 'Model Description' in row else str(row.iloc[0]).strip()
            lvl_str = str(row['Level']).strip() if 'Level' in row else str(row.iloc[3]).strip()
            item = str(row['Item']).strip() if 'Item' in row else str(row.iloc[4]).strip()
            desc = str(row['Description']) if pd.notna(row['Description']) else 'Component Part'
            usage = float(row['Usage']) if 'Usage' in row and pd.notna(row['Usage']) else 1.0

            if is_notest:
                is_new_group = (lvl_str == '.1')
                clean_key = re.sub(r'[\s-]*TEST$', '', model_desc, flags=re.IGNORECASE).strip().upper()
            else:
                is_new_group = (lvl_str == '0' or lvl_str == '0.0' or lvl_str == '.1')
                clean_key = model_desc.strip().upper()

            if not current_group_key or (is_new_group and current_group_key != clean_key):
                if current_group_key and current_components:
                    bom_dict[current_group_key] = current_components
                current_group_key = clean_key
                current_components = []

            if is_notest:
                lvl_num = lvl_str.count('.')
            else:
                if lvl_str == '0' or lvl_str == '0.0':
                    lvl_num = 0
                else:
                    lvl_num = lvl_str.count('.')

            if current_group_key:
                current_components.append({
                    "Level": lvl_str,
                    "Level_Num": lvl_num,
                    "Part": item,
                    "Description": desc,
                    "Usage": usage
                })

        if current_group_key and current_components:
            bom_dict[current_group_key] = current_components

        return bom_dict

    bom_fgd = parse_bom_sheet(df_fgd, is_notest=False)
    bom_notest = parse_bom_sheet(df_notest, is_notest=True)

    def process_data(dataframe):
        remaining_onhand = {str(k).strip(): float(v) for k, v in onhand_dict.items() if pd.notna(v)}
        summary = []
        bom_details = {}

        for model_name, group in dataframe.groupby('Model', sort=False):
            display_name = str(model_name).strip()
            upper_name = display_name.upper()
            total_req = float(group['Requirement'].max())
            cat_val = group.iloc[0]['Category'] if 'Category' in group.columns else 'movement'
            
            raw_date = group.iloc[0]['ReleaseDate']
            mfg_date = raw_date.strftime('%Y-%m-%d') if pd.notna(raw_date) else '-'

            wh_val = str(group.iloc[0]['Warehouse']).strip() if pd.notna(group.iloc[0]['Warehouse']) else '-'
            shop_val = str(group.iloc[0]['ShopNo']).strip() if pd.notna(group.iloc[0]['ShopNo']) else '-'
            if shop_val.endswith('.0'):
                shop_val = shop_val[:-2]

            model_comps = []
            is_notest_model = ('-NO' in upper_name or '-N' in upper_name)
            
            if is_notest_model:
                model_comps = bom_notest.get(upper_name, [])
                if not model_comps:
                    for k, v in bom_notest.items():
                        if k == upper_name or upper_name in k:
                            model_comps = v
                            break
            else:
                model_comps = bom_fgd.get(upper_name, [])
                if not model_comps:
                    for k, v in bom_fgd.items():
                        if k == upper_name or upper_name == k:
                            model_comps = v
                            break

            model_shortage = False
            processed_components = []

            if model_comps:
                for c in model_comps:
                    part_code = c['Part']
                    usage_qty = c['Usage']
                    req_qty = total_req * usage_qty
                    
                    target_root_lvl = 1 if is_notest_model else 0
                    if c['Level_Num'] == target_root_lvl:
                        req_qty = total_req

                    is_special_69 = str(part_code).endswith('69')
                    available_qty = float(remaining_onhand.get(part_code, 0.0))

                    if is_special_69:
                        issued_qty = 0.0
                        remaining_qty = available_qty
                        status = ""
                    else:
                        status = 'Ready' if available_qty >= req_qty else 'Shortage'
                        if status == 'Shortage' and c['Level_Num'] > target_root_lvl:
                            model_shortage = True
                        issued_qty = min(available_qty, req_qty)
                        remaining_qty = max(0.0, available_qty - issued_qty)
                        remaining_onhand[part_code] = remaining_qty

                    processed_components.append({
                        "Level": c['Level'],
                        "Level_Num": c['Level_Num'],
                        "Part": part_code,
                        "Description": c['Description'],
                        "Usage": usage_qty,
                        "Requirement": req_qty,
                        "OnHandBeforeIssue": available_qty,
                        "Issued": issued_qty,
                        "OnHand": remaining_qty,
                        "Status": status
                    })
            else:
                valid_parts = group.dropna(subset=['Part'])
                for _, row in valid_parts.iterrows():
                    part_code = str(row['Part']).strip()
                    part_desc = str(row['PartDesc']) if pd.notna(row['PartDesc']) else 'Component Part'
                    req_qty = float(row['Requirement'])
                    available_qty = float(remaining_onhand.get(part_code, 0.0))
                    
                    is_special_69 = part_code.endswith('69')

                    if is_special_69:
                        status = ""
                        issued_qty = 0.0
                        remaining_qty = available_qty
                    else:
                        status = 'Ready' if available_qty >= req_qty else 'Shortage'
                        if status == 'Shortage':
                            model_shortage = True
                        issued_qty = min(available_qty, req_qty)
                        remaining_qty = max(0.0, available_qty - issued_qty)
                        remaining_onhand[part_code] = remaining_qty

                    processed_components.append({
                        "Level": ".1" if is_notest_model else "0",
                        "Level_Num": 1 if is_notest_model else 0,
                        "Part": part_code,
                        "Description": part_desc,
                        "Usage": 1.0,
                        "Requirement": req_qty,
                        "OnHandBeforeIssue": available_qty,
                        "Issued": issued_qty,
                        "OnHand": remaining_qty,
                        "Status": status
                    })

            bom_details[display_name] = processed_components
            overall_status = 'Shortage' if model_shortage else 'Ready'
            item_code = processed_components[0]['Part'] if processed_components else '-'
            effective_on_hand = 0.0 if overall_status == 'Shortage' else total_req

            summary.append({
                "Model": display_name,
                "Warehouse": wh_val,
                "ShopNo": shop_val,
                "ItemCode": item_code,
                "Requirement": total_req,
                "OnHand": effective_on_hand,
                "Status": overall_status,
                "ManufactureDate": mfg_date,
                "Category": cat_val
            })
        return summary, bom_details

    summary_data, bom_data = process_data(parts_df_filtered)

    json_data = json.dumps(summary_data, ensure_ascii=False).replace("</", "<\\/")
    json_bom = json.dumps(bom_data, ensure_ascii=False).replace("</", "<\\/")

    html_template = """<!DOCTYPE html>
<html lang="en" class="light">
<head>
  <meta charset="UTF-8">
  <meta name="viewport" content="width=device-width, initial-scale=1.0">
  <title>Production Risk & Shortage Command Center</title>
  <script src="https://cdn.jsdelivr.net/npm/@tailwindcss/browser@4"></script>
  <style>
    input[type="date"]::-webkit-calendar-picker-indicator {
      background: transparent;
      bottom: 0;
      color: transparent;
      cursor: pointer;
      height: auto;
      left: 0;
      position: absolute;
      right: 0;
      top: 0;
      width: auto;
    }
  </style>
</head>
<body class="bg-gray-100 text-gray-800 font-sans min-h-screen flex">

  <!-- SIDEBAR -->
  <aside id="sidebar" style="width: 256px; min-width: 256px;" class="bg-white border-r border-gray-200 p-4 h-screen sticky top-0 flex flex-col justify-between shadow-xs select-none transition-all duration-300 overflow-y-auto">
    <div class="space-y-8">
      <div class="space-y-1">
        <div class="text-[10px] font-bold text-gray-400 uppercase tracking-widest px-3 mb-2 sidebar-text">Main Menu</div>
        
        <button onclick="filterStatus('All')" id="sidebar-all" class="w-full relative flex items-center gap-3 px-4 py-2.5 rounded-xl text-xs font-bold bg-gray-50 text-gray-900 transition-all cursor-pointer">
          <span id="indicator-all" class="absolute left-0 top-1/2 -translate-y-1/2 w-1 h-5 bg-black rounded-r-full"></span>
          <svg class="w-4 h-4 text-gray-700 shrink-0" fill="none" stroke="currentColor" viewBox="0 0 24 24"><path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M4 6a2 2 0 012-2h2a2 2 0 012 2v2a2 2 0 01-2 2H6a2 2 0 01-2-2V6zM14 6a2 2 0 012-2h2a2 2 0 012 2v2a2 2 0 01-2 2h-2a2 2 0 01-2-2V6zM4 16a2 2 0 012-2h2a2 2 0 012 2v2a2 2 0 01-2 2H6a2 2 0 01-2-2v-2zM14 16a2 2 0 012-2h2a2 2 0 012 2v2a2 2 0 01-2 2h-2a2 2 0 01-2-2v-2z"></path></svg>
          <span class="sidebar-text">Dashboard</span>
        </button>

        <button onclick="filterStatus('Shortage')" id="sidebar-shortage" class="w-full relative flex items-center gap-3 px-4 py-2.5 rounded-xl text-xs font-medium text-gray-600 hover:bg-gray-50 transition-all cursor-pointer">
          <span id="indicator-shortage" class="absolute left-0 top-1/2 -translate-y-1/2 w-1 h-5 bg-black rounded-r-full hidden"></span>
          <svg class="w-4 h-4 text-gray-500 shrink-0" fill="none" stroke="currentColor" viewBox="0 0 24 24"><path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M12 9v2m0 4h.01m-6.938 4h13.856c1.54 0 2.502-1.667 1.732-3L13.732 4c-.77-1.333-2.694-1.333-3.464 0L3.34 16c-.77 1.333.192 3 1.732 3z"></path></svg>
          <span class="sidebar-text">Shortage Blockers</span>
        </button>

        <button onclick="filterStatus('Ready')" id="sidebar-ready" class="w-full relative flex items-center gap-3 px-4 py-2.5 rounded-xl text-xs font-medium text-gray-600 hover:bg-gray-50 transition-all cursor-pointer">
          <span id="indicator-ready" class="absolute left-0 top-1/2 -translate-y-1/2 w-1 h-5 bg-black rounded-r-full hidden"></span>
          <svg class="w-4 h-4 text-gray-500 shrink-0" fill="none" stroke="currentColor" viewBox="0 0 24 24"><path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M9 12l2 2 4-4m6 2a9 9 0 11-18 0 9 9 0 0118 0z"></path></svg>
          <span class="sidebar-text">Ready to Produce</span>
        </button>
      </div>
    </div>

    <div class="space-y-4">
      <button onclick="toggleSidebar()" class="w-full flex items-center gap-2 px-2 py-2 text-xs font-bold text-gray-700 hover:bg-gray-50 rounded-lg transition-colors cursor-pointer">
        <span id="collapse-icon" class="text-sm font-black transition-transform duration-300">«</span>
        <span class="sidebar-text">Collapse sidebar</span>
      </button>
    </div>
  </aside>

  <main class="flex-1 flex flex-col min-h-screen">
    <header class="bg-white border-b border-gray-200 px-6 py-3 flex items-center justify-between gap-3 shadow-xs">
      <div class="w-72 shrink-0 relative">
        <span class="absolute inset-y-0 left-0 flex items-center pl-3 text-gray-400">
          <svg class="w-4 h-4" fill="none" stroke="currentColor" viewBox="0 0 24 24"><path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M21 21l-6-6m2-5a7 7 0 11-14 0 7 7 0 0114 0z"></path></svg>
        </span>
        <input type="text" id="search-input" placeholder="Quick search model or shop no..." class="w-full bg-gray-50 border border-gray-200 rounded-xl pl-9 pr-3 py-2 text-xs text-gray-800 placeholder-gray-400 focus:outline-none focus:border-black transition-all">
      </div>

      <div class="flex items-center gap-3 shrink-0">
        <div class="flex items-center gap-1 bg-gray-100 p-1 rounded-xl">
          <button onclick="switchView('asy')" id="tab-asy" class="relative px-3 py-1.5 rounded-lg text-xs font-semibold text-gray-500 hover:text-gray-900 transition-all cursor-pointer">
            ASY
            <span class="absolute bottom-0 left-2 right-2 h-0.5 bg-black rounded-full hidden"></span>
          </button>
          <button onclick="switchView('clock')" id="tab-clock" class="relative px-3 py-1.5 rounded-lg text-xs font-semibold text-gray-500 hover:text-gray-900 transition-all cursor-pointer">
            Clock
            <span class="absolute bottom-0 left-2 right-2 h-0.5 bg-black rounded-full hidden"></span>
          </button>
          <button onclick="switchView('dec')" id="tab-dec" class="relative px-3 py-1.5 rounded-lg text-xs font-semibold text-gray-500 hover:text-gray-900 transition-all cursor-pointer">
            DEC
            <span class="absolute bottom-0 left-2 right-2 h-0.5 bg-black rounded-full hidden"></span>
          </button>
          <button onclick="switchView('movement')" id="tab-movement" class="relative px-3 py-1.5 rounded-lg text-xs font-semibold text-gray-500 hover:text-gray-900 transition-all cursor-pointer">
            Movement
            <span class="absolute bottom-0 left-2 right-2 h-0.5 bg-black rounded-full hidden"></span>
          </button>
          <button onclick="switchView('pcbcoil')" id="tab-pcbcoil" class="relative px-3 py-1.5 rounded-lg text-xs font-semibold text-gray-500 hover:text-gray-900 transition-all cursor-pointer">
            PCB&COIL
            <span class="absolute bottom-0 left-2 right-2 h-0.5 bg-black rounded-full hidden"></span>
          </button>
        </div>
        
        <div class="h-6 w-[1px] bg-gray-200"></div>

        <div class="flex items-center gap-2">
          <div class="relative bg-gray-50 border border-gray-200 rounded-xl px-2.5 py-1.5 hover:border-black transition-colors cursor-pointer">
            <input type="date" id="date-filter-input" onchange="applyFilterAndSearch()" class="bg-transparent text-xs text-gray-700 uppercase focus:outline-none cursor-pointer w-26">
          </div>
          <button onclick="resetDateFilter()" class="cursor-pointer hover:scale-110 transition-transform select-none p-0.5 flex items-center justify-center" title="Clear data filter">
            <video src="https://cdn-icons-mp4.flaticon.com/512/8121/8121324.mp4" autoplay loop muted playsinline class="w-5 h-5 object-contain pointer-events-none"></video>
          </button>
        </div>
      </div>
    </header>

    <div id="dashboard-view" class="p-8 space-y-6 flex-1">
      
      <!-- Landing Page -->
      <div id="section-landing" class="space-y-4 py-8">
        <div class="max-w-xl">
          <h2 class="text-xl font-bold text-gray-900 mb-1">Select Production Category</h2>
          <p class="text-xs text-gray-500">Please choose a category below to check production status and component inventory.</p>
        </div>
        <div class="grid grid-cols-1 md:grid-cols-5 gap-6 pt-4">
          <div onclick="selectCategory('asy')" class="bg-white border-2 border-gray-200 hover:border-black p-6 rounded-2xl shadow-xs cursor-pointer transition-all hover:shadow-lg flex flex-col justify-between group h-40">
            <div class="flex justify-between items-start">
              <span class="text-[10px] font-bold text-gray-400 uppercase tracking-wider group-hover:text-black">Category</span>
              <span class="w-3 h-3 rounded-full bg-blue-500"></span>
            </div>
            <div>
              <h4 class="font-black text-gray-900 text-lg mb-1">ASY</h4>
              <p class="text-xs text-gray-500">Assembly Production</p>
            </div>
          </div>
          <div onclick="selectCategory('clock')" class="bg-white border-2 border-gray-200 hover:border-black p-6 rounded-2xl shadow-xs cursor-pointer transition-all hover:shadow-lg flex flex-col justify-between group h-40">
            <div class="flex justify-between items-start">
              <span class="text-[10px] font-bold text-gray-400 uppercase tracking-wider group-hover:text-black">Category</span>
              <span class="w-3 h-3 rounded-full bg-amber-500"></span>
            </div>
            <div>
              <h4 class="font-black text-gray-900 text-lg mb-1">Clock</h4>
              <p class="text-xs text-gray-500">Clock Mechanism</p>
            </div>
          </div>
          <div onclick="selectCategory('dec')" class="bg-white border-2 border-gray-200 hover:border-black p-6 rounded-2xl shadow-xs cursor-pointer transition-all hover:shadow-lg flex flex-col justify-between group h-40">
            <div class="flex justify-between items-start">
              <span class="text-[10px] font-bold text-gray-400 uppercase tracking-wider group-hover:text-black">Category</span>
              <span class="w-3 h-3 rounded-full bg-purple-500"></span>
            </div>
            <div>
              <h4 class="font-black text-gray-900 text-lg mb-1">DEC</h4>
              <p class="text-xs text-gray-500">DEC Components</p>
            </div>
          </div>
          <div onclick="selectCategory('movement')" class="bg-white border-2 border-gray-200 hover:border-black p-6 rounded-2xl shadow-xs cursor-pointer transition-all hover:shadow-lg flex flex-col justify-between group h-40">
            <div class="flex justify-between items-start">
              <span class="text-[10px] font-bold text-gray-400 uppercase tracking-wider group-hover:text-black">Category</span>
              <span class="w-3 h-3 rounded-full bg-emerald-500"></span>
            </div>
            <div>
              <h4 class="font-black text-gray-900 text-lg mb-1">Movement</h4>
              <p class="text-xs text-gray-500">Movement Production</p>
            </div>
          </div>
          <div onclick="selectCategory('pcbcoil')" class="bg-white border-2 border-gray-200 hover:border-black p-6 rounded-2xl shadow-xs cursor-pointer transition-all hover:shadow-lg flex flex-col justify-between group h-40">
            <div class="flex justify-between items-start">
              <span class="text-[10px] font-bold text-gray-400 uppercase tracking-wider group-hover:text-black">Category</span>
              <span class="w-3 h-3 rounded-full bg-rose-500"></span>
            </div>
            <div>
              <h4 class="font-black text-gray-900 text-lg mb-1">PCB&COIL</h4>
              <p class="text-xs text-gray-500">Electronics & Coils</p>
            </div>
          </div>
        </div>
      </div>

      <!-- Dashboard View with 3 Clickable Metric Cards -->
      <div id="section-table" class="space-y-6 hidden">
        
        <div class="flex items-center justify-between">
          <button onclick="backToLanding()" class="bg-white border border-gray-200 hover:bg-gray-50 text-gray-700 px-4 py-2 rounded-xl text-xs font-bold transition-all cursor-pointer shadow-xs flex items-center gap-1.5">
            <span>←</span> Back to Categories
          </button>
        </div>

        <!-- 3 Clickable Metric Cards -->
        <div class="grid grid-cols-1 md:grid-cols-3 gap-5">
          <div onclick="filterStatus('All')" id="card-total" class="bg-white border-2 border-gray-200 hover:border-black p-5 rounded-2xl shadow-xs flex items-center justify-between cursor-pointer transition-all hover:shadow-md group">
            <div class="space-y-1">
              <span class="text-[11px] font-bold text-gray-400 uppercase tracking-wider group-hover:text-gray-900 transition-colors">Total Models</span>
              <h3 id="stat-total" class="text-2xl font-black text-gray-900">0</h3>
            </div>
            <div class="w-10 h-10 rounded-xl bg-gray-50 border border-gray-100 flex items-center justify-center text-gray-600 group-hover:bg-black group-hover:text-white transition-all">
              <svg class="w-5 h-5" fill="none" stroke="currentColor" viewBox="0 0 24 24"><path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M19 11H5m14 0a2 2 0 012 2v6a2 2 0 01-2 2H5a2 2 0 01-2-2v-6a2 2 0 012-2m14 0V9a2 2 0 00-2-2M5 11V9a2 2 0 012-2m0 0V5a2 2 0 012-2h6a2 2 0 012 2v2M7 7h10"></path></svg>
            </div>
          </div>

          <div onclick="filterStatus('Ready')" id="card-ready" class="bg-white border-2 border-emerald-100 hover:border-emerald-500 p-5 rounded-2xl shadow-xs flex items-center justify-between cursor-pointer transition-all hover:shadow-md group">
            <div class="space-y-1">
              <span class="text-[11px] font-bold text-emerald-600 uppercase tracking-wider">Ready to Produce</span>
              <h3 id="stat-ready" class="text-2xl font-black text-emerald-600">0</h3>
            </div>
            <div class="w-10 h-10 rounded-xl bg-emerald-50 border border-emerald-100 flex items-center justify-center text-emerald-600 group-hover:bg-emerald-500 group-hover:text-white transition-all">
              <svg class="w-5 h-5" fill="none" stroke="currentColor" viewBox="0 0 24 24"><path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M9 12l2 2 4-4m6 2a9 9 0 11-18 0 9 9 0 0118 0z"></path></svg>
            </div>
          </div>

          <div onclick="filterStatus('Shortage')" id="card-shortage" class="bg-white border-2 border-red-100 hover:border-red-500 p-5 rounded-2xl shadow-xs flex items-center justify-between cursor-pointer transition-all hover:shadow-md group">
            <div class="space-y-1">
              <span class="text-[11px] font-bold text-red-600 uppercase tracking-wider">Shortage Blockers</span>
              <h3 id="stat-shortage" class="text-2xl font-black text-red-600">0</h3>
            </div>
            <div class="w-10 h-10 rounded-xl bg-red-50 border border-red-100 flex items-center justify-center text-red-600 group-hover:bg-red-600 group-hover:text-white transition-all">
              <svg class="w-5 h-5" fill="none" stroke="currentColor" viewBox="0 0 24 24"><path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M12 9v2m0 4h.01m-6.938 4h13.856c1.54 0 2.502-1.667 1.732-3L13.732 4c-.77-1.333-2.694-1.333-3.464 0L3.34 16c-.77 1.333.192 3 1.732 3z"></path></svg>
            </div>
          </div>
        </div>

        <!-- Table View -->
        <div class="bg-white p-6 rounded-2xl border border-gray-200 shadow-xs space-y-4">
          <div class="flex flex-col md:flex-row justify-between items-start md:items-center gap-4">
            <h3 id="table-main-title" class="text-sm font-bold text-gray-900">Model Production Status & Schedule</h3>
            <div class="flex items-center gap-3">
              <div class="flex items-center gap-2">
                <span class="text-xs font-bold text-gray-400 uppercase tracking-wider">Part Filter:</span>
                <div class="relative">
                  <select id="dashboard-part-filter" onchange="applyFilterAndSearch()" class="appearance-none bg-gray-50 hover:bg-gray-100 border border-gray-200 text-gray-800 text-xs font-bold rounded-lg px-3 py-1.5 pr-8 focus:outline-none focus:ring-2 focus:ring-black transition-all cursor-pointer">
                    <option value="ALL">All Types</option>
                  </select>
                  <div class="pointer-events-none absolute inset-y-0 right-0 flex items-center px-2 text-gray-500">
                    <svg class="w-3.5 h-3.5" fill="none" stroke="currentColor" viewBox="0 0 24 24"><path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M19 9l-7 7-7-7"></path></svg>
                  </div>
                </div>
              </div>
              <span id="table-count-label" class="text-xs font-semibold text-gray-400 bg-gray-50 px-3 py-1 rounded-full border border-gray-100">Showing active models</span>
            </div>
          </div>
          <div class="overflow-x-auto">
            <table class="w-full text-left border-collapse">
              <thead>
                <tr class="border-b border-gray-100 text-[11px] font-bold text-gray-400 uppercase tracking-wider bg-gray-50/50">
                  <th class="py-3.5 px-4 rounded-l-xl">Model Name</th>
                  <th class="py-3.5 px-4 text-center">Warehouse</th>
                  <th class="py-3.5 px-4 text-center">Manufacture Date</th>
                  <th class="py-3.5 px-4 text-center">Shop No.</th>
                  <th class="py-3.5 px-4">Status</th>
                  <th class="py-3.5 px-4 rounded-r-xl">Action</th>
                </tr>
              </thead>
              <tbody id="table-body" class="divide-y divide-gray-50 text-sm"></tbody>
            </table>
          </div>
        </div>

      </div>

      <!-- Shortage Inline Details View -->
      <div id="section-shortage-inline" class="hidden space-y-6">
        
        <div class="flex items-center justify-between">
          <button onclick="backToLanding()" class="bg-white border border-gray-200 hover:bg-gray-50 text-gray-700 px-4 py-2 rounded-xl text-xs font-bold transition-all cursor-pointer shadow-xs flex items-center gap-1.5">
            <span>←</span> Back to Categories
          </button>
        </div>

        <!-- 3 Clickable Metric Cards -->
        <div class="grid grid-cols-1 md:grid-cols-3 gap-5">
          <div onclick="filterStatus('All')" id="card-shortage-total" class="bg-white border-2 border-gray-200 hover:border-black p-5 rounded-2xl shadow-xs flex items-center justify-between cursor-pointer transition-all hover:shadow-md group">
            <div class="space-y-1">
              <span class="text-[11px] font-bold text-gray-400 uppercase tracking-wider group-hover:text-gray-900 transition-colors">Total Models</span>
              <h3 id="stat-shortage-total" class="text-2xl font-black text-gray-900">0</h3>
            </div>
            <div class="w-10 h-10 rounded-xl bg-gray-50 border border-gray-100 flex items-center justify-center text-gray-600 group-hover:bg-black group-hover:text-white transition-all">
              <svg class="w-5 h-5" fill="none" stroke="currentColor" viewBox="0 0 24 24"><path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M19 11H5m14 0a2 2 0 012 2v6a2 2 0 01-2 2H5a2 2 0 01-2-2v-6a2 2 0 012-2m14 0V9a2 2 0 00-2-2M5 11V9a2 2 0 012-2m0 0V5a2 2 0 012-2h6a2 2 0 012 2v2M7 7h10"></path></svg>
            </div>
          </div>

          <div onclick="filterStatus('Ready')" id="card-shortage-ready" class="bg-white border-2 border-emerald-100 hover:border-emerald-500 p-5 rounded-2xl shadow-xs flex items-center justify-between cursor-pointer transition-all hover:shadow-md group">
            <div class="space-y-1">
              <span class="text-[11px] font-bold text-emerald-600 uppercase tracking-wider">Ready to Produce</span>
              <h3 id="stat-shortage-ready" class="text-2xl font-black text-emerald-600">0</h3>
            </div>
            <div class="w-10 h-10 rounded-xl bg-emerald-50 border border-emerald-100 flex items-center justify-center text-emerald-600 group-hover:bg-emerald-500 group-hover:text-white transition-all">
              <svg class="w-5 h-5" fill="none" stroke="currentColor" viewBox="0 0 24 24"><path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M9 12l2 2 4-4m6 2a9 9 0 11-18 0 9 9 0 0118 0z"></path></svg>
            </div>
          </div>

          <div onclick="filterStatus('Shortage')" id="card-shortage-shortage" class="bg-white border-2 border-red-100 hover:border-red-500 p-5 rounded-2xl shadow-xs flex items-center justify-between cursor-pointer transition-all hover:shadow-md group ring-2 ring-red-500">
            <div class="space-y-1">
              <span class="text-[11px] font-bold text-red-600 uppercase tracking-wider">Shortage Blockers</span>
              <h3 id="stat-shortage-shortage" class="text-2xl font-black text-red-600">0</h3>
            </div>
            <div class="w-10 h-10 rounded-xl bg-red-50 border border-red-100 flex items-center justify-center text-red-600 group-hover:bg-red-600 group-hover:text-white transition-all">
              <svg class="w-5 h-5" fill="none" stroke="currentColor" viewBox="0 0 24 24"><path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M12 9v2m0 4h.01m-6.938 4h13.856c1.54 0 2.502-1.667 1.732-3L13.732 4c-.77-1.333-2.694-1.333-3.464 0L3.34 16c-.77 1.333.192 3 1.732 3z"></path></svg>
            </div>
          </div>
        </div>

        <div class="flex justify-between items-center bg-red-50/50 border border-red-200 p-5 rounded-2xl">
          <div>
            <h3 class="text-sm font-bold text-red-700">Shortage Blockers - Inline Details</h3>
            <p class="text-xs text-red-500 mt-0.5">Showing models with component shortage blockers for quick inspection.</p>
          </div>
          <!-- กรอบแสดงจำนวนพาร์ทที่ขาด (Shortage Parts) ทำงานตามฟิลเตอร์วันที่และประเภทเรียบร้อย -->
          <span id="shortage-count-badge" class="px-3 py-1.5 rounded-xl bg-red-600 text-white text-xs font-bold shadow-xs"></span>
        </div>
        <div id="shortage-inline-container" class="space-y-4"></div>
      </div>
    </div>

    <div id="tree-view" class="hidden p-8 space-y-6 flex-1">
      <div class="flex flex-col md:flex-row items-start md:items-center justify-between border-b border-gray-200 pb-4 gap-4">
        <div class="flex items-center gap-4">
          <button onclick="showDashboard()" class="bg-white border border-gray-200 hover:bg-gray-50 text-gray-700 px-4 py-2 rounded-xl text-xs font-bold shadow-xs cursor-pointer transition-all">
            ← Back to Dashboard
          </button>
          <h2 id="tree-title" class="text-base font-bold text-gray-900">BOM Hierarchy Tree</h2>
        </div>
        
        <div class="flex items-center gap-3">
          <span class="text-xs font-bold text-gray-400 uppercase tracking-wider">Filter Type:</span>
          <div class="relative">
            <select id="tree-part-filter" onchange="filterBOMTree()" class="appearance-none bg-gray-50 hover:bg-gray-100 border border-gray-200 text-gray-800 text-xs font-bold rounded-lg px-3 py-1.5 pr-8 focus:outline-none focus:ring-2 focus:ring-black transition-all cursor-pointer">
              <option value="ALL">All Types</option>
            </select>
            <div class="pointer-events-none absolute inset-y-0 right-0 flex items-center px-2 text-gray-500">
              <svg class="w-3.5 h-3.5" fill="none" stroke="currentColor" viewBox="0 0 24 24"><path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M19 9l-7 7-7-7"></path></svg>
            </div>
          </div>
        </div>
      </div>

      <div class="bg-white p-6 rounded-2xl border border-gray-200 shadow-xs space-y-4">
        <div class="overflow-x-auto">
          <table class="w-full text-left border-collapse">
            <thead>
              <tr class="border-b border-gray-100 text-[11px] font-bold text-gray-400 uppercase tracking-wider bg-gray-50/50">
                <th class="py-3.5 px-4 font-semibold">Part Code & Description</th>
                <th class="py-3.5 px-4 font-semibold text-center w-28">Usage</th>
                <th class="py-3.5 px-4 font-semibold w-36">Requirement</th>
                <th class="py-3.5 px-4 font-semibold w-36">On-Hand</th>
                <th class="py-3.5 px-4 font-semibold w-48 text-center">Status</th>
              </tr>
            </thead>
            <tbody id="tree-table-body" class="divide-y divide-gray-50 text-sm"></tbody>
          </table>
        </div>
      </div>
    </div>
  </main>

  <script>
    const rawData = REPLACE_JSON_DATA;
    const bomDetails = REPLACE_JSON_BOM;
    let currentFilter = 'All';
    let currentView = '';
    let currentActiveModel = '';
    let isSidebarCollapsed = false;

    function initDashboardPartFilter() {
      const filterSelect = document.getElementById('dashboard-part-filter');
      let suffixes = new Set();
      
      Object.values(bomDetails).forEach(comps => {
        comps.forEach(c => {
          let suf = getPartTypeSuffix(c.Description);
          if (suf) suffixes.add(suf);
        });
      });

      filterSelect.innerHTML = '<option value="ALL">All Types</option>';
      Array.from(suffixes).sort().forEach(suf => {
        let opt = document.createElement('option');
        opt.value = suf;
        opt.innerText = suf;
        filterSelect.appendChild(opt);
      });
    }

    function toggleSidebar() {
      isSidebarCollapsed = !isSidebarCollapsed;
      const sidebar = document.getElementById('sidebar');
      const sidebarTexts = document.querySelectorAll('.sidebar-text');
      const collapseIcon = document.getElementById('collapse-icon');

      if (isSidebarCollapsed) {
        sidebar.style.width = '80px';
        sidebar.style.minWidth = '80px';
        sidebarTexts.forEach(el => el.style.display = 'none');
        collapseIcon.innerText = '»';
      } else {
        sidebar.style.width = '256px';
        sidebar.style.minWidth = '256px';
        sidebarTexts.forEach(el => el.style.display = '');
        collapseIcon.innerText = '«';
      }
    }

    function selectCategory(view) {
      document.getElementById('section-landing').classList.add('hidden');
      document.getElementById('section-table').classList.remove('hidden');
      switchView(view);
    }

    function backToLanding() {
      currentView = '';
      document.getElementById('section-table').classList.add('hidden');
      document.getElementById('section-shortage-inline').classList.add('hidden');
      document.getElementById('section-landing').classList.remove('hidden');
      
      ['asy', 'clock', 'dec', 'movement', 'pcbcoil'].forEach(v => {
        const btn = document.getElementById('tab-' + v);
        const indicator = btn ? btn.querySelector('span') : null;
        if (btn) {
          btn.className = 'relative px-3.5 py-2 rounded-lg text-xs font-semibold text-gray-500 hover:text-gray-900 transition-all cursor-pointer';
          if (indicator) indicator.classList.add('hidden');
        }
      });
    }

    function switchView(view) {
      currentView = view;
      document.getElementById('section-landing').classList.add('hidden');

      if (currentFilter === 'Shortage') {
        document.getElementById('section-table').classList.add('hidden');
        document.getElementById('section-shortage-inline').classList.remove('hidden');
      } else {
        document.getElementById('section-shortage-inline').classList.add('hidden');
        document.getElementById('section-table').classList.remove('hidden');
      }

      ['asy', 'clock', 'dec', 'movement', 'pcbcoil'].forEach(v => {
        const btn = document.getElementById('tab-' + v);
        const indicator = btn ? btn.querySelector('span') : null;
        if (btn) {
          if (v === view) {
            btn.className = 'relative px-3.5 py-2 rounded-lg text-xs font-bold bg-white text-gray-900 shadow-xs transition-all cursor-pointer';
            if (indicator) indicator.classList.remove('hidden');
          } else {
            btn.className = 'relative px-3.5 py-2 rounded-lg text-xs font-semibold text-gray-500 hover:text-gray-900 transition-all cursor-pointer';
            if (indicator) indicator.classList.add('hidden');
          }
        }
      });

      applyFilterAndSearch();
    }

    function renderAll(data, baseFilteredData) {
      updateSummaryCards(baseFilteredData);
      if (currentFilter === 'Shortage') {
        renderShortageInline(data);
      } else {
        renderTable(data);
      }
    }

    function updateSummaryCards(baseFilteredData) {
      let total = baseFilteredData.length;
      let ready = baseFilteredData.filter(i => i.Status === 'Ready').length;
      let shortage = baseFilteredData.filter(i => i.Status === 'Shortage').length;

      document.getElementById('stat-total').innerText = total;
      document.getElementById('stat-ready').innerText = ready;
      document.getElementById('stat-shortage').innerText = shortage;

      document.getElementById('stat-shortage-total').innerText = total;
      document.getElementById('stat-shortage-ready').innerText = ready;
      document.getElementById('stat-shortage-shortage').innerText = shortage;
    }

    function renderTable(data) {
      const tbody = document.getElementById('table-body');
      tbody.innerHTML = '';
      document.getElementById('table-count-label').innerText = 'Showing ' + data.length + ' active models';

      if (data.length === 0) {
        tbody.innerHTML = '<tr><td colspan="6" class="py-8 text-center text-gray-400">No active or upcoming models found</td></tr>';
        return;
      }

      data.forEach((row) => {
        const badge = row.Status === 'Shortage'
          ? '<span class="px-3 py-1 rounded-lg text-xs font-bold bg-red-50 text-red-600 border border-red-200">Shortage</span>'
          : '<span class="px-3 py-1 rounded-lg text-xs font-bold bg-emerald-50 text-emerald-600 border border-emerald-200">Ready</span>';

        const tr = document.createElement('tr');
        tr.className = 'hover:bg-gray-50/60 transition-colors';
        tr.innerHTML = `
          <td class="py-3.5 px-4 font-bold text-gray-900">${row.Model}</td>
          <td class="py-3.5 px-4 text-center text-gray-600 font-mono text-xs font-medium">${row.Warehouse || "-"}</td>
          <td class="py-3.5 px-4 text-center text-gray-600 text-xs font-medium">${row.ManufactureDate || "-"}</td>
          <td class="py-3.5 px-4 text-center text-gray-600 font-mono text-xs font-medium">${row.ShopNo || "-"}</td>
          <td class="py-3.5 px-4">${badge}</td>
          <td class="py-3.5 px-4"></td>
        `;
        
        const btn = document.createElement('button');
        btn.className = 'inline-flex items-center gap-1.5 text-xs font-bold text-blue-600 hover:text-blue-700 bg-blue-50/80 hover:bg-blue-100 px-3.5 py-1.5 rounded-xl transition-all cursor-pointer border border-blue-200';
        btn.innerHTML = 'View BOM Tree ↗';
        btn.onclick = () => showBOMTree(row.Model);
        tr.cells[5].appendChild(btn);

        tbody.appendChild(tr);
      });
    }

    function renderShortageInline(data) {
      const container = document.getElementById('shortage-inline-container');
      container.innerHTML = '';

      const selectedPartType = document.getElementById('dashboard-part-filter').value;
      let totalShortagePartsCount = 0;

      data.forEach(row => {
        const comps = bomDetails[row.Model] || [];
        const shortageComps = comps.filter(c => {
          if (c.Status !== 'Shortage') return false;
          if (selectedPartType && selectedPartType !== 'ALL') {
            return getPartTypeSuffix(c.Description) === selectedPartType;
          }
          return true;
        });
        totalShortagePartsCount += shortageComps.length;
      });

      document.getElementById('shortage-count-badge').innerText = totalShortagePartsCount + ' Shortage Parts';

      if (data.length === 0) {
        container.innerHTML = '<div class="bg-white p-8 rounded-2xl text-center text-gray-400 border border-gray-200">Excellent! No shortage blockers found in this category.</div>';
        return;
      }

      data.forEach(row => {
        const comps = bomDetails[row.Model] || [];
        const shortageComps = comps.filter(c => {
          if (c.Status !== 'Shortage') return false;
          if (selectedPartType && selectedPartType !== 'ALL') {
            return getPartTypeSuffix(c.Description) === selectedPartType;
          }
          return true;
        });

        if (shortageComps.length === 0 && selectedPartType !== 'ALL') return;

        let compsHtml = '';
        if (shortageComps.length === 0) {
          compsHtml = '<p class="text-xs text-gray-400 italic p-3">No direct shortage components found matching filter.</p>';
        } else {
          compsHtml = `
            <div class="overflow-x-auto mt-2">
              <table class="w-full text-left text-xs bg-white rounded-xl border border-red-100">
                <thead class="bg-red-50/40 text-red-700 uppercase tracking-wider">
                  <tr>
                    <th class="p-3">Level</th>
                    <th class="p-3">Part Code & Description</th>
                    <th class="p-3 text-center">Requirement</th>
                    <th class="p-3 text-center">On-Hand</th>
                    <th class="p-3 text-center">Shortage Qty</th>
                  </tr>
                </thead>
                <tbody class="divide-y divide-red-50">
          `;
          shortageComps.forEach(sc => {
            let diff = Math.round((sc.OnHandBeforeIssue - sc.Requirement) * 100) / 100;
            compsHtml += `
              <tr class="hover:bg-red-50/20">
                <td class="p-3 font-mono font-bold text-gray-600">${sc.Level}</td>
                <td class="p-3 font-semibold text-gray-900"><span class="font-mono text-blue-600 mr-2">${sc.Part}</span> ${sc.Description}</td>
                <td class="p-3 text-center text-gray-700">${Math.round(sc.Requirement * 100) / 100}</td>
                <td class="p-3 text-center text-gray-700">${Math.round(sc.OnHandBeforeIssue * 100) / 100}</td>
                <td class="p-3 text-center font-bold text-red-600">${diff.toLocaleString()}</td>
              </tr>
            `;
          });
          compsHtml += `</tbody></table></div>`;
        }

        const card = document.createElement('div');
        card.className = 'bg-white border border-red-200 rounded-2xl p-6 shadow-xs space-y-3';
        card.innerHTML = `
          <div class="flex flex-col md:flex-row justify-between items-start md:items-center gap-2 border-b border-gray-100 pb-3">
            <div>
              <div class="flex items-center gap-2">
                <h4 class="font-bold text-gray-900 text-base">${row.Model}</h4>
                <span class="px-2.5 py-1 rounded-lg text-[10px] font-black bg-red-100 text-red-700 tracking-wider uppercase">SHORTAGE BLOCKER</span>
              </div>
              <div class="flex items-center gap-4 text-xs text-gray-500 mt-1.5 font-mono">
                <span>Warehouse: <strong class="text-gray-700">${row.Warehouse || "-"}</strong></span>
                <span>Mfg Date: <strong class="text-gray-700">${row.ManufactureDate || "-"}</strong></span>
                <span>Shop No: <strong class="text-gray-700">${row.ShopNo || "-"}</strong></span>
              </div>
            </div>
          </div>
          <div>
            <span class="text-xs font-bold text-red-600 uppercase tracking-wider">Shortage Components:</span>
            ${compsHtml}
          </div>
        `;
        container.appendChild(card);
      });
    }

    function filterStatus(status) {
      currentFilter = status;
      ['all', 'shortage', 'ready'].forEach(s => {
        const btn = document.getElementById('sidebar-' + s);
        const ind = document.getElementById('indicator-' + s);
        if (btn) {
          btn.className = 'w-full relative flex items-center gap-3 px-4 py-2.5 rounded-xl text-xs font-medium text-gray-600 hover:bg-gray-50 transition-all cursor-pointer';
        }
        if (ind) ind.classList.add('hidden');
      });

      const activeBtn = document.getElementById('sidebar-' + status.toLowerCase());
      const activeInd = document.getElementById('indicator-' + status.toLowerCase());
      if (activeBtn) {
        activeBtn.className = 'w-full relative flex items-center gap-3 px-4 py-2.5 rounded-xl text-xs font-bold bg-gray-50 text-gray-900 transition-all cursor-pointer';
      }
      if (activeInd) activeInd.classList.remove('hidden');

      ['total', 'ready', 'shortage'].forEach(cardType => {
        const cardEl1 = document.getElementById('card-' + cardType);
        const cardEl2 = document.getElementById('card-shortage-' + cardType);
        if (cardEl1) {
          if (cardType === status.toLowerCase()) cardEl1.classList.add('ring-2', 'ring-black');
          else cardEl1.classList.remove('ring-2', 'ring-black');
        }
        if (cardEl2) {
          if (cardType === status.toLowerCase()) {
            if (cardType === 'shortage') cardEl2.classList.add('ring-2', 'ring-red-500');
            else if (cardType === 'ready') cardEl2.classList.add('ring-2', 'ring-emerald-500');
            else cardEl2.classList.add('ring-2', 'ring-black');
          } else {
            cardEl2.classList.remove('ring-2', 'ring-black', 'ring-red-500', 'ring-emerald-500');
          }
        }
      });

      if (currentView) {
        if (status === 'Shortage') {
          document.getElementById('section-table').classList.add('hidden');
          document.getElementById('section-shortage-inline').classList.remove('hidden');
        } else {
          document.getElementById('section-shortage-inline').classList.add('hidden');
          document.getElementById('section-table').classList.remove('hidden');
        }
        applyFilterAndSearch();
      }
    }

    function resetDateFilter() {
      document.getElementById('date-filter-input').value = '';
      applyFilterAndSearch();
    }

    function applyFilterAndSearch() {
      if (!currentView) return;
      const query = document.getElementById('search-input').value.toLowerCase();
      const selectedDate = document.getElementById('date-filter-input').value;
      const selectedPartType = document.getElementById('dashboard-part-filter').value;
      let baseFiltered = rawData;

      if (currentView) {
        baseFiltered = baseFiltered.filter(i => i.Category === currentView);
      }

      if (selectedDate) {
        baseFiltered = baseFiltered.filter(i => i.ManufactureDate === selectedDate);
      }

      if (selectedPartType && selectedPartType !== 'ALL') {
        baseFiltered = baseFiltered.filter(i => {
          const comps = bomDetails[i.Model] || [];
          return comps.some(c => getPartTypeSuffix(c.Description) === selectedPartType);
        });
      }

      if (query) {
        baseFiltered = baseFiltered.filter(i => {
          const modelMatch = i.Model.toLowerCase().includes(query);
          const shopNoMatch = String(i.ShopNo || '').toLowerCase().includes(query);
          return modelMatch || shopNoMatch;
        });
      }

      let statusFiltered = baseFiltered;
      if (currentFilter === 'Shortage') {
        statusFiltered = baseFiltered.filter(i => i.Status === 'Shortage');
      } else if (currentFilter === 'Ready') {
        statusFiltered = baseFiltered.filter(i => i.Status === 'Ready');
      }

      renderAll(statusFiltered, baseFiltered);
    }

    document.getElementById('search-input').addEventListener('input', applyFilterAndSearch);

    function getPartTypeSuffix(desc) {
      if (!desc) return '';
      let cleanDesc = desc.trim();
      if (cleanDesc.toUpperCase().endsWith('NOTEST')) {
        return cleanDesc.slice(-6).toUpperCase();
      }
      if (cleanDesc.length >= 3) {
        return cleanDesc.slice(-3).toUpperCase();
      }
      return cleanDesc.toUpperCase();
    }

    function showBOMTree(modelName) {
      currentActiveModel = modelName;
      document.getElementById('dashboard-view').classList.add('hidden');
      document.getElementById('tree-view').classList.remove('hidden');
      document.getElementById('tree-title').innerText = 'BOM Hierarchy Tree: ' + modelName;

      const comps = bomDetails[modelName] || [];
      const filterSelect = document.getElementById('tree-part-filter');
      filterSelect.innerHTML = '<option value="ALL">All Types</option>';

      let suffixes = new Set();
      comps.forEach(c => {
        let suf = getPartTypeSuffix(c.Description);
        if (suf) suffixes.add(suf);
      });

      Array.from(suffixes).sort().forEach(suf => {
        let opt = document.createElement('option');
        opt.value = suf;
        opt.innerText = suf;
        filterSelect.appendChild(opt);
      });
      filterSelect.value = 'ALL';

      renderBOMTreeTable(comps);
      window.scrollTo({ top: 0, behavior: 'smooth' });
    }

    function filterBOMTree() {
      const selectedType = document.getElementById('tree-part-filter').value;
      const comps = bomDetails[currentActiveModel] || [];

      if (selectedType === 'ALL') {
        renderBOMTreeTable(comps);
      } else {
        const filteredComps = comps.filter(c => getPartTypeSuffix(c.Description) === selectedType);
        renderBOMTreeTable(filteredComps);
      }
    }

    function renderBOMTreeTable(comps) {
      const tbody = document.getElementById('tree-table-body');
      tbody.innerHTML = '';

      if (comps.length === 0) {
        tbody.innerHTML = '<tr><td colspan="5" class="py-8 text-center text-gray-400">No components found matching this filter</td></tr>';
        return;
      }

      comps.forEach((c, index) => {
        c.index = index;
        c.children = [];
      });

      let stack = [];
      let rootItems = [];
      comps.forEach(c => {
        while (stack.length > 0 && stack[stack.length - 1].Level_Num >= c.Level_Num) {
          stack.pop();
        }
        if (stack.length === 0) {
          rootItems.push(c);
        } else {
          stack[stack.length - 1].children.push(c);
        }
        stack.push(c);
      });

      function renderRows(item, depth, parentPath) {
        let levelNum = item.Level_Num;
        let hasChildren = item.children && item.children.length > 0;
        let rowId = 'row-' + item.index;
        let pathKey = parentPath ? parentPath + '-' + item.index : 'root-' + item.index;

        let statusBadge = '';
        if (item.Status !== "") {
          if (item.Status === 'Shortage') {
            let diff = Math.round((item.OnHandBeforeIssue - item.Requirement) * 100) / 100;
            statusBadge = `<span class="px-3 py-1 rounded-lg text-xs font-bold bg-red-50 text-red-600 border border-red-300">SHORTAGE (${diff.toLocaleString()})</span>`;
          } else {
            if (item.OnHandBeforeIssue > item.Requirement) {
              let rem = Math.round((item.OnHandBeforeIssue - item.Requirement) * 100) / 100;
              statusBadge = `<span class="px-3 py-1 rounded-lg text-xs font-bold bg-emerald-50 text-emerald-600 border border-emerald-200">REMAINING: +${rem.toLocaleString()}</span>`;
            } else {
              statusBadge = '<span class="px-3 py-1 rounded-lg text-xs font-bold bg-gray-100 text-gray-600 border border-gray-200">PARENT IN STOCK</span>';
            }
          }
        }

        let minLevel = item.isNotest ? 1 : 0;
        let indentPx = (levelNum - minLevel) * 28;
        if (indentPx < 0) indentPx = 0;

        let levelBadgeClass = levelNum === minLevel ? 'bg-blue-100 text-blue-700 border-blue-300' : 'bg-gray-100 text-gray-600 border-gray-300';
        let toggleIcon = hasChildren ? `<span id="icon-${rowId}" class="text-gray-400 text-xs transition-transform transform rotate-90 cursor-pointer mr-2">▼</span>` : '<span class="w-4 inline-block mr-2"></span>';

        let tr = document.createElement('tr');
        tr.id = rowId;
        tr.dataset.path = pathKey;
        tr.className = `hover:bg-gray-50/80 transition-colors text-sm ${levelNum === minLevel ? 'bg-gray-50/60 font-bold' : 'border-t border-gray-100'}`;
        
        tr.innerHTML = `
          <td class="py-3.5 px-4">
            <div class="flex items-center" style="padding-left: ${indentPx}px;">
              ${toggleIcon}
              <span class="px-2 py-0.5 rounded text-xs font-mono font-bold border ${levelBadgeClass} mr-2">L${levelNum}</span>
              <span class="font-mono text-xs font-bold text-blue-600 mr-2">${item.Part}</span>
              <span class="text-gray-800">${item.Description}</span>
            </div>
          </td>
          <td class="py-3.5 px-4 text-center font-semibold text-gray-700">${item.Usage}</td>
          <td class="py-3.5 px-4 font-semibold text-gray-700">${Math.round(item.Requirement * 100) / 100} EA</td>
          <td class="py-3.5 px-4 font-semibold text-gray-700">${Math.round(item.OnHandBeforeIssue * 100) / 100}</td>
          <td class="py-3.5 px-4 text-center">${statusBadge}</td>
        `;
        tbody.appendChild(tr);

        let childTrs = [];
        if (hasChildren) {
          item.children.forEach(child => {
            childTrs.push(...renderRows(child, depth + 1, pathKey));
          });

          let isExpanded = true;
          let toggleHandler = (e) => {
            isExpanded = !isExpanded;
            let icon = document.getElementById('icon-' + rowId);
            if (icon) {
              if (isExpanded) {
                icon.classList.add('rotate-90');
                childTrs.forEach(t => t.style.display = '');
              } else {
                icon.classList.remove('rotate-90');
                childTrs.forEach(t => t.style.display = 'none');
              }
            }
          };
          tr.onclick = toggleHandler;
        }

        return [tr, ...childTrs];
      }

      rootItems.forEach(root => {
        renderRows(root, 0, '');
      });
    }

    function showDashboard() {
      document.getElementById('tree-view').classList.add('hidden');
      document.getElementById('dashboard-view').classList.remove('hidden');
      window.scrollTo({ top: 0, behavior: 'smooth' });
    }

    initDashboardPartFilter();
  </script>
</body>
</html>
"""

    html_content = html_template.replace("REPLACE_JSON_DATA", json_data).replace("REPLACE_JSON_BOM", json_bom)

    with open('index.html', 'w', encoding='utf-8') as f:
        f.write(html_content)
    print("Dashboard generated successfully with full date filter integration for shortage part counts!")

if __name__ == '__main__':
    generate_full_dashboard()