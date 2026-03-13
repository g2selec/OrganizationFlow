import streamlit as st
import pandas as pd
import numpy as np
import json
from streamlit_echarts import st_echarts

# --- 1. PAGE SETUP & CSS OPTIMIZATION ---
st.set_page_config(layout="wide", page_title="Org Flow Dashboard")

# Crush the default Streamlit white space to maximize chart size
st.markdown("""
    <style>
        .block-container {
            padding-top: 2rem !important;
            padding-bottom: 0rem !important;
        }
    </style>
""", unsafe_allow_html=True)

# Smaller, compact title
st.markdown("### 📊 Org Flow Architecture")

# --- 2. FILE UPLOADER & CONTROLS ---
with st.sidebar:
    st.header("Data Input & Settings")
    uploaded_file = st.file_uploader("Upload Excel Data", type=["xlsx"])
    st.markdown("---")
    start_depth = st.slider("Initial Explode Depth", min_value=1, max_value=5, value=3)

if uploaded_file:
    # Load Data & Clean
    df_raw = pd.read_excel(uploaded_file, sheet_name="RawData", header=1)
    df_dd = pd.read_excel(uploaded_file, sheet_name="DD")
    df_raw.columns = df_raw.columns.astype(str).str.strip()
    df_dd.columns = df_dd.columns.astype(str).str.strip()

    df_raw = df_raw.dropna(subset=["Name", "Role"])
    df_raw["Name"] = df_raw["Name"].astype(str).str.strip()
    
    for col in ["Reports To", "Team No", "Governance Body / Committee", "Group", "Team Type"]:
        if col in df_raw.columns:
            df_raw[col] = df_raw[col].fillna("").astype(str).str.strip().replace("nan", "")
        else:
            df_raw[col] = "" 

    # --- LEVEL ENGINE ---
    if 'Role' in df_dd.columns and 'Level' in df_dd.columns:
        role_to_level = dict(zip(df_dd['Role'], df_dd['Level']))
        df_raw['Mapped Level'] = df_raw['Role'].map(role_to_level).astype(str).str.upper().str.replace('L', '')
        df_raw['Mapped Level'] = pd.to_numeric(df_raw['Mapped Level'], errors='coerce')
    else:
        df_raw['Mapped Level'] = np.nan

    # --- GLOBAL HIERARCHY MAPPING ---
    name_to_level_global = dict(zip(df_raw['Name'], df_raw['Mapped Level']))
    all_names_global = set(df_raw["Name"].tolist())
    emp_to_mgr_global = {row["Name"]: str(row["Reports To"]).strip() for _, row in df_raw.iterrows()}
    
    hod_names_global = set()
    for name, mgr in emp_to_mgr_global.items():
        if not mgr or mgr.lower() == "nan" or mgr.lower() == "sir":
            hod_names_global.add(name)
        if mgr and mgr.lower() != "nan" and mgr.lower() != "sir" and mgr not in all_names_global:
            hod_names_global.add(mgr)

    # --- SMART GROUP FILTERING ---
    with st.sidebar:
        st.markdown("---")
        st.header("Filter & Export")
        unique_groups = sorted(list(set(g for g in df_raw["Group"].tolist() if g and g.lower() != "na")))
        selected_group = st.selectbox("Select Group to View:", ["All Groups"] + unique_groups)

    if selected_group != "All Groups":
        emps_in_group = set(df_raw[df_raw["Group"] == selected_group]["Name"].tolist())
        allowed_names = set(emps_in_group)
        for emp in emps_in_group:
            current = emp
            while True:
                mgr = emp_to_mgr_global.get(current)
                if not mgr or mgr.lower() == "nan" or mgr.lower() == "sir":
                    break
                allowed_names.add(mgr)
                current = mgr
                
        df_raw = df_raw[df_raw["Name"].isin(allowed_names)]
        hod_names = hod_names_global.intersection(allowed_names)
        name_to_level = {k: v for k, v in name_to_level_global.items() if k in allowed_names}
    else:
        hod_names = hod_names_global
        name_to_level = name_to_level_global

    all_names = set(df_raw["Name"].tolist())

    # --- 3. ECHARTS TREE BUILDER ENGINE ---
    missing_level_alerts = []
    node_data = {}      
    children_map = {}   

    def add_node(node_id, label, color):
        if node_id not in node_data:
            node_data[node_id] = {"name": str(label), "color": color}
        if node_id not in children_map:
            children_map[node_id] = []

    def add_edge(parent_id, child_id):
        if parent_id not in children_map:
            children_map[parent_id] = []
        if child_id not in children_map[parent_id]:
            children_map[parent_id].append(child_id)

    gov_id = "GOV_MAIN"
    add_node(gov_id, "Governance Body\n(Sir) [L0]", "#ffcccc")

    # PHASE 1: Employee Nodes
    for index, row in df_raw.iterrows():
        name, role, lvl = row["Name"], row["Role"], row["Mapped Level"]
        emp_id = f"EMP_{name}"
        is_hod = name in hod_names

        lvl_text = f"L{int(lvl)}" if pd.notna(lvl) else "Unmapped"
        emp_label = f"{name}\n({role})\n[{lvl_text}]"
        node_color = "#ffeb99" if is_hod else "#fff2cc"
        
        add_node(emp_id, emp_label, node_color)

    # PHASE 1B: Phantom HODs
    for hod in hod_names:
        if hod not in all_names:
            add_node(f"EMP_{hod}", f"{hod}\n(HOD)\n[L1]", "#ffeb99")
            name_to_level[hod] = 1 
            add_edge(gov_id, f"EMP_{hod}")

    # PHASE 2: ROUTING & MATH
    for index, row in df_raw.iterrows():
        name, mgr, emp_lvl = row["Name"], str(row["Reports To"]).strip(), row["Mapped Level"]
        group, tt, tno = row["Group"], row["Team Type"], row["Team No"]

        emp_id = f"EMP_{name}"

        if not mgr or mgr.lower() == "nan" or mgr.lower() == "sir":
            mgr = "Sir"
            parent_node = gov_id
            mgr_lvl = 0 
        elif mgr in hod_names:
            current_parent = f"EMP_{mgr}"
            if group and group.lower() != "na" and (selected_group == "All Groups" or group == selected_group):
                grp_id = f"GRP_{group}"
                add_node(grp_id, group, "#ccffcc")
                add_edge(current_parent, grp_id)
                current_parent = grp_id
                
                if tt and tt.lower() != "na":
                    tt_id = f"TT_{group}_{tt}"
                    add_node(tt_id, tt, "#cce5ff")
                    add_edge(current_parent, tt_id)
                    current_parent = tt_id
            parent_node = current_parent
            mgr_lvl = name_to_level.get(mgr, np.nan)
        else:
            mgr_id = f"EMP_{mgr}"
            current_parent = mgr_id
            grand_mgr = emp_to_mgr_global.get(mgr)
            
            if grand_mgr in hod_names and tno and tno.lower() != "na" and (selected_group == "All Groups" or group == selected_group):
                tno_id = f"TNO_{group}_{tt}_{tno}"
                add_node(tno_id, f"Team:\n{tno}", "#e6ccff")
                add_edge(current_parent, tno_id)
                current_parent = tno_id
            parent_node = current_parent
            mgr_lvl = name_to_level.get(mgr, np.nan)

        # GAP VALIDATION
        if pd.notna(emp_lvl) and pd.notna(mgr_lvl):
            gap = int(emp_lvl - mgr_lvl)
            if gap > 1:
                missing_lvls = [f"L{l}" for l in range(int(mgr_lvl) + 1, int(emp_lvl))]
                missing_str = ", ".join(missing_lvls)
                missing_id = f"MISSING_{parent_node}_{missing_str.replace(' ', '_')}"
                
                add_node(missing_id, f"⚠️ Missing\n{missing_str}", "#ff9999")
                add_edge(parent_node, missing_id)
                add_edge(missing_id, emp_id)
                
                alert_msg = f"Under **{mgr}**, missing **{missing_str}** detected!"
                if alert_msg not in missing_level_alerts:
                    missing_level_alerts.append(alert_msg)
            else:
                add_edge(parent_node, emp_id)
        else:
            add_edge(parent_node, emp_id)

    # PHASE 4: NESTED JSON
    def build_echarts_tree(node_id, visited=None):
        if visited is None: visited = set()
        if node_id in visited: return None 
        visited.add(node_id)
        
        node_info = node_data[node_id]
        
        tree_node = {
            "name": node_info["name"],
            "label": {
                "backgroundColor": node_info["color"],
                "borderColor": "#555",
                "borderWidth": 1,
                "padding": [8, 10], 
                "borderRadius": 5,
                "color": "#000",
                "fontSize": 12,
                "lineHeight": 18,
                "align": "center"
            }
        }
        
        children = []
        for child_id in children_map.get(node_id, []):
            child_tree = build_echarts_tree(child_id, visited.copy())
            if child_tree:
                children.append(child_tree)
                
        if children:
            tree_node["children"] = children
            
        return tree_node

    final_tree_data = build_echarts_tree(gov_id)

    # --- 4. RENDER FULL-WIDTH CHART ---
    options = {
        "tooltip": {"trigger": "item", "triggerOn": "mousemove"},
        "series": [
            {
                "type": "tree",
                "data": [final_tree_data],
                "orient": "TB", 
                "top": "5%", "bottom": "5%", "left": "2%", "right": "2%",
                "symbolSize": 12, 
                "initialTreeDepth": start_depth, 
                "roam": True, 
                "expandAndCollapse": True,
                "animationDuration": 550, "animationDurationUpdate": 750,
                "edgeShape": "polyline", 
                "lineStyle": {"width": 2, "color": "#aaa"},
                "label": {"position": "bottom", "verticalAlign": "middle", "align": "center"},
                "leaves": {"label": {"position": "bottom", "verticalAlign": "middle", "align": "center"}}
            }
        ]
    }
    
    # Render the interactive chart
    st_echarts(options=options, height="800px")

    # --- THE MAGIC HTML EXPORTER ---
    # We wrap the exact JSON data into a standalone HTML file
    html_template = f"""
    <!DOCTYPE html>
    <html>
    <head>
        <meta charset="utf-8">
        <title>Org Flow Export: {selected_group}</title>
        <script src="https://cdn.jsdelivr.net/npm/echarts@5.4.3/dist/echarts.min.js"></script>
        <style>
            html, body, #main {{ width: 100%; height: 100%; margin: 0; padding: 0; background-color: #ffffff; }}
        </style>
    </head>
    <body>
        <div id="main"></div>
        <script>
            var chart = echarts.init(document.getElementById('main'));
            var option = {json.dumps(options)};
            
            // Expand all nodes automatically for printing
            option.series[0].initialTreeDepth = -1; 
            
            chart.setOption(option);
        </script>
    </body>
    </html>
    """

    with st.sidebar:
        st.download_button(
            label="📥 Download Interactive HTML (For PDF Print)",
            data=html_template,
            file_name=f"Org_Flow_{selected_group.replace(' ', '_')}.html",
            mime="text/html",
            help="Download this file, open it in Chrome/Edge, and use 'Print to PDF' for perfect uncropped resolution!"
        )

    # --- 5. RENDER FOOTER DASHBOARD ---
    st.markdown("---")
    
    sum_col1, sum_col2, sum_col3 = st.columns(3)
    
    with sum_col1:
        st.subheader("Data Summary")
        st.metric("Total Employees", len(df_raw))

    with sum_col2:
        st.subheader("⚠️ Validation")
        if missing_level_alerts:
            for alert in missing_level_alerts:
                st.warning(alert)
        else:
            st.success("All reporting lines are contiguous.")

    with sum_col3:
        st.subheader("Unmapped Roles")
        unmapped = df_raw[df_raw['Mapped Level'].isna()]
        if not unmapped.empty:
            st.dataframe(unmapped[['Name', 'Role']], hide_index=True)
        else:
            st.success("All roles mapped perfectly!")

else:
    st.info("👈 Please upload your Excel file in the sidebar to begin.")
