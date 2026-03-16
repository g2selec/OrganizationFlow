import streamlit as st
import pandas as pd
import numpy as np
import json
from streamlit_echarts import st_echarts

# --- 1. PAGE SETUP & CSS OPTIMIZATION ---
st.set_page_config(layout="wide", page_title="Org Flow Dashboard")

st.markdown("""
    <style>
        .block-container {
            padding-top: 2rem !important;
            padding-bottom: 0rem !important;
        }
    </style>
""", unsafe_allow_html=True)

st.markdown("### 📊 Org Flow Architecture")

# --- 2. FILE UPLOADER & CONTROLS ---
with st.sidebar:
    st.header("Data Input & Settings")
    uploaded_file = st.file_uploader("Upload Excel Data", type=["xlsx"])
    st.markdown("---")
    start_depth = st.slider("Initial Explode Depth", min_value=1, max_value=6, value=1)
    cluster_l7 = st.checkbox("Collapse L7 Employees into Clusters", value=True)

if uploaded_file:
    # Load Data & Clean
    df_raw = pd.read_excel(uploaded_file, sheet_name="RawData", header=1)
    df_dd = pd.read_excel(uploaded_file, sheet_name="DD")
    df_raw.columns = df_raw.columns.astype(str).str.strip()
    df_dd.columns = df_dd.columns.astype(str).str.strip()

    df_raw = df_raw.dropna(subset=["Name", "Role"])
    df_raw["Name"] = df_raw["Name"].astype(str).str.strip()
    
    for col in ["Reports To", "Team No", "Governance Body / Committee", "Group", "Team Type", "Domain"]:
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

    # --- GLOBAL CACHE & ROW-LOCKED HEADCOUNT ENGINE ---
    df_raw_all = df_raw.copy()
    all_names_global = set(df_raw_all["Name"].tolist())
    
    top_level_heads = set()
    true_hods = set()
    mgr_to_emps_dict = {} # Now using sets to prevent double-counting matrix employees
    
    for _, row in df_raw_all.iterrows():
        name = str(row["Name"]).strip()
        role = str(row["Role"]).strip()
        mgr = str(row["Reports To"]).strip()
        
        if "hod" in role.lower():
            true_hods.add(name)
            
        if not mgr or mgr.lower() in ["nan", "sir", ""]:
            top_level_heads.add(name)
            
        if mgr and mgr.lower() not in ["nan", "sir", ""] and mgr not in all_names_global:
            top_level_heads.add(mgr)
            true_hods.add(mgr)
            
        if mgr and mgr.lower() not in ["nan", "sir", ""]:
            if mgr not in mgr_to_emps_dict:
                mgr_to_emps_dict[mgr] = set()
            mgr_to_emps_dict[mgr].add(name)

    mgr_silo_teams = {}
    for _, row in df_raw_all.iterrows():
        m = str(row["Reports To"]).strip()
        s = str(row["Group"]).strip()
        if not s or s.lower() == "na":
            s = "Ungrouped"
        t = str(row["Team No"]).strip()
        if m and t and t.lower() != "na":
            key = f"{m}_{s}"
            if key not in mgr_silo_teams:
                mgr_silo_teams[key] = set()
            mgr_silo_teams[key].add(t)

    # Recursive math engine to calculate total unique reports
    total_reports_cache = {}
    def get_total_reports(mgr_name, visited=None):
        if visited is None: visited = set()
        if mgr_name in visited: return set() 
        visited.add(mgr_name)
        
        if mgr_name in total_reports_cache:
            return total_reports_cache[mgr_name]
        
        directs = mgr_to_emps_dict.get(mgr_name, set())
        all_reports = set(directs)
        for d in directs:
            all_reports.update(get_total_reports(d, visited.copy()))
            
        total_reports_cache[mgr_name] = all_reports
        return all_reports

    for emp_name in list(mgr_to_emps_dict.keys()) + list(all_names_global):
        get_total_reports(emp_name)

    all_managers = set(mgr_to_emps_dict.keys())

    # --- STICKY NOTE EXTRACTION ---
    sticky_text = ""
    try:
        df_notes = pd.read_excel(uploaded_file, sheet_name="Notes", header=None)
        if df_notes.shape[0] >= 22 and df_notes.shape[1] >= 7:
            notes_slice = df_notes.iloc[21:24, 6].dropna().astype(str).tolist()
            if notes_slice:
                sticky_text = "\n".join(notes_slice).strip()
    except Exception as e:
        pass

    # --- SMART GROUP & HOD FILTERING (ROW LOCKED) ---
    with st.sidebar:
        st.markdown("---")
        st.header("Filter & Export")
        
        unique_hods = sorted(list(true_hods))
        selected_hod = st.selectbox("Select HOD to View:", ["All HODs"] + unique_hods)
        
        unique_groups = sorted(list(set(g for g in df_raw["Group"].tolist() if g and g.lower() != "na")))
        selected_group = st.selectbox("Select Group to View:", ["All Groups"] + unique_groups)

    # 🌟 NEW FIX: Row-Locked HOD Tracing 🌟
    def get_row_ultimate_hod(emp_name, emp_silo, current_row):
        curr_name = emp_name
        curr_silo = emp_silo
        curr_row = current_row
        visited = set()
        
        while curr_name:
            if curr_name in visited: break
            visited.add(curr_name)
            
            if curr_name in true_hods:
                return curr_name
                
            mgr_name = str(curr_row.get("Reports To", "")).strip()
            if not mgr_name or mgr_name.lower() in ["nan", "sir", ""]:
                break
                
            mgr_rows = df_raw_all[(df_raw_all["Name"] == mgr_name) & (df_raw_all["Group"] == curr_silo)]
            if not mgr_rows.empty:
                curr_row = mgr_rows.iloc[0]
            else:
                mgr_rows = df_raw_all[df_raw_all["Name"] == mgr_name]
                if not mgr_rows.empty:
                    curr_row = mgr_rows.iloc[0]
                else:
                    break
            curr_name = mgr_name
            
        return None

    if selected_group != "All Groups":
        df_raw = df_raw[df_raw["Group"] == selected_group]

    if selected_hod != "All HODs":
        keep_mask = []
        for _, row in df_raw.iterrows():
            e_name = str(row["Name"]).strip()
            e_silo = str(row["Group"]).strip()
            if not e_silo or e_silo.lower() == "na":
                e_silo = "Ungrouped"
            uhod = get_row_ultimate_hod(e_name, e_silo, row)
            keep_mask.append(uhod == selected_hod)
        df_raw = df_raw[keep_mask]

    all_names = set(df_raw["Name"].tolist())

    # --- 3. ECHARTS TREE BUILDER ENGINE ---
    missing_level_alerts = []
    node_data = {}      
    children_map = {}   
    built_nodes = {} 
    l7_clusters = {} 

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

    def add_emp_node_from_row(node_id, row, is_hod):
        if node_id not in node_data:
            name = str(row.get("Name", "")).strip()
            role = str(row.get("Role", "")).strip()
            domain = str(row.get("Domain", "")).strip()
            
            if not role or role.lower() == "nan":
                role = "HOD" if is_hod else "Unknown"
            if not domain or domain.lower() == "nan" or domain.lower() == "na":
                domain = "No Domain"
            
            lvl = row.get("Mapped Level")
            lvl_text = f"L{int(lvl)}" if pd.notna(lvl) else "Unmapped"
            
            direct_count = len(mgr_to_emps_dict.get(name, set()))
            total_count = len(total_reports_cache.get(name, set()))
            
            if total_count > 0:
                count_str = f"👥 D:{direct_count} | T:{total_count}"
                emp_label = f"{name}\n({role} | {domain})\n[{lvl_text}]\n{count_str}"
            else:
                emp_label = f"{name}\n({role} | {domain})\n[{lvl_text}]"
                
            node_color = "#ffeb99" if is_hod else "#fff2cc" 
            add_node(node_id, emp_label, node_color)

    gov_id = "GOV_MAIN"
    add_node(gov_id, "Governance Body\n(Sir) [L0]", "#ffcccc")

    # 🌟 THE ROW-LOCKED RECURSIVE ENGINE 🌟
    def trace_up_node(emp_name, emp_silo, current_row):
        is_top_level = emp_name in top_level_heads
        is_true_hod = emp_name in true_hods
        
        cache_key = emp_name if is_top_level else f"{emp_name}_{emp_silo}"

        if cache_key in built_nodes:
            return built_nodes[cache_key]

        mgr_name = str(current_row.get("Reports To", "")).strip()

        # --- TIER 1: DOMAIN HEAD ---
        if is_top_level:
            hod_grp = str(current_row.get("Group", "")).strip()
            if not hod_grp or hod_grp.lower() == "na":
                hod_grp = "Ungrouped Domain"

            top_grp_id = f"TOP_GRP_{hod_grp}"
            add_node(top_grp_id, f"🏢 {hod_grp}", "#ffd27f")
            add_edge(gov_id, top_grp_id)

            emp_id = f"EMP_{emp_name}"
            add_emp_node_from_row(emp_id, current_row, is_hod=is_true_hod)

            emp_lvl = current_row.get("Mapped Level")
            if pd.notna(emp_lvl):
                gap = int(emp_lvl - 0) 
                if gap > 1:
                    missing_lvls = [f"L{l}" for l in range(1, int(emp_lvl))]
                    missing_str = ", ".join(missing_lvls)
                    missing_id = f"MISSING_{top_grp_id}_{missing_str.replace(' ', '_')}"

                    add_node(missing_id, f"⚠️ Missing\n{missing_str}", "#ff9999")
                    add_edge(top_grp_id, missing_id)
                    add_edge(missing_id, emp_id)
                else:
                    add_edge(top_grp_id, emp_id)
            else:
                add_edge(top_grp_id, emp_id)

            built_nodes[cache_key] = emp_id
            return emp_id

        # --- TIER 2: RECURSE UP TO MANAGER ---
        mgr_rows = df_raw_all[(df_raw_all["Name"] == mgr_name) & (df_raw_all["Group"] == emp_silo)]
        if not mgr_rows.empty:
            mgr_row = mgr_rows.iloc[0]
        else:
            mgr_rows = df_raw_all[df_raw_all["Name"] == mgr_name]
            if not mgr_rows.empty:
                mgr_row = mgr_rows.iloc[0]
            else:
                mgr_row = pd.Series({"Name": mgr_name, "Reports To": "Sir", "Role": "HOD", "Mapped Level": np.nan, "Group": emp_silo})

        parent_id = trace_up_node(mgr_name, emp_silo, mgr_row)
        is_mgr_top_level = mgr_name in top_level_heads
        current_parent = parent_id

        # INJECT STRUCTURAL CONTAINERS
        if is_mgr_top_level:
            if emp_silo and emp_silo != "Ungrouped":
                grp_id = f"GRP_{emp_silo}"
                add_node(grp_id, emp_silo, "#ccffcc")
                add_edge(current_parent, grp_id)
                current_parent = grp_id

                tt = str(current_row.get("Team Type", "")).strip()
                if tt and tt.lower() != "na":
                    tt_id = f"TT_{emp_silo}_{tt}"
                    add_node(tt_id, tt, "#cce5ff")
                    add_edge(current_parent, tt_id)
                    current_parent = tt_id
        else:
            tno = str(current_row.get("Team No", "")).strip()
            tt = str(current_row.get("Team Type", "")).strip()
            mgr_tno = str(mgr_row.get("Team No", "")).strip()
            
            if tno and tno.lower() != "na":
                draw_team_box = True
                managed_teams = mgr_silo_teams.get(f"{mgr_name}_{emp_silo}", set())
                if tno == mgr_tno and len(managed_teams) <= 1:
                    draw_team_box = False
                    
                if draw_team_box:
                    tno_id = f"TNO_{emp_silo}_{mgr_name}_{tt}_{tno}"
                    add_node(tno_id, f"Team:\n{tno}", "#e6ccff")
                    add_edge(current_parent, tno_id)
                    current_parent = tno_id

        # GAP VALIDATION
        emp_lvl = current_row.get("Mapped Level")
        mgr_lvl = mgr_row.get("Mapped Level")
        final_parent = current_parent

        if pd.notna(emp_lvl) and pd.notna(mgr_lvl):
            gap = int(emp_lvl - mgr_lvl)
            if gap > 1:
                missing_lvls = [f"L{l}" for l in range(int(mgr_lvl) + 1, int(emp_lvl))]
                missing_str = ", ".join(missing_lvls)
                missing_id = f"MISSING_{current_parent}_{missing_str.replace(' ', '_')}"

                add_node(missing_id, f"⚠️ Missing\n{missing_str}", "#ff9999")
                add_edge(current_parent, missing_id)
                final_parent = missing_id

                alert_msg = f"Under **{mgr_name}** in {emp_silo}, missing **{missing_str}** detected!"
                if alert_msg not in missing_level_alerts:
                    missing_level_alerts.append(alert_msg)

        # 🌟 THE CLUSTERING LOGIC 🌟
        emp_id = f"EMP_{emp_name}_{emp_silo}"
        
        if cluster_l7 and emp_lvl == 7 and emp_name not in all_managers:
            if final_parent not in l7_clusters:
                l7_clusters[final_parent] = []
            l7_clusters[final_parent].append(emp_name)
        else:
            add_emp_node_from_row(emp_id, current_row, is_hod=is_true_hod)
            add_edge(final_parent, emp_id)

        built_nodes[cache_key] = emp_id
        return emp_id

    # --- FIRE THE ENGINE ---
    for index, row in df_raw.iterrows():
        emp_name = str(row["Name"]).strip()
        emp_silo = str(row["Group"]).strip()
        if not emp_silo or emp_silo.lower() == "na":
            emp_silo = "Ungrouped"
            
        trace_up_node(emp_name, emp_silo, row)

    # 🌟 BUILD THE L7 CLUSTER NODES 🌟
    for parent_id, names in l7_clusters.items():
        cluster_id = f"CLUSTER_L7_{parent_id}"
        names.sort()
        names_str = "\n".join(names)
        label = f"👥 L7 Technicians ({len(names)})\n{names_str}"
        add_node(cluster_id, label, "#e6f2ff")
        add_edge(parent_id, cluster_id)

    # --- 🌟 CUSTOM LEFT-TO-RIGHT SORTING 🌟 ---
    custom_order = {
        "TOP_GRP_Electrical": 1,
        "TOP_GRP_Process": 2,
        "TOP_GRP_Power Electronics": 3,
        "TOP_GRP_G4-PS": 4,   
        "TOP_GRP_G4-VFD": 5   
    }
    if gov_id in children_map:
        children_map[gov_id].sort(key=lambda x: custom_order.get(x, 99))

    # --- PHASE 4: NESTED JSON ---
    def build_echarts_tree(node_id, visited=None):
        if visited is None: visited = set()
        if node_id in visited: return None 
        visited.add(node_id)
        
        node_info = node_data[node_id]
        
        tree_node = {
            "id": node_id,
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

    # 🌟 BUILD THE TRUE FLOATING STICKY NOTE GRAPHIC 🌟
    graphic_elements = []
    if sticky_text:
        lines = sticky_text.split('\n')
        note_height = max(80, len(lines) * 20 + 60)
        
        graphic_elements.append({
            "type": "group",
            # "left": 40, 
            # "top": 40,
            "draggable": True, 
            "children": [
                {
                    "type": "rect",
                    "z": 100,
                    "shape": { "width": 280, "height": note_height },
                    "style": {
                        "fill": "#fff9c4", 
                        "stroke": "#fbc02d",
                        "lineWidth": 1,
                        "shadowBlur": 8,
                        "shadowColor": "rgba(0,0,0,0.2)"
                    }
                },
                {
                    "type": "text",
                    "z": 101,
                    "left": 15,
                    "top": 15,
                    "style": {
                        "text": f"📌 Governance Notes:\n\n{sticky_text}",
                        "fill": "#333",
                        "font": "13px Arial",
                        "lineHeight": 20
                    }
                }
            ]
        })

    # --- 4. RENDER FULL-WIDTH CHART ---
    options = {
        "tooltip": {"trigger": "item", "triggerOn": "mousemove"},
        "graphic": graphic_elements, 
        "series": [
            {
                "type": "tree",
                "data": [final_tree_data],
                "orient": "TB", 
                "top": "8%", "bottom": "8%", "left": "2%", "right": "2%",
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
    
    st_echarts(options=options, height="800px")

    # --- THE MAGIC HTML EXPORTER ---
    html_template = f"""
    <!DOCTYPE html>
    <html>
    <head>
        <meta charset="utf-8">
        <title>Org Flow Export</title>
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
            option.series[0].initialTreeDepth = -1; 
            chart.setOption(option);
        </script>
    </body>
    </html>
    """

    with st.sidebar:
        st.download_button(
            label="📥 Download Interactive HTML",
            data=html_template,
            file_name=f"Org_Flow_Export.html",
            mime="text/html"
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
