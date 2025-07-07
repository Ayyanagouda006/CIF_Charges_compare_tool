# ------------------- CIF Charges Entry UI -------------------
# LCL Destination Charges Comparison Calculator with Save & History Tabs
# ----------------------------------------------------------------------
import streamlit as st
import pandas as pd
import numpy as np
import os
import re
from io import BytesIO

# ----------------------------------------------------------------------
# 0.  Page setup (MUST be first Streamlit call)
# ----------------------------------------------------------------------
st.set_page_config(
    page_title="LCL Destination Charges Comparison Calculator",
    page_icon="📦",
    layout="wide",
)

# Ensure data directories exist ------------------------------------------------
DATA_DIR = "Data"
SAVED_DIR = os.path.join(DATA_DIR, "Saved")
EXCHANGE_PATH = os.path.join(DATA_DIR, "Exchange Rates.xlsx")
os.makedirs(SAVED_DIR, exist_ok=True)

# ---------------------------------------------------------------------------------------
# 1.  Helper – Load & cache exchange‑rate table
# ---------------------------------------------------------------------------------------
@st.cache_data(show_spinner=False)
def load_exchange_rates() -> pd.DataFrame:
    """Returns the exchange‑rate DataFrame directly from Excel."""
    return pd.read_excel(EXCHANGE_PATH)

def save_exchange_rates(df: pd.DataFrame):
    df.to_excel(EXCHANGE_PATH, index=False)

exchange_df = load_exchange_rates()

@st.cache_data(show_spinner=False)
def get_currency_list(df: pd.DataFrame):
    return sorted(df["Currency"].dropna().unique().tolist())

currency_options = get_currency_list(exchange_df)

# ==============================================================================
# MAIN NAVIGATION TABS
# ==============================================================================
main_tabs = st.tabs(["📊 Comparison Calculator", "📂 Saved Comparisons", "💱 Exchange Rates"])

# ==============================================================================
# TAB 1: COMPARISON CALCULATOR
# ==============================================================================
with main_tabs[0]:
    st.title("LCL Destination Charges Comparison Calculator")

    # ------------------------------------------------------------------
    # 2.  Container‑level inputs
    # ------------------------------------------------------------------
    with st.expander("📦 Container Information", expanded=True):
        c1, c2, c3 = st.columns(3)
        with c1:
            container_type = st.selectbox("Container Type", ["20 Standard", "40 Standard"])
            loadability    = st.text_input("Loadability (numeric)", "0")
        with c2:
            box_rate       = st.text_input("Box Rate (USD)", "0")
        with c3:
            origin_charges = st.text_input("Origin Charges (INR)", "0")

    # ------------------------------------------------------------------
    # 3.  Session‑state setup for dynamic agents
    # ------------------------------------------------------------------
    if "agent_ids" not in st.session_state:
        st.session_state.agent_ids  = [1]
        st.session_state.agent_names = {1: "Agent 1"}

    if st.button("➕ Add Agent"):
        new_id = max(st.session_state.agent_ids) + 1
        st.session_state.agent_ids.append(new_id)
        st.session_state.agent_names[new_id] = f"Agent {new_id}"

    def delete_agent(agent_id):
        st.session_state.agent_ids.remove(agent_id)
        st.session_state.agent_names.pop(agent_id, None)

    # ------------------------------------------------------------------
    # 4.  Data‑collector helpers
    # ------------------------------------------------------------------
    def extract_agent_data() -> pd.DataFrame:
        rows = []
        for agent_id in st.session_state.agent_ids:
            agent_name = st.session_state.get(f"agent_name_{agent_id}", f"Agent {agent_id}")

            # destination charges rows 1‑8
            for i in range(1, 9):
                rows.append({
                    "Agent Name": agent_name,
                    "Description": st.session_state.get(f"{agent_id}_desc_{i}", ""),
                    "Currency":    st.session_state.get(f"{agent_id}_currency_{i}", ""),
                    "Per CBM":     st.session_state.get(f"{agent_id}_cbm_{i}", ""),
                    "Per Ton":     st.session_state.get(f"{agent_id}_ton_{i}", ""),
                    "Minimum":     st.session_state.get(f"{agent_id}_min_{i}", ""),
                    "Maximum":     st.session_state.get(f"{agent_id}_max_{i}", ""),
                    "Per BL":      st.session_state.get(f"{agent_id}_bl_{i}", "")
                })

            # remarks row
            rows.append({
                "Agent Name": agent_name,
                "Description": "Remarks",
                "Currency":    st.session_state.get(f"{agent_id}_desc_9", ""),
                "Per CBM": "", "Per Ton": "", "Minimum": "", "Maximum": "", "Per BL": ""
            })

            # rebate row
            rows.append({
                "Agent Name": agent_name,
                "Description": "Rebate",
                "Currency": st.session_state.get(f"{agent_id}_rebate_currency", ""),
                "Per CBM":   st.session_state.get(f"{agent_id}_rebate_cbm", ""),
                "Per Ton":   st.session_state.get(f"{agent_id}_rebate_ton", ""),
                "Minimum": "", "Maximum": "",
                "Per BL":    st.session_state.get(f"{agent_id}_rebate_bl", "")
            })

        df = pd.DataFrame(rows)
        # keep only non‑empty descriptions
        return df[df["Description"].fillna("").str.strip() != ""]

    # ------------------------------------------------------------------
    # 5.  Agent‑entry form UI
    # ------------------------------------------------------------------
    def agent_form(agent_id: int):
        c1, c2 = st.columns([5, 1])
        with c1:
            st.text_input("***Agent Name***", key=f"agent_name_{agent_id}",
                          value=st.session_state.agent_names[agent_id])
        with c2:
            if st.button("❌", key=f"del_{agent_id}"):
                delete_agent(agent_id)
                st.rerun()

        st.markdown("***Destination Charges (CIF)***")
        head_cols = st.columns([3, 1, 1, 1, 1, 1, 1])
        for col, h in zip(head_cols,
                          ["Charge Head", "Currency", "Per CBM", "Per Ton",
                           "Minimum", "Maximum", "Per BL"]):
            col.markdown(f"**{h}**")

        for i in range(1, 9):
            cols = st.columns([3, 1, 1, 1, 1, 1, 1])
            cols[0].text_input("", key=f"{agent_id}_desc_{i}",
                               label_visibility="collapsed", placeholder=f"Charge Head {i}")
            cols[1].selectbox("", currency_options, key=f"{agent_id}_currency_{i}",
                              label_visibility="collapsed",
                              index=currency_options.index("USD")
                              if "USD" in currency_options else 0)
            cols[2].text_input("", key=f"{agent_id}_cbm_{i}",  label_visibility="collapsed")
            cols[3].text_input("", key=f"{agent_id}_ton_{i}",  label_visibility="collapsed")
            cols[4].text_input("", key=f"{agent_id}_min_{i}",  label_visibility="collapsed")
            cols[5].text_input("", key=f"{agent_id}_max_{i}",  label_visibility="collapsed")
            cols[6].text_input("", key=f"{agent_id}_bl_{i}",   label_visibility="collapsed")

        st.text_input("Charge Head 9 Notes", key=f"{agent_id}_desc_9",
                      placeholder="If Cartons, …")

        st.markdown("***Rebates***")

        rebate_cols = st.columns(4)
        rebate_headers = ["Currency", "Per CBM", "Per Ton", "Per BL"]
        for col, header in zip(rebate_cols, rebate_headers):
            col.markdown(f"**{header}**")

        r1, r2, r3, r4 = st.columns(4)
        r1.selectbox("", currency_options,
                     key=f"{agent_id}_rebate_currency",
                     label_visibility="collapsed",
                     index=currency_options.index("USD") if "USD" in currency_options else 0)
        r2.text_input("", key=f"{agent_id}_rebate_cbm", label_visibility="collapsed")
        r3.text_input("", key=f"{agent_id}_rebate_ton", label_visibility="collapsed")
        r4.text_input("", key=f"{agent_id}_rebate_bl", label_visibility="collapsed")

    # render each agent tab
    tabs = st.tabs([f"Agent {aid}" for aid in st.session_state.agent_ids])
    for tab, aid in zip(tabs, st.session_state.agent_ids):
        with tab:
            agent_form(aid)

    # ------------------------------------------------------------------
    # 6.  Comparison engine
    # ------------------------------------------------------------------
    def agent_compare(df, exchange_df, loadability, box_rate, origin_charge):
        money_cols = ['Per CBM', 'Per Ton', 'Minimum', 'Maximum', 'Per BL']

        # -------- 1. INR → USD factor & cost/WM
        inr_rate   = exchange_df.loc[exchange_df['Currency'].eq('INR'),
                                     'Exchange Rate to USD'].astype(float).squeeze()
        origin_usd = origin_charge * inr_rate
        cost_per_wm = (box_rate + origin_usd) / loadability

        # -------- 2. Clean numeric columns
        df[money_cols] = (df[money_cols]
                          .replace(r'^\s*$', np.nan, regex=True)
                          .apply(pd.to_numeric, errors='coerce')
                          .fillna(0))

        # -------- 3. Currency → USD map
        rate_map = dict(zip(exchange_df['Currency'],
                            exchange_df['Exchange Rate to USD'].astype(float)))
        rate_map.setdefault('USD', 1.0)

        # -------- 4. Per‑agent calculation
        rows_out = []
        for agent, grp in df.groupby('Agent Name', sort=False):
            rebate_df  = grp[grp['Description'] == 'Rebate']
            remarks_df = grp[grp['Description'] == 'Remarks']
            charge_df  = grp[~grp['Description'].isin(['Rebate', 'Remarks'])]

            remark = remarks_df['Currency'].iloc[0] if not remarks_df.empty else ""

            # --- rebate figures
            if rebate_df.empty:
                rebate_cbm = rebate_per_ton = rebate_bl = 0.0
            else:
                r_cur = rebate_df.iloc[0]['Currency']
                r_rate = rate_map.get(r_cur, np.nan)
                rebate_cbm = rebate_df.iloc[0]['Per CBM'] * r_rate if not np.isnan(r_rate) else 0
                rebate_bl  = rebate_df.iloc[0]['Per BL']  * r_rate if not np.isnan(r_rate) else 0
                rebate_per_ton = rebate_df.iloc[0]['Per Ton']  * r_rate if not np.isnan(r_rate) else 0

            # --- charge totals
            totals = charge_df.apply(
                lambda row: row[money_cols] * rate_map.get(row['Currency'], np.nan),
                axis=1
            ).sum()
            tot_cbm, tot_bl, tot_ton = totals['Per CBM'], totals['Per BL'], totals['Per Ton']

            # --- build output row
            out = {"Agent Name": agent, "Remarks": remark}
            for n in range(1, 31):
                tpc = tot_cbm * n
                tpt = tot_ton * (n/2)
                if tpc > tpt:
                    con = tpc
                    rcon = rebate_cbm * n
                else:
                    con = tpt
                    rcon = rebate_per_ton * (n/2)
                dest_chg = tot_bl + con
                out[f"CBM {n}"] = (cost_per_wm * n) + dest_chg - (rcon) - rebate_bl
            rows_out.append(out)

        return pd.DataFrame(rows_out)

    st.markdown("### 🛠️ Actions")
    calc_btn, dl_placeholder, save_placeholder = st.columns([1, 1, 1])

    # 7‑A Calculate
    if calc_btn.button("🧮 Calculate"):
        try:
            load_f   = float(loadability)
            box_f    = float(box_rate)
            origin_f = float(origin_charges)
        except ValueError:
            st.error("Loadability, Box Rate, and Origin Charges must be numeric.")
            st.stop()

        in_df  = extract_agent_data()
        out_df = agent_compare(in_df, exchange_df, load_f, box_f, origin_f)

        st.session_state["container_info"] = pd.DataFrame({
            "Field": ["Container Type", "Loadability", "Box Rate (USD)", "Origin Charges (INR)"],
            "Value": [container_type, load_f, box_f, origin_f]
        })
        st.session_state["last_input_df"]  = in_df
        st.session_state["last_result_df"] = out_df

        st.success("Calculation complete.")
        st.dataframe(out_df)

    # 7‑B Download (only if data exists)
    def to_safe_sheet(name: str) -> str:
        # Trim to 31 chars, remove forbidden chars
        name = re.sub(r"[\[\]\*:/\\?]", "", name)[:31]
        return name or "Sheet"

    if all(k in st.session_state for k in ("container_info", "last_input_df", "last_result_df")):
        with dl_placeholder:
            buf = BytesIO()
            with pd.ExcelWriter(buf, engine="xlsxwriter") as writer:
                st.session_state["container_info"].to_excel(writer, sheet_name="Info", index=False)

                for agent, grp in st.session_state["last_input_df"].groupby("Agent Name", sort=False):
                    grp.to_excel(writer, sheet_name=to_safe_sheet(agent), index=False)

                st.session_state["last_result_df"].to_excel(writer, sheet_name="Comparison", index=False)

            buf.seek(0)
            st.download_button(
                "📥 Download Excel",
                data=buf.getvalue(),
                file_name="cif_charge_report.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

        # 7‑C Save Feature ----------------------------------------------------
        with save_placeholder:
            if "save_mode" not in st.session_state:
                st.session_state.save_mode = False

            if not st.session_state.save_mode:
                if st.button("💾 Save Comparison"):
                    st.session_state.save_mode = True
                    st.rerun()
            else:
                st.text_input("Enter a name for this comparison:", key="save_filename")
                confirm_col, cancel_col = st.columns([1, 1])
                if confirm_col.button("✅ Confirm Save"):
                    filename = st.session_state.get("save_filename", "").strip()
                    if not filename:
                        st.error("Filename cannot be empty.")
                    else:
                        safe_name = re.sub(r"[^A-Za-z0-9 _-]", "", filename).replace(" ", "_")
                        file_path = os.path.join(SAVED_DIR, f"{safe_name}.xlsx")
                        if os.path.exists(file_path):
                            st.warning("A file with this name already exists and will be overwritten.")
                        with pd.ExcelWriter(file_path, engine="xlsxwriter") as writer:
                            st.session_state["container_info"].to_excel(writer, sheet_name="Info", index=False)
                            for agent, grp in st.session_state["last_input_df"].groupby("Agent Name", sort=False):
                                grp.to_excel(writer, sheet_name=to_safe_sheet(agent), index=False)
                            st.session_state["last_result_df"].to_excel(writer, sheet_name="Comparison", index=False)
                        st.success(f"Comparison saved as '{safe_name}.xlsx' in the Saved folder.")
                        st.session_state.save_mode = False
                if cancel_col.button("❌ Cancel"):
                    st.session_state.save_mode = False
                    st.rerun()
    else:
        with dl_placeholder:
            st.caption("Run **Calculate** first to enable download and save options.")

# ==============================================================================
# TAB 2: SAVED COMPARISONS
# ==============================================================================
with main_tabs[1]:
    st.title("📂 Saved Comparisons")

    saved_files = [f for f in os.listdir(SAVED_DIR) if f.lower().endswith(".xlsx")]

    if not saved_files:
        st.info("No saved comparisons found. Return to the first tab, perform a calculation, and save it.")
    else:
        selected_file = st.selectbox("Select a saved comparison to view:", saved_files)
        if selected_file:
            file_path = os.path.join(SAVED_DIR, selected_file)

            # ------------------------------------------------------------------
            # PREVIEW SELECTED WORKBOOK
            # ------------------------------------------------------------------
            with pd.ExcelFile(file_path) as xls:
                sheet_names = xls.sheet_names
                view_tabs = st.tabs(sheet_names)
                for sheet, t in zip(sheet_names, view_tabs):
                    with t:
                        df_sheet = pd.read_excel(xls, sheet_name=sheet)
                        st.dataframe(df_sheet)

            # ------------------------------------------------------------------
            # DOWNLOAD & DELETE ACTIONS
            # ------------------------------------------------------------------
            act_dl, act_del = st.columns(2)
            with act_dl:
                with open(file_path, "rb") as fp:
                    st.download_button(
                        label="📥 Download this comparison",
                        data=fp.read(),
                        file_name=selected_file,
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    )
            with act_del:
                if "delete_mode" not in st.session_state:
                    st.session_state.delete_mode = False

                if not st.session_state.delete_mode:
                    if st.button("🗑️ Delete this comparison"):
                        st.session_state.delete_mode = True
                        st.rerun()
                else:
                    st.warning("Are you sure you want to delete this file? This action cannot be undone.")
                    c_yes, c_no = st.columns([1, 1])
                    if c_yes.button("✅ Yes, delete"):
                        try:
                            os.remove(file_path)
                            st.success(f"'{selected_file}' has been deleted.")
                            st.session_state.delete_mode = False
                            st.rerun()
                        except Exception as e:
                            st.error(f"Error deleting file: {e}")
                    if c_no.button("❌ Cancel"):
                        st.session_state.delete_mode = False
                        st.rerun()

with main_tabs[2]:
    st.title("💱 Edit Exchange Rates")
    st.caption("You can update or add new exchange rates. Click save to apply changes.")

    edited_df = st.data_editor(
        exchange_df,
        num_rows="dynamic",
        use_container_width=True,
        key="exchange_editor"
    )

    if st.button("💾 Save Exchange Rates"):
        if "Currency" in edited_df.columns and "Exchange Rate to USD" in edited_df.columns:
            try:
                edited_df["Exchange Rate to USD"] = pd.to_numeric(edited_df["Exchange Rate to USD"])
                save_exchange_rates(edited_df)
                st.success("Exchange rates saved successfully. Please refresh to see changes.")
                st.cache_data.clear()
            except Exception as e:
                st.error(f"Error saving exchange rates: {e}")
        else:
            st.error("Please ensure 'Currency' and 'Exchange Rate to USD' columns exist.")