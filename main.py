# ------------------- CIF Charges Entry UI -------------------
import streamlit as st
import pandas as pd
import numpy as np
from io import BytesIO
import re
# ─────────────────────────────────────────────────────────────
# 0.  Page setup (MUST be first Streamlit call)
# ─────────────────────────────────────────────────────────────
st.set_page_config(
    page_title="CIF Charges Entry",
    page_icon="📦",
    layout="wide",
)

st.title("CIF Charges Entry UI")

# ─────────────────────────────────────────────────────────────
# 1.  Load exchange rates & currency list
# ─────────────────────────────────────────────────────────────
exchange_df = pd.read_excel("Data/Exchange Rates.xlsx")  # ← replace with your path

@st.cache_data
def get_currency_list(df):
    return sorted(df["Currency"].dropna().unique().tolist())

currency_options = get_currency_list(exchange_df)

# ─────────────────────────────────────────────────────────────
# 2.  Container‑level inputs
# ─────────────────────────────────────────────────────────────
with st.expander("📦 Container Information", expanded=True):
    c1, c2, c3 = st.columns(3)
    with c1:
        container_type = st.selectbox("Container Type", ["20 Standard", "40 Standard"])
        loadability    = st.text_input("Loadability (numeric)", "0")
    with c2:
        box_rate       = st.text_input("Box Rate (USD)", "0")
    with c3:
        origin_charges = st.text_input("Origin Charges (INR)", "0")

# ─────────────────────────────────────────────────────────────
# 3.  Session‑state setup for dynamic agents
# ─────────────────────────────────────────────────────────────
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

# ─────────────────────────────────────────────────────────────
# 4.  Data‑collector helpers
# ─────────────────────────────────────────────────────────────
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

# ─────────────────────────────────────────────────────────────
# 5.  Agent‑entry form UI
# ─────────────────────────────────────────────────────────────
def agent_form(agent_id: int):
    c1, c2 = st.columns([5, 1])
    with c1:
        st.text_input("Agent Name", key=f"agent_name_{agent_id}",
                      value=st.session_state.agent_names[agent_id])
    with c2:
        if st.button("❌", key=f"del_{agent_id}"):
            delete_agent(agent_id)
            st.rerun()

    st.markdown("**Destination Charges (CIF)**")
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

    st.markdown("**Rebates**")
    st.selectbox("Rebate Currency", currency_options,
                 key=f"{agent_id}_rebate_currency",
                 index=currency_options.index("USD") if "USD" in currency_options else 0)
    rebate_cols = st.columns(3)
    rebate_headers = ["Per CBM", "Per Ton", "Per BL"]
    for col, header in zip(rebate_cols, rebate_headers):
        col.markdown(f"**{header}**")

    r1, r2, r3 = st.columns(3)
    r1.text_input("", key=f"{agent_id}_rebate_cbm", label_visibility="collapsed")
    r2.text_input("", key=f"{agent_id}_rebate_ton", label_visibility="collapsed")
    r3.text_input("", key=f"{agent_id}_rebate_bl", label_visibility="collapsed")

# render each agent tab
tabs = st.tabs([f"Agent {aid}" for aid in st.session_state.agent_ids])
for tab, aid in zip(tabs, st.session_state.agent_ids):
    with tab:
        agent_form(aid)

# ─────────────────────────────────────────────────────────────
# 6.  Comparison engine
# ─────────────────────────────────────────────────────────────
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
            rebate_cbm = rebate_bl = 0.0
        else:
            r_cur = rebate_df.iloc[0]['Currency']
            r_rate = rate_map.get(r_cur, np.nan)
            rebate_cbm = rebate_df.iloc[0]['Per CBM'] * r_rate if not np.isnan(r_rate) else 0
            rebate_bl  = rebate_df.iloc[0]['Per BL']  * r_rate if not np.isnan(r_rate) else 0

        # --- charge totals
        totals = charge_df.apply(
            lambda row: row[money_cols] * rate_map.get(row['Currency'], np.nan),
            axis=1
        ).sum()
        tot_cbm, tot_bl = totals['Per CBM'], totals['Per BL']

        # --- build output row
        out = {"Agent Name": agent, "Remarks": remark}
        for n in range(1, 31):
            dest_chg = tot_bl + tot_cbm * n
            out[f"CBM {n}"] = (cost_per_wm * n) + dest_chg - (rebate_cbm * n) - rebate_bl
        rows_out.append(out)

    return pd.DataFrame(rows_out)

st.markdown("### 🛠️ Actions")
calc_btn, dl_placeholder = st.columns(2)

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
else:
    with dl_placeholder:
        st.caption("Run **Calculate** first to enable download.")