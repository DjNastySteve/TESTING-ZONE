
from io import BytesIO
import xlsxwriter

def generate_agency_report(df, agency_name):
    output = BytesIO()
    workbook = None

    # Clean up blanks
    for col in ["Category 1", "Customer Name", "Sales Rep"]:
        df[col] = df[col].replace("(Blanks)", pd.NA)

    # Prep summary data
    total_sales = df[sales_column].sum()
    top_customers = df.groupby("Customer Name")[sales_column].sum().sort_values(ascending=False).head(10)
    bottom_customers = df.groupby("Customer Name")[sales_column].sum().sort_values(ascending=True).head(10)
    category_sales = df.groupby("Category 1")[sales_column].sum().sort_values(ascending=False)

    growth = df.groupby("Customer Name")[[sales_column, "Prior Sales"]].sum()
    growth["$ Growth"] = growth[sales_column] - growth["Prior Sales"]
    growth["% Growth"] = growth["$ Growth"] / growth["Prior Sales"].replace(0, pd.NA) * 100
    top_growth_dollars = growth.sort_values("$ Growth", ascending=False).head(3)
    top_growth_percent = growth[growth["% Growth"] > 0].sort_values("% Growth", ascending=False).head(3)
    top_decline_dollars = growth.sort_values("$ Growth", ascending=True).head(3)

    # Recap summary text
    diff_total = df[sales_column].sum() - df["Prior Sales"].sum()
    summary_lines = []
    summary_lines.append(f"""Hope everyone’s doing well! Here's your {agency_name} recap:\n""")
    if diff_total < 0:
        summary_lines.append(f"We ended the period down ${abs(diff_total):,.0f} vs last year. Still some wins to celebrate.\n")
    else:
        summary_lines.append(f"We ended the period up ${abs(diff_total):,.0f} over last year — great momentum!\n")
    summary_lines.append("\nTop dealers:")
    for dealer, row in top_growth_dollars.iterrows():
        summary_lines.append(f"- {dealer}: +${row['$ Growth']:,.0f}")
    summary_lines.append("\nDealers who pulled back:")
    for dealer, row in top_decline_dollars.iterrows():
        summary_lines.append(f"- {dealer}: -${abs(row['$ Growth']):,.0f}")
    summary_lines.append("""\n🔥 Product to plug: Rhyme Downlights. Sleek, simple, and a showroom favorite.
    Let’s lean into wins, check in on our quiet ones, and light it up ⚡""")
    summary_text = "\n".join(summary_lines)

    with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
        workbook = writer.book
        money_fmt = workbook.add_format({'num_format': '$#,##0'})
        bold = workbook.add_format({'bold': True})
        wrap_fmt = workbook.add_format({'text_wrap': True, 'valign': 'top'})

        # Summary
        df.to_excel(writer, index=False, sheet_name="Summary")
        ws_summary = writer.sheets["Summary"]
        for col_num, value in enumerate(df.columns):
            ws_summary.set_column(col_num, col_num, 18, money_fmt if "Sales" in value else None)

        # Auto Summary
        ws_auto = workbook.add_worksheet("Auto Summary")
        ws_auto.set_column("A:A", 100, wrap_fmt)
        ws_auto.write("A1", summary_text)

        # Deep Dive
        deep = workbook.add_worksheet("Deep Dive")
        deep.write("A1", "Best-Selling Product Categories", bold)
        for i, (cat, val) in enumerate(category_sales.items()):
            deep.write(i + 2, 0, cat)
            deep.write(i + 2, 1, val, money_fmt)

        deep.write("D1", "Top Dealers", bold)
        for i, (name, val) in enumerate(top_customers.items()):
            deep.write(i + 2, 3, name)
            deep.write(i + 2, 4, val, money_fmt)

        deep.write("G1", "Bottom Dealers", bold)
        for i, (name, val) in enumerate(bottom_customers.items()):
            deep.write(i + 2, 6, name)
            deep.write(i + 2, 7, val, money_fmt)

        deep.write("A20", "Client-Level Detail", bold)
        for col_idx, col in enumerate(df.columns):
            deep.write(21, col_idx, col, bold)
        for row_idx, row in df.iterrows():
            for col_idx, val in enumerate(row):
                deep.write(22 + row_idx, col_idx, val, money_fmt if "Sales" in df.columns[col_idx] else None)
        deep.autofilter(21, 0, 21 + len(df), len(df.columns) - 1)

    return output.getvalue()



import streamlit as st
import pandas as pd
import matplotlib.pyplot as plt

st.set_page_config(page_title="Proluxe Sales Dashboard", layout="wide")

# Declare view/territory filters early
view_option = st.sidebar.radio("📅 Select View", ["YTD", "MTD"])
territory = st.sidebar.radio("📌 Select Sales Manager", ["All", "Cole", "Jake", "Proluxe"])

@st.cache_data
def load_data(file):
    sales_df = pd.read_excel(file, sheet_name="Sales Data YTD")
    mtd_df = pd.read_excel(file, sheet_name="Monthly Goal Sales Data")

    cole_reps = ['609', '617', '621', '623', '625', '626']
    jake_reps = ['601', '614', '616', '619', '620', '622', '627']

    rep_map = pd.DataFrame({
        "REP": cole_reps + jake_reps + ['Home'],
        "Rep Name": ["Cole"] * len(cole_reps) + ["Jake"] * len(jake_reps) + ["Proluxe"]
    })

    for df in [sales_df, mtd_df]:
        df.columns = df.columns.str.strip()
        df["Sales Rep"] = df["Sales Rep"].astype(str)
        df[sales_column] = pd.to_numeric(df[sales_column], errors="coerce").fillna(0)

    return sales_df, mtd_df, rep_map

sales_df, mtd_df, rep_map = load_data("FY25.PLX.xlsx")

sales_df, mtd_df, rep_map = load_data("FY25.PLX.xlsx")

df = mtd_df.copy() if view_option == "MTD" else sales_df.copy()
df["Sales Rep"] = df["Sales Rep"].astype(str)
df = df.merge(rep_map, left_on="Sales Rep", right_on="REP", how="left")

rep_agency_mapping = {
"601": "New Era", "627": "Phoenix", "609": "Morris-Tait", "614": "Access", "616": "Synapse",
"617": "NuTech", "619": "Connected Sales", "620": "Frontline", "621": "ProAct", "622": "PSG",
"623": "LK", "625": "Sound-Tech", "626": "Audio Americas"
}

budgets = {
"Cole": 3769351.32, "Jake": 3027353.02, "Proluxe": 743998.29, "All": 7538702.63
}
agency_budget_mapping = {
"New Era": 890397.95, "Phoenix": 712318.36, "Morris-Tait": 831038.09, "Access": 237439.45,
"Synapse": 237439.45, "NuTech": 474878.91, "Connected Sales": 356159.18, "Frontline": 118719.73,
"ProAct": 385839.11, "PSG": 474878.91, "LK": 1187197.26, "Sound-Tech": 890397.95, "Audio Americas": 0
}

df["Agency"] = df["Sales Rep"].map(rep_agency_mapping)


agencies = sorted(df["Agency"].dropna().unique())
selected_agency = st.sidebar.selectbox("🏢 Filter by Agency", ["All"] + agencies)

df_filtered = df if territory == "All" else df[df["Rep Name"] == territory]
df_filtered = df_filtered if selected_agency == "All" else df_filtered[df_filtered["Agency"] == selected_agency]


if view_option == "MTD":
    budget_column = 'Proluxe FY25 Monthly Budget'
    sales_column = 'FY25 Current MTD'
    banner_html = "<div style='background-color:#111; padding:0.8em 1em; border-radius:0.5em; color:#DDD;'>📅 <b>Now Viewing:</b> <span style='color:#00FFAA;'>Month-To-Date</span> Performance</div>"
else:
    banner_html = "<div style='background-color:#111; padding:0.8em 1em; border-radius:0.5em; color:#DDD;'>📅 <b>Now Viewing:</b> <span style='color:#FFD700;'>Year-To-Date</span> Performance</div>"
st.markdown(banner_html, unsafe_allow_html=True)


total_sales = df_filtered[sales_column].sum()
budget = agency_budget_mapping.get(selected_agency, 0) if selected_agency != "All" else budgets.get(territory, 0)
percent_to_goal = (total_sales / budget * 100) if budget > 0 else 0
total_customers = df_filtered["Customer Name"].nunique()

col1, col2, col3, col4 = st.columns(4)
col1.metric("📦 Customers", f"{total_customers:,}")
col2.metric("💰 FY25 Sales", f"${total_sales:,.2f}")
col3.metric("🎯 FY25 Budget", f"${budget:,.2f}")
col4.metric("📊 % to Goal", f"{percent_to_goal:.1f}%")
st.progress(min(int(percent_to_goal), 100))

# Top & Bottom Customers
st.subheader("🏆 Top 10 Customers by Sales")
top10 = df_filtered.groupby(["Customer Name", "Agency"])[sales_column].sum().sort_values(ascending=False).head(10).reset_index()
top10["Sales ($)"] = top10[sales_column].apply(lambda x: f"${x:,.2f}")
st.table(top10[["Customer Name", "Agency", "Sales ($)"]])

st.subheader("🚨 Bottom 10 Customers by Sales")
bottom10 = df_filtered.groupby(["Customer Name", "Agency"])[sales_column].sum().sort_values().head(10).reset_index()
bottom10["Sales ($)"] = bottom10[sales_column].apply(lambda x: f"${x:,.2f}")
st.table(bottom10[["Customer Name", "Agency", "Sales ($)"]])

# Agency Bar Chart
st.subheader("🏢 Agency Sales Comparison")
agency_grouped = df_filtered.groupby("Agency")[sales_column].sum().sort_values()
fig, ax = plt.subplots(figsize=(10, 5))
bars = ax.barh(agency_grouped.index, agency_grouped.values, color="#00c3ff")
ax.bar_label(bars, fmt="%.0f", label_type="edge")
ax.set_xlabel("Current Sales ($)")
st.pyplot(fig)

# Download filtered data
st.subheader("📁 Export")
csv_export = df_filtered.to_csv(index=False)
st.download_button("⬇ Download Filtered Data as CSV", csv_export, "Filtered_FY25_Sales.csv", "text/csv")

# st.dataframe(df[["Sales Rep", "Rep Name", "Agency"]].drop_duplicates().head(10))
# --- Phase 3: Advanced Excel Export by Rep Agency ---
from io import BytesIO

# Final integrated export logic
selected_export_agency = st.sidebar.selectbox("Select Agency to Export", ["All"] + sorted(df_filtered["Agency"].dropna().unique()))
if st.sidebar.button("📥 Download Full Excel Report"):
    export_df = df_filtered if selected_export_agency == "All" else df_filtered[df_filtered["Agency"] == selected_export_agency]
    excel_data = generate_agency_report(export_df, selected_export_agency)
    st.download_button("Download Report", data=excel_data,
                      file_name=f"{selected_export_agency}_Proluxe_Report.xlsx",
                      mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")