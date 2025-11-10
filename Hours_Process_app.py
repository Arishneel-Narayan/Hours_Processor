import streamlit as st
import pandas as pd
import matplotlib.pyplot as plt
import matplotlib.colors as mcolors
import numpy as np
import datetime
import io

# ----------------------------
# CONFIGURATION & HARDCODED LISTS
# ----------------------------
# Page config
st.set_page_config(page_title="Technician Hours Tracker", layout="wide")

MANUAL_TECHS = [
    "SRIJAN", "EPELI_23", "AMITESHWAR_22", "KAUSHIK_23", "RATHNAYAKA",
    "NITUN_22", "ROMAN_01", "NIKLESH_25", "SUMIT_22", "NAND_22",
    "VILIAME", "ANISH", "PAUL_22", "SAILESH_22", "KAPIL_25",
    "NITENDRA_25", "GOSAI_25", "JOE_22", "RAHUL_24", "ROHIT",
]
MANUAL_TECHS = sorted(MANUAL_TECHS, key=str.upper)

TECHNICIAN_COL = "Technician"
WORK_ORDER_COL = "Work order for labor reporting"
HOURS_COL = "Labor reporting time (duration)"

# ----------------------------
# HELPER FUNCTIONS
# ----------------------------
def time_to_hours(t):
    """Converts various time formats (timedelta, str, int) to decimal hours."""
    if pd.isnull(t): return 0.0
    if isinstance(t, pd.Timedelta): return t.total_seconds() / 3600.0
    if isinstance(t, (pd.Timestamp, datetime.datetime)): return t.hour + t.minute / 60.0 + t.second / 3600.0
    if isinstance(t, datetime.time): return t.hour + t.minute / 60.0 + t.second / 3600.0
    if isinstance(t, (int, float, np.integer, np.floating)):
        val = float(t)
        return val * 24.0 if 0 <= val <= 1 else val
    try:
        td = pd.to_timedelta(t)
        return td.total_seconds() / 3600.0
    except Exception:
        pass
    try:
        dt = pd.to_datetime(t)
        return dt.hour + dt.minute / 60.0 + dt.second / 3600.0
    except Exception:
        pass
    try:
        return float(str(t))
    except Exception:
        return 0.0

def get_bar_color(hours, max_hours):
    """Determines bar color based on hours worked."""
    max_for_norm = max(max_hours, 20.0)
    if hours < 20:
        return "#FF0000" # Red
    else:
        # Orange to Yellow to Green gradient
        norm_val = (hours - 20) / (max_for_norm - 20) if (max_for_norm - 20) > 0 else 0
        cmap = mcolors.LinearSegmentedColormap.from_list("", ["orange", "yellow", "green"])
        return cmap(norm_val)

# ----------------------------
# STREAMLIT UI
# ----------------------------
st.title("📊 Technician Hours Tracker")

# Sidebar Configuration
with st.sidebar:
    st.header("Configuration")
    chart_title = st.text_input("Chart Title", value="Technician Hours Summary")
    sort_by_hours = st.checkbox("Sort by Total Hours (Descending)", value=True)
    
    st.subheader("Upload Data")
    uploaded_files = st.file_uploader(
        "Upload Occupation Excel files", 
        type=["xls", "xlsx"], 
        accept_multiple_files=True
    )
    
    with st.expander("View Target Technician List"):
        st.write(MANUAL_TECHS)

# Main App Logic
if uploaded_files:
    all_data = []
    for uploaded_file in uploaded_files:
        try:
            # Determine engine based on file extension
            file_ext = uploaded_file.name.split('.')[-1].lower()
            engine = "xlrd" if file_ext == "xls" else "openpyxl"
            
            df = pd.read_excel(uploaded_file, engine=engine)

            # Basic validation
            missing_cols = [col for col in [TECHNICIAN_COL, WORK_ORDER_COL, HOURS_COL] if col not in df.columns]
            if missing_cols:
                st.warning(f"Skipping {uploaded_file.name}: Missing columns {missing_cols}")
                continue

            # Clean and process
            df = df.dropna(subset=[TECHNICIAN_COL, HOURS_COL])
            df[TECHNICIAN_COL] = df[TECHNICIAN_COL].astype(str).str.strip().str.upper()
            df[HOURS_COL] = df[HOURS_COL].apply(time_to_hours).astype(float)
            all_data.append(df[[TECHNICIAN_COL, WORK_ORDER_COL, HOURS_COL]])
            
        except Exception as e:
            st.error(f"Failed to process {uploaded_file.name}: {e}")

    if all_data:
        # --- Combine and Summarize ---
        combined_df = pd.concat(all_data, ignore_index=True)

        grouped = combined_df.groupby(TECHNICIAN_COL).agg(
            Work_Orders_Completed=(WORK_ORDER_COL, "nunique"),
            Total_Hours_Worked=(HOURS_COL, "sum")
        ).reset_index()

        # Merge with manual list to ensure all target techs are present
        manual_df = pd.DataFrame({TECHNICIAN_COL: MANUAL_TECHS})
        summary = manual_df.merge(grouped, on=TECHNICIAN_COL, how="left")
        summary["Work_Orders_Completed"] = summary["Work_Orders_Completed"].fillna(0).astype(int)
        summary["Total_Hours_Worked"] = summary["Total_Hours_Worked"].fillna(0.0).round(2)

        if sort_by_hours:
            summary = summary.sort_values(by="Total_Hours_Worked", ascending=False, ignore_index=True)

        # --- Metrics Display ---
        col1, col2, col3, col4 = st.columns(4)
        col1.metric("Total Technicians", summary.shape[0])
        col2.metric("Total Hours", f"{summary['Total_Hours_Worked'].sum():.2f}")
        col3.metric("Avg Hours/Tech", f"{summary['Total_Hours_Worked'].mean():.2f}")
        col4.metric("Total Work Orders", summary["Work_Orders_Completed"].sum())

        # --- Plotting ---
        fig_height = max(6, 0.4 * len(summary))
        fig, ax = plt.subplots(figsize=(10, fig_height))

        y_pos = np.arange(len(summary))
        hours_data = summary["Total_Hours_Worked"].values
        
        # Dynamic max for color scaling
        current_max_hours = summary["Total_Hours_Worked"].max() if not summary.empty else 40
        colors = [get_bar_color(h, current_max_hours) for h in hours_data]

        bars = ax.barh(y_pos, hours_data, align="center", color=colors)
        ax.set_yticks(y_pos)
        ax.set_yticklabels(summary[TECHNICIAN_COL])
        ax.invert_yaxis()  # Highest first if sorted
        ax.set_xlabel("Total Hours Worked")
        ax.set_title(chart_title)

        x_limit = max(80, current_max_hours * 1.1)
        ax.set_xlim(0, x_limit)

        # Target line (40h)
        ax.axvline(40, color="blue", linestyle="--", linewidth=1)
        # Placed text slightly off the 40 line for visibility
        ax.text(40.5, len(summary)-0.5, "Target (40h)", color="blue", va="bottom", ha="left", fontsize=9, fontweight="bold")

        # Annotate bars
        for bar, val in zip(bars, hours_data):
            w = bar.get_width()
            label_x = w + (x_limit * 0.01) if w < x_limit * 0.9 else w - (x_limit * 0.02)
            ha_val = "left" if w < x_limit * 0.9 else "right"
            ax.text(label_x, bar.get_y() + bar.get_height() / 2, f"{val:.2f}", 
                   va="center", ha=ha_val, fontsize=9)

        st.pyplot(fig)
        # ----------------------------
        # DOWNLOAD SECTION (REVISED)
        # ----------------------------
        st.subheader("Download Outputs")

        # 1. Prepare Excel download
        excel_buffer = io.BytesIO()
        with pd.ExcelWriter(excel_buffer, engine='xlsxwriter') as writer:
            summary.to_excel(writer, index=False, sheet_name='Summary')
        
        # 2. Prepare Plot download
        img_buffer = io.BytesIO()
        # Save the figure to the in-memory buffer
        fig.savefig(img_buffer, format="png", dpi=150, bbox_inches="tight")

        # 3. Create columns for the buttons
        col1, col2 = st.columns(2)

        with col1:
            st.download_button(
                label="📥 Download Summary Excel",
                data=excel_buffer.getvalue(),
                file_name="technician_summary.xlsx",
                mime="application/vnd.ms-excel",
                use_container_width=True
            )
        
        with col2:
            st.download_button(
                label="🖼️ Download Graph PNG",
                data=img_buffer.getvalue(),
                file_name=f"{chart_title.replace(' ', '_').lower()}.png",
                mime="image/png",
                use_container_width=True
            )

    else:
        st.info("Uploaded files contained no valid data for processing.")
else:
    st.info("👆 Please upload Occupation Excel files in the sidebar to begin.")
