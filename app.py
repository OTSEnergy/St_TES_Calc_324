"""
Steffes Analytics Tool - Cloud Web App Version
-------------------------------------------------
This is the Streamlit Cloud compatible version.
Instead of reading local files, it accepts an uploaded Excel file,
processes it entirely in memory using pandas & openpyxl,
and stores the results in Streamlit session state.
"""

import streamlit as st
import pandas as pd
import matplotlib.pyplot as plt
import seaborn as sns
import numpy as np

# Configure the browser tab title and the default wide page layout
st.set_page_config(layout="wide", page_title="ETS Simulation Viewer")
st.title("ETS Scenario Evaluation & Simulation")

def pull_excel_data(uploaded_file):
    with st.spinner("Crunching Excel data... this usually takes 30-60 seconds depending on file size."):
        try:
            def get_data(sheet_name, start_col, end_col):
                uploaded_file.seek(0)
                return pd.read_excel(uploaded_file, sheet_name=sheet_name, usecols=f"{start_col}:{end_col}", header=None, nrows=8765, engine="openpyxl").values

            # ---------------------------------------------------------
            # DATASET 1: The Baseline (Heat Pump Only)
            # ---------------------------------------------------------
            data_hp_only = get_data("Heat Pump Only Energy Calc ", "L", "AE")
            df_base = pd.DataFrame(data_hp_only)
            df_base = df_base.rename(columns={0: 'Timestamp', 1: 'Temp', 5: 'Whole House', 13: 'Baseline_HP', 16: 'Baseline_Backup'})
            for req_col in ['Timestamp', 'Temp', 'Whole House', 'Baseline_HP', 'Baseline_Backup']:
                if req_col not in df_base.columns: df_base[req_col] = 0
            df_base = df_base[['Timestamp', 'Temp', 'Whole House', 'Baseline_HP', 'Baseline_Backup']].iloc[1:8761]

            # ---------------------------------------------------------
            # DATASET 2: Scenario 1 (ETS Offset ER Heat)
            # ---------------------------------------------------------
            data_ets1 = get_data("ETS_OffsetERHeat", "L", "AF")
            df_ets1 = pd.DataFrame(data_ets1)
            df_ets1 = df_ets1.rename(columns={0: 'Timestamp', 7: 'ETS1_S', 14: 'ETS1_HP', 18: 'ETS1_Backup', 19: 'ETS1_ETS', 20: 'ETS1_AF'})
            for req_col in ['Timestamp', 'ETS1_HP', 'ETS1_Backup', 'ETS1_ETS', 'ETS1_S', 'ETS1_AF']:
                if req_col not in df_ets1.columns: df_ets1[req_col] = 0
            df_ets1 = df_ets1[['Timestamp', 'ETS1_HP', 'ETS1_Backup', 'ETS1_ETS', 'ETS1_S', 'ETS1_AF']].iloc[1:8761]
            
            for col in ['ETS1_HP', 'ETS1_Backup', 'ETS1_ETS', 'ETS1_S', 'ETS1_AF']:
                df_ets1[col] = pd.to_numeric(df_ets1[col], errors='coerce').fillna(0)
            df_ets1['ETS1_TotalSystemPower'] = df_ets1['ETS1_HP'] + df_ets1['ETS1_Backup'] + df_ets1['ETS1_ETS']
            df_ets1['ETS1_WholeHouse'] = df_ets1['ETS1_S'] + df_ets1['ETS1_AF']

            # ---------------------------------------------------------
            # DATASET 3: Scenario 2 (ETS Offset ER Heat + Peak HP)
            # ---------------------------------------------------------
            data_ets2 = get_data("ETS_OffsetERHeat+PeakHP", "L", "AF")
            df_ets2 = pd.DataFrame(data_ets2)
            df_ets2 = df_ets2.rename(columns={0: 'Timestamp', 7: 'ETS2_S', 14: 'ETS2_HP', 18: 'ETS2_Backup', 19: 'ETS2_ETS', 20: 'ETS2_AF'})
            for req_col in ['Timestamp', 'ETS2_HP', 'ETS2_Backup', 'ETS2_ETS', 'ETS2_S', 'ETS2_AF']:
                if req_col not in df_ets2.columns: df_ets2[req_col] = 0
            df_ets2 = df_ets2[['Timestamp', 'ETS2_HP', 'ETS2_Backup', 'ETS2_ETS', 'ETS2_S', 'ETS2_AF']].iloc[1:8761]
            
            for col in ['ETS2_HP', 'ETS2_Backup', 'ETS2_ETS', 'ETS2_S', 'ETS2_AF']:
                df_ets2[col] = pd.to_numeric(df_ets2[col], errors='coerce').fillna(0)
            df_ets2['ETS2_TotalSystemPower'] = df_ets2['ETS2_HP'] + df_ets2['ETS2_Backup'] + df_ets2['ETS2_ETS']
            df_ets2['ETS2_WholeHouse'] = df_ets2['ETS2_S'] + df_ets2['ETS2_AF']

            # ---------------------------------------------------------
            # METADATA EXTRACTION: Settings and Summary Context
            # ---------------------------------------------------------
            uploaded_file.seek(0)
            df_selections = pd.read_excel(uploaded_file, sheet_name="Model Selections ", header=None, engine="openpyxl")
            
            def safe_get(df, r, c, default=0.0):
                if r < len(df) and c < len(df.columns):
                    val = df.iloc[r, c]
                    return float(val) if pd.notna(val) else default
                return default

            hp_vals = [safe_get(df_selections, 10, 20), safe_get(df_selections, 10, 21), safe_get(df_selections, 10, 22)]
            ets_vals = [safe_get(df_selections, 11, 20), safe_get(df_selections, 11, 21), safe_get(df_selections, 11, 22)]
            peak_vals = [safe_get(df_selections, 12, 20), safe_get(df_selections, 12, 21), safe_get(df_selections, 12, 22)]
            
            peak_hours = []
            if len(df_selections) > 44:
                end_row = min(68, len(df_selections))
                peak_hours_raw = df_selections.iloc[44:end_row, 0].values
                for row_val in peak_hours_raw:
                    if pd.notna(row_val):
                        try:
                            peak_hours.append(int(float(row_val)))
                        except ValueError:
                            pass
            
            def extract_kv(start_row, end_row):
                kv = {}
                end_row = min(end_row, len(df_selections) - 1)
                if start_row > end_row: return kv
                for r in range(start_row, end_row + 1):
                    key = df_selections.iloc[r, 0] if 0 < len(df_selections.columns) else None
                    if pd.notna(key) and str(key).strip():
                        val = df_selections.iloc[r, 1] if 1 < len(df_selections.columns) else ""
                        val = val if pd.notna(val) else ""
                        if isinstance(val, float) and val.is_integer():
                            val = int(val)
                        kv[str(key).strip()] = str(val).strip()
                return kv

            settings = {}
            settings.update(extract_kv(2, 6))      # A3:B7
            settings.update(extract_kv(16, 18))    # A17:B19
            settings.update(extract_kv(35, 35))    # A36:B36
            
            summary_data = {
                "No ETS": {
                    "HP": float(hp_vals[0]) if pd.notna(hp_vals[0]) else 0.0,
                    "ETS": float(ets_vals[0]) if pd.notna(ets_vals[0]) else 0.0,
                    "Peak": float(peak_vals[0]) if pd.notna(peak_vals[0]) else 0.0
                },
                "ETS - Offset ER": {
                    "HP": float(hp_vals[1]) if pd.notna(hp_vals[1]) else 0.0,
                    "ETS": float(ets_vals[1]) if pd.notna(ets_vals[1]) else 0.0,
                    "Peak": float(peak_vals[1]) if pd.notna(peak_vals[1]) else 0.0
                },
                "ETS - Offset ER and Peak HP": {
                    "HP": float(hp_vals[2]) if pd.notna(hp_vals[2]) else 0.0,
                    "ETS": float(ets_vals[2]) if pd.notna(ets_vals[2]) else 0.0,
                    "Peak": float(peak_vals[2]) if pd.notna(peak_vals[2]) else 0.0
                },
                "Peak_Hours": peak_hours,
                "Model_Settings": settings
            }

            # FINAL CONCATENATION
            df_base['Temp'] = pd.to_numeric(df_base['Temp'], errors='coerce')
            df_base['Whole House'] = pd.to_numeric(df_base['Whole House'], errors='coerce').fillna(0)
            for col in ['Baseline_HP', 'Baseline_Backup']:
                df_base[col] = pd.to_numeric(df_base[col], errors='coerce').fillna(0)
            df_base['Baseline_TotalSystemPower'] = df_base['Baseline_HP'] + df_base['Baseline_Backup']

            df_final = pd.concat([df_base.reset_index(drop=True), 
                                  df_ets1.drop(columns=['Timestamp']).reset_index(drop=True),
                                  df_ets2.drop(columns=['Timestamp']).reset_index(drop=True)], axis=1)

            # Sanitize the Timestamp directly in python memory so the Date picker functions perfectly
            try:
                df_final['Timestamp'] = pd.to_datetime(df_final['Timestamp'], errors='coerce')
                # If there are completely corrupted rows that coerce to NaT (Not a Time), fill them with a dummy sequence
                if df_final['Timestamp'].isna().any():
                     df_final['Timestamp'] = pd.date_range(start='2023-01-01 00:00:00', periods=len(df_final), freq='H')
            except Exception:
                df_final['Timestamp'] = pd.date_range(start='2023-01-01 00:00:00', periods=len(df_final), freq='H')

            st.session_state['df'] = df_final
            st.session_state['summary'] = summary_data
            st.session_state['filename'] = uploaded_file.name
            
            st.success("Data extracted successfully!")

        except Exception as e:
            st.error(f"Error accessing Excel file: {e}. If the data didn't load, please screenshot this red box!")

# Initialize state
if 'df' not in st.session_state:
    st.session_state['df'] = None
if 'summary' not in st.session_state:
    st.session_state['summary'] = None
if 'filename' not in st.session_state:
    st.session_state['filename'] = None

df = st.session_state['df']
summary = st.session_state['summary']

# Show uploaded filename as a subtle sub-heading beneath the main title
if st.session_state['filename']:
    st.markdown(
        f"<p style='margin-top:-18px; color:gray; font-size:0.9em;'>📂 {st.session_state['filename']}</p>",
        unsafe_allow_html=True
    )

# ---------------------------------------------------------
# SIDEBAR CONFIGURATION
# ---------------------------------------------------------
st.sidebar.header("Data Connection")

uploaded_file = st.sidebar.file_uploader("Upload Excel Model (.xlsx)", type=['xlsx', 'xlsm'])

if uploaded_file is not None:
    if st.sidebar.button("Run Simulation / Pull Data"):
        pull_excel_data(uploaded_file)
        df = st.session_state.get('df')
        summary = st.session_state.get('summary')

# ---------------------------------------------------------
# GRAPHING & TABS ENGINE
# ---------------------------------------------------------

# If df is properly hydrated from CSV, we render the entire UI suite
if df is not None:
    min_temp_idx = None
    
    # Global Extents representing exactly whatever boundary the Dataframe returned (typically 12 months)
    global_min = df['Timestamp'].min().date()
    global_max = df['Timestamp'].max().date()
    default_start = global_min
    default_end = global_max
    
    # 1. PEAK WEEK CALCULATION 
    # Finds the single 168-hour (7-day) continuous block displaying the lowest average outdoor temperature
    if len(df) >= 168:
        rolling_avg = df['Temp'].rolling(window=168).mean()
        min_temp_idx = rolling_avg.idxmin()
        if pd.notna(min_temp_idx):
            default_end = df.loc[min_temp_idx, 'Timestamp'].date()
            default_start = df.loc[min_temp_idx - 167, 'Timestamp'].date()

    # 2. MODEL METADATA SIDEBAR
    # Writes out the settings parsed strictly from the Excel JSON payload
    if summary and 'Model_Settings' in summary:
        st.sidebar.markdown("---")
        st.sidebar.header("Model Settings")
        for k, v in summary['Model_Settings'].items():
            st.sidebar.markdown(f"**{k}:** {v}")

    # 3. DATE BOUNDARY UI
    # We dynamically lock the starting date to existing globals, then prompt for a specific numerical duration.
    # The application programmatically evaluates end dates to prevent invalid data index crashing out of bounds.
    st.sidebar.markdown("---")
    st.sidebar.header("Custom Date Range")
    
    user_start = st.sidebar.date_input("Start Date", value=default_start, min_value=global_min, max_value=global_max)
    
    default_days = (default_end - default_start).days
    if default_days <= 0: default_days = 7
    
    max_possible_days = (global_max - user_start).days
    if max_possible_days < 1: max_possible_days = 1
    if default_days > max_possible_days: default_days = max_possible_days
    
    num_days = st.sidebar.number_input("Number of Days", min_value=1, max_value=int(max_possible_days), value=int(default_days), step=1)
    
    user_end = user_start + pd.Timedelta(days=num_days)
    
    # Mask down the full 8760 sheet purely into the User's selected window
    user_mask = (df['Timestamp'].dt.date >= user_start) & (df['Timestamp'].dt.date <= user_end)
    user_df = df.loc[user_mask]

    # Render top-level application navigation tabs
    tab1, tab2, tab3, tab4, tab5, tab6, tab7 = st.tabs(["Peak Week", "User Selected Dates", "Heat Maps", "Winter Avg Profiles", "10 Coldest Winter Days", "Temp vs Load", "Data Tables"])

    
    # =========================================================
    # TAB 1: PEAK WEEK COMPARISON
    # =========================================================
    # Displays identical side-by-side time series of Power metrics bounded precisely 
    # over the 168 hours constituting the absolute lowest average thermal condition.
    with tab1:
        st.header("Peak Week Comparison")
        st.markdown("Comparing Total System Power during the 168-hour window with the lowest average temperature.")
        
        rolling_temp = df['Temp'].rolling(window=168).mean()
        end_idx = rolling_temp.idxmin()
        start_idx = end_idx - 167
        peak_week_df = df.iloc[start_idx:end_idx+1]
        
        col1, col2 = st.columns(2)
        
        with col1:
            st.subheader("Scenario 1: ETS Offset ER Heat")
            fig1, ax1 = plt.subplots(figsize=(8, 5), dpi=300)
            
            # Left Axis plotting Power in solid lines
            ax1.plot(peak_week_df['Timestamp'], peak_week_df['Baseline_TotalSystemPower'], label='Baseline', color='black', linewidth=1.5)
            ax1.plot(peak_week_df['Timestamp'], peak_week_df['ETS1_TotalSystemPower'], label='Scenario 1', color='#1f77b4', linewidth=1.5, alpha=0.9)
            ax1.set_ylabel('Power (kW)', fontsize=10)
            ax1.set_xlabel('Date/Time', fontsize=10)
            
            # Secondary Right Axis plotting Temperature dynamically mimicking the load curve
            ax1_temp = ax1.twinx()
            ax1_temp.plot(peak_week_df['Timestamp'], peak_week_df['Temp'], label='Temperature', color='gray', linestyle=':', linewidth=1.5, alpha=0.7)
            ax1_temp.set_ylabel('Temperature (°F)', fontsize=10, color='gray')
            ax1_temp.tick_params(axis='y', labelcolor='gray')
            
            # Consolidate mult-axis legends into a single box
            lines_1, labels_1 = ax1.get_legend_handles_labels()
            lines_2, labels_2 = ax1_temp.get_legend_handles_labels()
            ax1.legend(lines_1 + lines_2, labels_1 + labels_2, loc='upper right')
            
            ax1.grid(True, linestyle='--', alpha=0.7)
            ax1.tick_params(axis='x', rotation=45)
            fig1.tight_layout()
            st.pyplot(fig1)
            
        with col2:
            st.subheader("Scenario 2: ETS Offset ER Heat + Peak HP")
            fig2, ax2 = plt.subplots(figsize=(8, 5), dpi=300)
            ax2.plot(peak_week_df['Timestamp'], peak_week_df['Baseline_TotalSystemPower'], label='Baseline', color='black', linewidth=1.5)
            ax2.plot(peak_week_df['Timestamp'], peak_week_df['ETS2_TotalSystemPower'], label='Scenario 2', color='#d62728', linewidth=1.5, alpha=0.9)
            ax2.set_ylabel('Power (kW)', fontsize=10)
            ax2.set_xlabel('Date/Time', fontsize=10)
            
            ax2_temp = ax2.twinx()
            ax2_temp.plot(peak_week_df['Timestamp'], peak_week_df['Temp'], label='Temperature', color='gray', linestyle=':', linewidth=1.5, alpha=0.7)
            ax2_temp.set_ylabel('Temperature (°F)', fontsize=10, color='gray')
            ax2_temp.tick_params(axis='y', labelcolor='gray')
            
            # Combine legends
            lines_1, labels_1 = ax2.get_legend_handles_labels()
            lines_2, labels_2 = ax2_temp.get_legend_handles_labels()
            ax2.legend(lines_1 + lines_2, labels_1 + labels_2, loc='upper right')
            
            ax2.grid(True, linestyle='--', alpha=0.7)
            ax2.tick_params(axis='x', rotation=45)
            fig2.tight_layout()
            st.pyplot(fig2)


    with tab2:
        st.header("User Selected Dates")
        st.markdown(f"Custom view from {user_start} to {user_end}.")
        if not user_df.empty:
            col1, col2 = st.columns(2)
            with col1:
                st.subheader("Baseline vs. Scenario 1 (ETS Offset ER Heat)")
                fig_pw1, ax_pw1 = plt.subplots(figsize=(8, 5), dpi=300)
                ax_pw1.plot(user_df['Timestamp'], user_df['Baseline_TotalSystemPower'], color='black', linewidth=2, label='Baseline')
                ax_pw1.plot(user_df['Timestamp'], user_df['ETS1_TotalSystemPower'], color='#1f77b4', linewidth=2, linestyle='--', label='Scenario 1')
                ax_pw1.set_ylabel('Total System Power (kW)', color='black', fontsize=10)
                ax_pw1.tick_params(axis='y', labelcolor='black')
                
                ax_temp1 = ax_pw1.twinx()
                ax_temp1.plot(user_df['Timestamp'], user_df['Temp'], color='gray', linewidth=1.5, linestyle=':', alpha=0.7, label='S.Temp (°F)')
                ax_temp1.set_ylabel('Temperature (°F)', color='gray', fontsize=10)
                ax_temp1.tick_params(axis='y', labelcolor='gray')
                
                lines_pw1, labels_pw1 = ax_pw1.get_legend_handles_labels()
                lines_temp1, labels_temp1 = ax_temp1.get_legend_handles_labels()
                ax_pw1.legend(lines_pw1 + lines_temp1, labels_pw1 + labels_temp1, loc='upper left')
                
                plt.setp(ax_pw1.xaxis.get_majorticklabels(), rotation=45, ha="right")
                ax_pw1.grid(True, linestyle='--', alpha=0.6)
                fig_pw1.tight_layout()
                st.pyplot(fig_pw1)

            with col2:
                st.subheader("Baseline vs. Scenario 2 (ETS Offset ER Heat + Peak HP)")
                fig_pw2, ax_pw2 = plt.subplots(figsize=(8, 5), dpi=300)
                ax_pw2.plot(user_df['Timestamp'], user_df['Baseline_TotalSystemPower'], color='black', linewidth=2, label='Baseline')
                ax_pw2.plot(user_df['Timestamp'], user_df['ETS2_TotalSystemPower'], color='#d62728', linewidth=2, linestyle='--', label='Scenario 2')
                ax_pw2.set_ylabel('Total System Power (kW)', color='black', fontsize=10)
                ax_pw2.tick_params(axis='y', labelcolor='black')
                
                ax_temp2 = ax_pw2.twinx()
                ax_temp2.plot(user_df['Timestamp'], user_df['Temp'], color='gray', linewidth=1.5, linestyle=':', alpha=0.7, label='S.Temp (°F)')
                ax_temp2.set_ylabel('Temperature (°F)', color='gray', fontsize=10)
                ax_temp2.tick_params(axis='y', labelcolor='gray')
                
                lines_pw2, labels_pw2 = ax_pw2.get_legend_handles_labels()
                lines_temp2, labels_temp2 = ax_temp2.get_legend_handles_labels()
                ax_pw2.legend(lines_pw2 + lines_temp2, labels_pw2 + labels_temp2, loc='upper left')
                
                plt.setp(ax_pw2.xaxis.get_majorticklabels(), rotation=45, ha="right")
                ax_pw2.grid(True, linestyle='--', alpha=0.6)
                fig_pw2.tight_layout()
                st.pyplot(fig_pw2)
        else:
            st.warning("No data found for the selected date range.")

    # =========================================================
    # TAB 3: HEATMAP RENDERING
    # =========================================================
    # Converts continuous 8760 timeline data into a strictly wrapped 365x24 grid.
    with tab3:
        st.header("8760-Hour Heat Maps")
        st.markdown("Displays the total system power (kW) for each case across the entire 8760 hours.")
        
        # Guard clause - prevents crashes if dealing with short incomplete datasets
        if len(df) >= 8760:
            cases = {
                "Baseline (Heat Pump Only)": "Baseline_TotalSystemPower",
                "ETS Offset ER Heat": "ETS1_TotalSystemPower",
                "ETS Offset ER Heat + Peak HP": "ETS2_TotalSystemPower"
            }
            
            # Hardcoded consistent coloring thresholds between 0-10 so graphics match visually
            global_min = 0
            global_max = 10
            
            for title, col_name in cases.items():
                fig, ax = plt.subplots(figsize=(15, 4), dpi=300)
                
                # Slicing the first 8760 records precisely, and folding it into rows of 24 cols. 
                # Transposing (.T) so Days sit horizontally and Hours build vertically
                data_matrix = df[col_name].values[:8760].reshape(365, 24).T
                cmap = sns.color_palette("rocket_r", as_cmap=True) 
                
                # heatmap with max cap at 10, and extend triangle on the colorbar
                ax_hm = sns.heatmap(data_matrix, cmap=cmap, ax=ax, vmin=global_min, vmax=global_max, cbar_kws={'label': 'Power (kW)', 'extend': 'max'})
                
                # Update colorbar to demonstrate capping above 10
                cbar = ax_hm.collections[0].colorbar
                cbar.set_ticks([0, 2, 4, 6, 8, 10])
                cbar.set_ticklabels(['0', '2', '4', '6', '8', '10+'])
                
                # Visual tweaks: 0 Hour (Midnight) normally draws at top, invert puts it at bottom
                ax.invert_yaxis()
                ax.set_title(f'Heat Map: {title}', fontsize=12, fontweight='bold')
                ax.set_ylabel('Hour of Day (0-23)')
                ax.set_xlabel('Day of Year (1-365)')
                
                # Step axis markers strictly to chunks of 30 days vs numbering every single column (unreadable)
                ax.set_xticks(np.arange(0, 365, 30))
                ax.set_xticklabels(np.arange(1, 365, 30), rotation=0)
                st.pyplot(fig)
        else:
            st.warning("Needs exactly 8760 rows to generate 365x24 heatmaps.")
    
    # =========================================================
    # TAB 4: WINTER AVERAGE PROFILES
    # =========================================================
    with tab4:
        st.header("Winter Hourly Average Profiles")
        st.markdown("Average power profile for each hour of the day during Winter months (Dec, Jan, Feb).")
        
        # Index filtering dropping all rows outside of Dec(12) Jan(1) Feb(2)
        winter_months = [12, 1, 2]
        winter_df = df[df['Timestamp'].dt.month.isin(winter_months)]
        
        if not winter_df.empty:
            # Grouping by mathematically averaging the Power columns for "Hour 0", "Hour 1", etc...
            winter_avg = winter_df.groupby(winter_df['Timestamp'].dt.hour)[['Baseline_TotalSystemPower', 'ETS1_TotalSystemPower', 'ETS2_TotalSystemPower']].mean().reset_index()
            
            col1, col2 = st.columns(2)
            
            with col1:
                st.subheader("Scenario 1: ETS Offset ER Heat")
                fig1, ax1 = plt.subplots(figsize=(8, 5), dpi=300)
                ax1.fill_between(winter_avg['Timestamp'], winter_avg['Baseline_TotalSystemPower'], color='black', alpha=0.3, label='Baseline')
                ax1.fill_between(winter_avg['Timestamp'], winter_avg['ETS1_TotalSystemPower'], color='#1f77b4', alpha=0.5, label='Scenario 1')
                ax1.plot(winter_avg['Timestamp'], winter_avg['Baseline_TotalSystemPower'], color='black', linewidth=1.5)
                ax1.plot(winter_avg['Timestamp'], winter_avg['ETS1_TotalSystemPower'], color='#1f77b4', linewidth=1.5)
                ax1.set_ylabel('Average Power (kW)', fontsize=10)
                ax1.set_xlabel('Hour of Day', fontsize=10)
                ax1.set_xticks(range(0, 24))
                ax1.legend()
                ax1.grid(True, linestyle='--', alpha=0.5)
                fig1.tight_layout()
                st.pyplot(fig1)

            with col2:
                st.subheader("Scenario 2: ETS Offset ER Heat + Peak HP")
                fig2, ax2 = plt.subplots(figsize=(8, 5), dpi=300)
                ax2.fill_between(winter_avg['Timestamp'], winter_avg['Baseline_TotalSystemPower'], color='black', alpha=0.3, label='Baseline')
                ax2.fill_between(winter_avg['Timestamp'], winter_avg['ETS2_TotalSystemPower'], color='#d62728', alpha=0.5, label='Scenario 2')
                ax2.plot(winter_avg['Timestamp'], winter_avg['Baseline_TotalSystemPower'], color='black', linewidth=1.5)
                ax2.plot(winter_avg['Timestamp'], winter_avg['ETS2_TotalSystemPower'], color='#d62728', linewidth=1.5)
                ax2.set_ylabel('Average Power (kW)', fontsize=10)
                ax2.set_xlabel('Hour of Day', fontsize=10)
                ax2.set_xticks(range(0, 24))
                ax2.legend()
                ax2.grid(True, linestyle='--', alpha=0.5)
                fig2.tight_layout()
                st.pyplot(fig2)
        else:
            st.warning("Insufficient winter data found in simulation_results.csv")

    # =========================================================
    # TAB 5: TOP 10 COLDEST DAYS PROFILES
    # =========================================================
    with tab5:
        st.header("10 Coldest Winter Days Average Profiles")
        st.markdown("Average power profile for the 10 absolute coldest days during the Winter months (Dec, Jan, Feb).")
        
        winter_months = [12, 1, 2]
        winter_df_all = df[df['Timestamp'].dt.month.isin(winter_months)].copy()
        
        if not winter_df_all.empty:
            winter_df_all['Date'] = winter_df_all['Timestamp'].dt.date
            # Mathematically resolve daily rolling average temperature, sort the top 10 lowest values
            daily_temp = winter_df_all.groupby('Date')['Temp'].mean().reset_index()
            coldest_10_dates = daily_temp.sort_values('Temp').head(10)['Date']
            
            # Repull those exact 10 days out of the main masterframe into isolated frame
            coldest_10_df = df[df['Timestamp'].dt.date.isin(coldest_10_dates)]
            cold_avg = coldest_10_df.groupby(coldest_10_df['Timestamp'].dt.hour)[['Baseline_TotalSystemPower', 'ETS1_TotalSystemPower', 'ETS2_TotalSystemPower']].mean().reset_index()
            
            col1, col2 = st.columns(2)
            
            with col1:
                st.subheader("Scenario 1: ETS Offset ER Heat")
                fig1, ax1 = plt.subplots(figsize=(8, 5), dpi=300)
                ax1.fill_between(cold_avg['Timestamp'], cold_avg['Baseline_TotalSystemPower'], color='black', alpha=0.3, label='Baseline')
                ax1.fill_between(cold_avg['Timestamp'], cold_avg['ETS1_TotalSystemPower'], color='#1f77b4', alpha=0.5, label='Scenario 1')
                ax1.plot(cold_avg['Timestamp'], cold_avg['Baseline_TotalSystemPower'], color='black', linewidth=1.5)
                ax1.plot(cold_avg['Timestamp'], cold_avg['ETS1_TotalSystemPower'], color='#1f77b4', linewidth=1.5)
                ax1.set_ylabel('Average Power (kW)', fontsize=10)
                ax1.set_xlabel('Hour of Day', fontsize=10)
                ax1.set_xticks(range(0, 24))
                ax1.legend()
                ax1.grid(True, linestyle='--', alpha=0.5)
                fig1.tight_layout()
                st.pyplot(fig1)

            with col2:
                st.subheader("Scenario 2: ETS Offset ER Heat + Peak HP")
                fig2, ax2 = plt.subplots(figsize=(8, 5), dpi=300)
                ax2.fill_between(cold_avg['Timestamp'], cold_avg['Baseline_TotalSystemPower'], color='black', alpha=0.3, label='Baseline')
                ax2.fill_between(cold_avg['Timestamp'], cold_avg['ETS2_TotalSystemPower'], color='#d62728', alpha=0.5, label='Scenario 2')
                ax2.plot(cold_avg['Timestamp'], cold_avg['Baseline_TotalSystemPower'], color='black', linewidth=1.5)
                ax2.plot(cold_avg['Timestamp'], cold_avg['ETS2_TotalSystemPower'], color='#d62728', linewidth=1.5)
                ax2.set_ylabel('Average Power (kW)', fontsize=10)
                ax2.set_xlabel('Hour of Day', fontsize=10)
                ax2.set_xticks(range(0, 24))
                ax2.legend()
                ax2.grid(True, linestyle='--', alpha=0.5)
                fig2.tight_layout()
                st.pyplot(fig2)
        else:
            st.warning("Insufficient winter data found in simulation_results.csv")

    # =========================================================
    # TAB 6: SCATTER PLOTS (TEMP VS LOAD SCATTERING)
    # =========================================================
    with tab6:
        st.header("Whole House Power vs. Outdoor Temperature")
        st.markdown("Scatter plots showing the relationship between outdoor temperature and the *Whole House* load (not just HVAC/ETS).")
        
        # Verifies custom extraction columns natively built into dataframes
        if 'ETS1_WholeHouse' in df.columns and 'ETS2_WholeHouse' in df.columns:
            # Establishing mathematical bounding boxes (Extents) mapping highest and lowest
            # data points across the entirety of the visual scope for strict aesthetic continuity across the 3 subplots.
            temp_min, temp_max = df['Temp'].min(), df['Temp'].max()
            wh_min = min(df['Whole House'].min(), df['ETS1_WholeHouse'].min(), df['ETS2_WholeHouse'].min())
            wh_max = max(df['Whole House'].max(), df['ETS1_WholeHouse'].max(), df['ETS2_WholeHouse'].max())
            
            # Add a visual 5% buffer box
            temp_buf = (temp_max - temp_min) * 0.05 if pd.notna(temp_max) else 10
            wh_buf = (wh_max - wh_min) * 0.05 if pd.notna(wh_max) else 10
            
            x_lims = (temp_min - temp_buf, temp_max + temp_buf)
            y_lims = (wh_min - wh_buf, wh_max + wh_buf)
            
            col1, col2, col3 = st.columns(3)
            
            # Evaluate JSON cache for designated 'Peak Hour Settings', 
            # building boolean masks so graphs conditionally highlight Peak dots heavily 
            peak_hours = summary.get('Peak_Hours', []) if summary else []
            if peak_hours:
                is_peak = df['Timestamp'].dt.hour.isin(peak_hours)
                df_peak = df[is_peak]
                df_offpeak = df[~is_peak]
            else:
                df_peak = pd.DataFrame(columns=df.columns)
                df_offpeak = df
            
            with col1:
                st.subheader("Baseline (Heat Pump Only)")
                fig_base, ax_base = plt.subplots(figsize=(6, 5), dpi=300)
                # Plot transparent overlapping data
                ax_base.scatter(df_offpeak['Temp'], df_offpeak['Whole House'], facecolors='none', edgecolors='black', alpha=0.3, s=10, label='Off-Peak')
                if not df_peak.empty:
                     ax_base.scatter(df_peak['Temp'], df_peak['Whole House'], color='black', alpha=0.6, edgecolors='none', s=10, label='On-Peak')
                
                ax_base.set_xlim(x_lims)
                ax_base.set_ylim(y_lims)
                ax_base.set_xlabel('Temperature (°F)')
                ax_base.set_ylabel('Whole House Power (kW)')
                ax_base.grid(True, linestyle='--', alpha=0.5)
                ax_base.legend()
                fig_base.tight_layout()
                st.pyplot(fig_base)

            with col2:
                st.subheader("Scenario 1: ETS Offset ER Heat")
                fig_ets1, ax_ets1 = plt.subplots(figsize=(6, 5), dpi=300)
                ax_ets1.scatter(df_offpeak['Temp'], df_offpeak['ETS1_WholeHouse'], facecolors='none', edgecolors='#1f77b4', alpha=0.3, s=10, label='Off-Peak')
                if not df_peak.empty:
                     ax_ets1.scatter(df_peak['Temp'], df_peak['ETS1_WholeHouse'], color='#1f77b4', alpha=0.6, edgecolors='none', s=10, label='On-Peak')
                     
                ax_ets1.set_xlim(x_lims)
                ax_ets1.set_ylim(y_lims)
                ax_ets1.set_xlabel('Temperature (°F)')
                ax_ets1.set_ylabel('Whole House Power (kW)')
                ax_ets1.grid(True, linestyle='--', alpha=0.5)
                ax_ets1.legend()
                fig_ets1.tight_layout()
                st.pyplot(fig_ets1)

            with col3:
                st.subheader("Scenario 2: ETS Offset ER Heat + Peak HP")
                fig_ets2, ax_ets2 = plt.subplots(figsize=(6, 5), dpi=300)
                ax_ets2.scatter(df_offpeak['Temp'], df_offpeak['ETS2_WholeHouse'], facecolors='none', edgecolors='#d62728', alpha=0.3, s=10, label='Off-Peak')
                if not df_peak.empty:
                     ax_ets2.scatter(df_peak['Temp'], df_peak['ETS2_WholeHouse'], color='#d62728', alpha=0.6, edgecolors='none', s=10, label='On-Peak')
                
                ax_ets2.set_xlim(x_lims)
                ax_ets2.set_ylim(y_lims)
                ax_ets2.set_xlabel('Temperature (°F)')
                ax_ets2.set_ylabel('Whole House Power (kW)')
                ax_ets2.grid(True, linestyle='--', alpha=0.5)
                ax_ets2.legend()
                fig_ets2.tight_layout()
                st.pyplot(fig_ets2)
        else:
            st.info("Whole House metrics for ETS scenarios not found. Please click 'Pull Latest Data from Excel' to fetch the updated S and AF columns.")

    # =========================================================
    # TAB 7: EXPLORATORY DATA TABLES (Sanity Checking)
    # =========================================================
    with tab7:
        st.header("First Week Data (First 168 Hours)")
        
        # Hard limits table to first 168 items to prevent breaking HTML limits in browser DOM
        first_week = df.head(168).copy()
        
        with st.expander("Scenario 1: Baseline (Heat Pump Only)"):
            st.dataframe(first_week[['Timestamp', 'Temp', 'Whole House', 'Baseline_HP', 'Baseline_Backup', 'Baseline_TotalSystemPower']], use_container_width=True)
            
        with st.expander("Scenario 2: ETS Offset ER Heat"):
            st.dataframe(first_week[['Timestamp', 'Temp', 'ETS1_HP', 'ETS1_Backup', 'ETS1_ETS', 'ETS1_TotalSystemPower']], use_container_width=True)
            
        with st.expander("Scenario 3: ETS Offset ER Heat + Peak HP"):
            st.dataframe(first_week[['Timestamp', 'Temp', 'ETS2_HP', 'ETS2_Backup', 'ETS2_ETS', 'ETS2_TotalSystemPower']], use_container_width=True)

# Guard logic against clean boots with missing configuration payload
else:
    st.info("No simulation data found. Please set your inputs in the sidebar and run the simulation process.")
