import streamlit as st
import pandas as pd
import datetime
import re
import requests
import icalendar
import recurring_ical_events

# --- PAGE CONFIG ---
st.set_page_config(page_title="Holiday Finder", page_icon="📅", layout="wide")

st.title("📅 Holiday & Life Planner")

# --- URL PARAMETERS ---
query_params = st.query_params
default_name = query_params.get("name", "")
default_cal = query_params.get("cal", "")

# --- SIDEBAR ---
with st.sidebar:
    st.header("1. Settings")
    my_name = st.text_input("Your Name (Excel):", value=default_name, placeholder="Smith, John")
    
    st.header("2. Google Calendar")
    st.caption("Paste your 'Secret address in iCal format' here.")
    ical_url = st.text_input("iCal URL (.ics):", value=default_cal, placeholder="https://calendar.google.com/.../basic.ics")

    # Update URL params for bookmarking
    if my_name: st.query_params["name"] = my_name
    if ical_url: st.query_params["cal"] = ical_url

    st.header("3. Upload")
    uploaded_file = st.file_uploader("Drag & Drop Excel File", type=['xlsx'])

# --- HELPER FUNCTIONS ---

def clean_leave_type(text):
    """Removes spaces/digits for grouping."""
    if not isinstance(text, str): text = str(text)
    return re.sub(r'[\s\d]+', '', text)

def fetch_google_events(url, start_range, end_range):
    """Fetches iCal data, handles multi-day, and filters for 'AL ref'."""
    if not url: return {}
    
    try:
        response = requests.get(url)
        response.raise_for_status()
        calendar = icalendar.Calendar.from_ical(response.content)
        
        # Handle recurring events
        events = recurring_ical_events.of(calendar).between(start_range, end_range)
        
        events_map = {}
        
        for event in events:
            summary = str(event.get('SUMMARY', 'Busy'))
            
            # Extract Start
            dt_start = event.get('DTSTART').dt
            if isinstance(dt_start, datetime.datetime):
                start_date = dt_start.date()
            else:
                start_date = dt_start

            # Extract End
            if event.get('DTEND'):
                dt_end = event.get('DTEND').dt
                if isinstance(dt_end, datetime.datetime):
                    end_date = dt_end.date()
                else:
                    end_date = dt_end
            else:
                end_date = start_date + datetime.timedelta(days=1)

            # Safety loop for multi-day events
            if start_date == end_date:
                end_date = start_date + datetime.timedelta(days=1)

            current_curr = start_date
            while current_curr < end_date:
                if current_curr in events_map:
                    if summary not in events_map[current_curr]: 
                        events_map[current_curr].append(summary)
                else:
                    events_map[current_curr] = [summary]
                
                current_curr += datetime.timedelta(days=1)
        
        # --- LOGIC UPDATE: 'AL ref' FILTER ---
        final_map = {}
        for date_key, summary_list in events_map.items():
            # 1. Check if any event in the list contains "AL ref" (case insensitive)
            al_ref_events = [s for s in summary_list if "al ref" in s.lower()]
            
            if al_ref_events:
                # If found, ONLY show these specific events
                final_map[date_key] = "; ".join(al_ref_events)
            else:
                # Otherwise, show everything
                final_map[date_key] = "; ".join(summary_list)

        return final_map

    except Exception as e:
        st.error(f"Could not load Google Calendar: {e}")
        return {}

def extract_holidays(file, target_name):
    try:
        df = pd.read_excel(file, header=None)
    except Exception as e:
        st.error(f"Error reading Excel: {e}")
        return []

    my_booked_dates = []
    current_block_dates = {} 

    for index, row in df.iterrows():
        if len(row) < 3: continue
        col_b_val = row[1] 
        col_c_val = row[2]

        if isinstance(col_c_val, datetime.datetime):
            current_block_dates = {}
            for col_idx in range(2, len(row)):
                cell_value = row[col_idx]
                if isinstance(cell_value, datetime.datetime):
                    current_block_dates[col_idx] = cell_value
            continue 

        if current_block_dates and str(col_b_val).strip().lower() == target_name.strip().lower():
            for col_idx, date_val in current_block_dates.items():
                if col_idx < len(row): 
                    leave_code = row[col_idx]
                    code_str = str(leave_code).strip()
                    
                    if pd.notna(leave_code) and code_str not in ['', '0', 'nan', 'None', '0.0']:
                        my_booked_dates.append({
                            'Date': date_val, 
                            'Original Type': code_str,
                            'Type': clean_leave_type(code_str)
                        })
    return my_booked_dates

# --- EXECUTION ---
if uploaded_file and my_name:
    results = extract_holidays(uploaded_file, my_name)
    
    if results:
        df_results = pd.DataFrame(results)
        df_results['Date'] = pd.to_datetime(df_results['Date'])
        df_results = df_results.sort_values(by='Date')

        # --- GOOGLE CALENDAR INTEGRATION ---
        if ical_url:
            with st.spinner("Fetching Google Calendar events..."):
                min_date = df_results['Date'].min() - datetime.timedelta(days=7)
                max_date = df_results['Date'].max() + datetime.timedelta(days=7)
                
                gcal_events = fetch_google_events(ical_url, min_date, max_date)
                
                df_results['My Calendar'] = df_results['Date'].dt.date.map(gcal_events)
                df_results['My Calendar'] = df_results['My Calendar'].fillna("-")
        else:
            df_results['My Calendar'] = "No Link Provided"

        # --- SUMMARY ---
        col1, col2 = st.columns([1, 2])
        
        with col1:
            st.subheader(f"📊 Summary")
            summary_counts = df_results['Type'].value_counts().reset_index()
            summary_counts.columns = ['Leave Type', 'Days']
            st.dataframe(summary_counts, use_container_width=True, hide_index=True)

        with col2:
            st.subheader("📅 Schedule")
            display_df = df_results.copy()
            display_df['Date'] = display_df['Date'].dt.strftime('%d-%m-%Y')
            
            cols = ['Date', 'Original Type', 'My Calendar']
            st.dataframe(display_df[cols], use_container_width=True, hide_index=True)

        csv = df_results.to_csv(index=False).encode('utf-8')
        st.download_button(
            label="📥 Download Plan as CSV",
            data=csv,
            file_name=f"{my_name}_holiday_plan.csv",
            mime='text/csv',
        )
    else:
        st.warning("No holidays found.")

elif not uploaded_file:
    st.info("👆 Please upload the Excel file.")