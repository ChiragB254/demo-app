import streamlit as st
import pandas as pd
from io import BytesIO
import base64
from datetime import datetime, timedelta

# Load the logo
logo_path = "tangerine-full.png"

def add_logo(logo_path):
    with open(logo_path, "rb") as image_file:
        img_str = base64.b64encode(image_file.read()).decode("utf-8")
    st.markdown(
        f"""
        <style>
        [data-testid="stSidebar"] {{
            background-image: url(data:image/png;base64,{img_str});
            background-size: 200px 60px;
            background-repeat: no-repeat;
            background-position: 20px 20px;
        }}
        </style>
        """, unsafe_allow_html=True
    )

add_logo(logo_path)

# Color theme (from logo color scheme)
theme_color = "#f47321"

st.markdown(
    f"""
    <style>
    .css-1offfwp {{
        background-color: {theme_color};
    }}
    </style>
    """, unsafe_allow_html=True
)

# Load the data
manager_data_df = pd.read_excel("manager_agent_data.xlsx")
agent_data_df = pd.read_excel("agent_data_sept_2024.xlsx")



# Ensure 'date' column is in datetime format
agent_data_df['date'] = pd.to_datetime(agent_data_df['date'], errors='coerce')

# Extract manager and agent data
managers = manager_data_df['Manager'].unique()
agents_data = {manager: manager_data_df[manager_data_df['Manager'] == manager]['Agent'].tolist() for manager in managers}

# Initialize session state for tracking total hours per agent
if 'total_hours' not in st.session_state:
    st.session_state['total_hours'] = {agent: 6.5 for manager in managers for agent in agents_data[manager]}

# User Interface
st.image(logo_path, width=300)
st.title("Tangerine Agent Hours Management")
st.markdown("---")

# Select Manager
selected_manager = st.selectbox("Select Manager Name", ["Select a Manager"] + list(managers))

# Select Date Range
today = datetime.today()
latest_date = today - timedelta(days=1)
date_range = st.date_input(
    "Select Date Range",
    [latest_date - timedelta(days=7), latest_date],
    min_value=datetime(2020, 1, 1),
    max_value=latest_date
)
start_date, end_date = date_range

# Display agents based on manager selection
if selected_manager != "Select a Manager":
    st.write(f"Agents under {selected_manager}:")
    agents = agents_data[selected_manager]
else:
    st.write("Select a manager to see agents.")
    agents = []

# Function to update total hours based on button clicks
def update_hours(agent, change):
    st.session_state['total_hours'][agent] += change

# Total hours per agent
for agent in agents:
    col1, col2, col3 = st.columns([2, 2, 6])

    with col1:
        st.write(agent)

    with col2:
        if st.button("+", key=f"add_{agent}"):
            update_hours(agent, 1)

        if st.button("-", key=f"sub_{agent}"):
            update_hours(agent, -1)

    with col3:
        default_hours = st.session_state['total_hours'][agent]
        new_hours = st.number_input(f"Total Hours ({agent})", value=default_hours, step=0.25)
        st.session_state['total_hours'][agent] = new_hours

# Generate Report
if st.button("Generate Report"):
    # Filter data by selected date range and agents
    filtered_data = agent_data_df[
        (agent_data_df['date'] >= pd.to_datetime(start_date)) &
        (agent_data_df['date'] <= pd.to_datetime(end_date)) &
        (agent_data_df['agent'].isin(agents))
    ]

    # Ensure necessary columns exist before accessing them
    if not filtered_data.empty:
        # Merge with manager_data_df to get the manager's name
        merged_data = pd.merge(filtered_data, manager_data_df[['Agent', 'Manager']], 
                                left_on='agent', right_on='Agent', how='left')

        # Calculate productivity for each agent
        merged_data['Total Hours'] = merged_data['agent'].apply(lambda agent: st.session_state['total_hours'][agent])
        merged_data['Productivity Score'] = (
            ((merged_data['alerts'] * 6) / 60) +
            ((merged_data['manual_alerts'] * 6) / 60) +
            ((merged_data['marked'] * 10) / 60)
        ) / merged_data['Total Hours']

        # Highlight rows with a productivity score less than 0.80
        def highlight_low_productivity(row):
            return ['background-color: red'] * len(row) if row['Productivity Score'] < 0.80 else [''] * len(row)

        # Generate the final report
        report_columns = ['date', 'Manager', 'agent', 'Total Hours', 'alerts', 'manual_alerts', 'marked', 'Productivity Score', 'Location']
        missing_columns = [col for col in report_columns if col not in merged_data.columns]

        if not missing_columns:
            report_data = merged_data[report_columns]
            styled_report = report_data.style.apply(highlight_low_productivity, axis=1)

            # Display the styled table in the app
            st.dataframe(styled_report)

            # Export to Excel
            def to_excel(df):
                output = BytesIO()
                with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                    df.to_excel(writer, sheet_name='Sheet1', index=False)

                    # Get the xlsxwriter workbook and worksheet objects.
                    workbook = writer.book
                    worksheet = writer.sheets['Sheet1']

                    # Apply conditional formatting for low productivity
                    format_red = workbook.add_format({'bg_color': '#FFC7CE', 'font_color': '#9C0006'})
                    worksheet.conditional_format('G2:G{}'.format(len(df) + 1), {'type': 'cell', 'criteria': '<', 'value': 0.80, 'format': format_red})

                processed_data = output.getvalue()
                return processed_data

            excel_data = to_excel(report_data)

            # Encode as base64 for download
            b64 = base64.b64encode(excel_data).decode()

            st.markdown(
                f'<a href="data:application/octet-stream;base64,{b64}" download="agent_hours_report.xlsx">Download Excel Report</a>',
                unsafe_allow_html=True
            )
        else:
            st.error(f"Missing columns in filtered data: {', '.join(missing_columns)}")
    else:
        st.error("No data found for the selected date range and agents.")
