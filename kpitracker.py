# Turn an Excel Spreadsheet into an Interactive Dashboard using Python (Streamlit)
# Tutotial accessed @ https://www.youtube.com/watch?v=Sb0A9i6d320&t=194s

import pandas as pd
import plotly.express as px
import plotly
import streamlit as st

# Create Streamlit dashboard
st.set_page_config(page_title="Significant Event Dashboard",
                   page_icon=":horse:",
                   layout="wide")

@st.cache
def get_data_from_excel():
    df = pd.ExcelFile(r'C:\Users\galej\OneDrive - Tabcorp\Desktop\Raceday KPI Tracker Test.xlsx')
    # Select spreadsheet, skip first 2 rows and access colums C to P)
    df_sigevent = pd.read_excel(df, sheet_name='Significant Events',
                                    skiprows=1,
                                    usecols='C:P')



    print(df_sigevent.head())

    # Remove 00:00:00 values from date column
    df_sigevent['Date'] = pd.to_datetime(df_sigevent['Date'])
    df_sigevent['Date_New'] = df_sigevent['Date'].dt.date
    # remove unwanted columns
    df_sigevent = df_sigevent.drop(labels=['Date',
                                           'BRAVO',
                                           'Sell Code',
                                           'Race/s',
                                           'Summary',
                                           'Meeting',
                                           'Controller',
                                           'Controller2',
                                           'Controller3'],
                                   axis=1)


    # Rearrange Columnds
    df_sigevent = df_sigevent[['Date_New', 'Significant_Event', 'Sport_Code', 'Internal_External', 'Jurisdiction']]

    print(df_sigevent.to_string())
    return df_sigevent

df_sigevent = get_data_from_excel()

# #sidebar
st.sidebar.header("Filter here:")
significant_event = st.sidebar.multiselect(
    "Significant_Event",
    options=df_sigevent["Significant_Event"].unique(),
    default=df_sigevent["Significant_Event"].unique()
)
sport_code = st.sidebar.multiselect(
    "Sport_Code",
    options=df_sigevent["Sport_Code"].unique(),
    default=df_sigevent["Sport_Code"].unique()
)
int_ext = st.sidebar.multiselect(
    "Internal_External:",
    options=df_sigevent["Internal_External"].unique(),
    default=df_sigevent["Internal_External"].unique()
)
juris = st.sidebar.multiselect(
    "Jurisdiction:",
    options=df_sigevent["Jurisdiction"].unique(),
    default=df_sigevent["Jurisdiction"].unique()
)

# Create query selections for dataframe
df_selection = df_sigevent.query(
    "Significant_Event == @significant_event & "
    "Sport_Code == @sport_code &"
    "Internal_External == @int_ext &"
    "Jurisdiction == @juris"
)

st.title(":bar_chart: Significant Events")

# Display dataframe with query selectors on Streamlit
# st.dataframe(df_selection)
print("-----------------------")

# Create Sig Event total data frame
sig_event_dataframe = df_selection["Significant_Event"].value_counts()
print(sig_event_dataframe)

print("-----------------------")
print("-----------------------")
print(sig_event_dataframe[1:2])
print("-----------------------")
sig_total = sig_event_dataframe.sum(axis=0, skipna=True)
print("Sigevent total: ",  sig_total)
print("-----------------------")

# create Sport Code data frame
sport_code_dataframe = df_selection["Sport_Code"].value_counts()
print(sport_code_dataframe)
print("--------------")
# sport_code_dataframe.columns = ["Code", "Frequency"]
print(sport_code_dataframe.to_string())

print("-----------------------")


# SIGMA LEVEL
# Assign internal sig events to defects, races till date and  sig event opportunities per race
def sigma_level():
    internal_sig_events = 6

    defects = internal_sig_events
    races = 26198
    opportunity_per_race = 2

    # DPO = Determine how many opportunities for possible defects per race
    defects_per_opportunity = defects / (races * opportunity_per_race)

    # DPMO = Multiply defect per opportunity by 1,000,000
    defects_per_million_opportunities = defects_per_opportunity * 1000000

    # Sigma Level measured to 4.0
    if defects_per_million_opportunities <= 2:
        sigma = 6.00
        print("Sigma Level: ", sigma)
    elif defects_per_million_opportunities <= 5:
        sigma = 5.90
        print("Sigma Level: ", sigma)
    elif defects_per_million_opportunities <= 9:
        sigma = 5.80
        print("Sigma Level: ", sigma)
    elif defects_per_million_opportunities <= 21:
        sigma = 5.70
        print("Sigma Level: ", sigma)
    elif defects_per_million_opportunities <= 13:
        sigma = 5.60
        print("Sigma Level: ", sigma)
    elif defects_per_million_opportunities <= 32:
        sigma = 5.50
        print("Sigma Level: ", sigma)
    elif defects_per_million_opportunities <= 48:
        sigma = 5.40
        print("Sigma Level: ", sigma)
    elif defects_per_million_opportunities <= 72:
        sigma = 5.40
        print("Sigma Level: ", sigma)
    elif defects_per_million_opportunities <= 108:
        sigma = 5.20
        print("Sigma Level: ", sigma)
    elif defects_per_million_opportunities <= 159:
        sigma = 5.10
        print("Sigma Level: ", sigma)
    elif defects_per_million_opportunities <= 233:
        sigma = 5.00
        print("Sigma Level: ", sigma)
    elif defects_per_million_opportunities <= 337:
        sigma = 4.90
        print("Sigma Level: ", sigma)
    elif defects_per_million_opportunities <= 483:
        sigma = 4.80
        print("Sigma Level: ", sigma)
    elif defects_per_million_opportunities <= 687:
        sigma = 4.70
        print("Sigma Level: ", sigma)
    elif defects_per_million_opportunities <= 968:
        sigma = 4.60
        print("Sigma Level: ", sigma)
    elif defects_per_million_opportunities <= 1350:
        sigma = 4.50
        print("Sigma Level: ", sigma)
    elif defects_per_million_opportunities <= 1866:
        sigma = 4.40
        print("Sigma Level: ", sigma)
    elif defects_per_million_opportunities <= 2555:
        sigma = 4.30
        print("Sigma Level: ", sigma)
    elif defects_per_million_opportunities <= 3467:
        sigma = 4.20
        print("Sigma Level: ", sigma)
    elif defects_per_million_opportunities <= 4661:
        sigma = 4.10
        print("Sigma Level: ", sigma)
    elif defects_per_million_opportunities <= 6210:
        sigma = 4.0
        print("Sigma Level: ", sigma)
    else:
        print("Sigma Level out of control")
    return sigma

sigma = sigma_level()

# Organise Dashboard
left_column, middle1_column, middle2_collum, right_column = st.columns(4)
with left_column:
    st.subheader("Sigma Level: ")
    st.subheader(sigma)

with middle1_column:
    st.subheader("Sig Event Total: ")
    st.subheader(sig_total)

with middle2_collum:
    st.subheader("Domestic: ")

with right_column:
    st.subheader("International: ")




st.markdown("---")


# Sig Event Occurrence Bar Chart
sig_event_occurrence = px.bar(
    sig_event_dataframe,
    x=sig_event_dataframe,
    y=sig_event_dataframe.index,
    orientation="h",
    title="<b>Significant Events</b>",
    color_discrete_sequence=["#7EE3FF"] * len(sig_event_dataframe),
    template="plotly_white",
)
sig_event_occurrence.update_layout(
    plot_bgcolor="rgba(0,0,0,0)",
    xaxis=(dict(showgrid=False))
)


# Sport Code Occurrence Pie Chart
sport_code_occurrence = px.pie(
    sport_code_dataframe,
    values=sport_code_dataframe,
    names=sport_code_dataframe.index,
    title= "<b> Race Events </b>"
)
sport_code_occurrence.update_traces(
    textposition='inside',
    # textinfo='percent',
)

# sport_code_occurrence.update_layout(
#
# )

# st.plotly_chart(sport_code_occurrence)

l_column, r_column = st.columns(2)
with l_column:
    st.plotly_chart(sig_event_occurrence)

with r_column:
    st.plotly_chart(sport_code_occurrence)

# STREAMLIT CSS STYLING
css_style = """
            <style>
            /* MAIN */
                .css-10trblm {
                    color: #7EE3FF;
                }
                .stApp {
                    background-image: linear-gradient(to top, #025BAD, #074179);
                }
                .main-svg {
                    border: 5px solid;
                    border-radius: 10px;
                    border-color: #025BAD;
                }
                
            /* SIDEBAR */
                .css-163ttbj {
                    background-image: linear-gradient(to top, #025BAD, #074179);
                    border-right: 5px solid;
                    border-color: #025BAD;
                }
                .css-163ttbj h2 {
                    color: white;
                }
                .css-15tx938 {
                    color: white;
                }
                .st-c9 {
                    background-color: #16EFD2;
                    color: #085353;
                }
                .css-9s5bis {
                    color: #16EFD2;
                }
                
            /* HEADER & FOOTER */
                header {
                    visibility: hidden;
                }
                footer {
                    visibility: hidden;
                }
            </style>
            """
st.markdown(css_style, unsafe_allow_html=True)