# @Email:  miroslav.mrnak@gmail.com
# @Project:  Kapa Dashboard Streamlit

# INFO
# Každý rok je potřeba doplnit aktuální svátky!!! Mohlo by se řešit knihovnou

import datetime
import calendar
import pandas as pd  # pip install pandas openpyxl
import plotly.express as px  # pip install plotly-express
import streamlit as st  # pip install streamlit

# emojis: https://www.webfx.com/tools/emoji-cheat-sheet/
st.set_page_config(
    page_title="KaPa STIHL Dashboard",
    page_icon=":bar_chart:",
    layout="wide"
    )
st.header('KaPa STIHL 2022 // Version 0.01 ')

## ---- KONSTANTY ----

# zdrojová data
KAPA_FOLDER = "Kapa_aktualni\\"
KAPA_KON_FILE = "STIHL_Midd_term capacity planning_KON.xlsm"
KAPA_BER_FILE = "STIHL_Midd_term capacity planning_BER.xlsm"
KAPA_H2_FILE = "STIHL_Midd_term capacity planning_H2.xlsm"
KAPA_ETS_FILE = "STIHL_Midd_term capacity planning_ETS.xlsm"

#AKKODIS Barvy - zatím nepoužito
yellow = '#fdb71a'
blue = '#001b33'

# Groupy v plánování
groups = {
    'konstrukce': ['CAD', 'KMNG', 'PCAD'],
    'vypocty': ['FEM', 'PFEM', 'BHMNG', 'BKMNG', 'CFD', 'PCFD', 'IT', 'PIT'],
    'h2': ['H2', 'H2MN', 'PH2', 'HR'],
    'ets': ['EE', 'PEE'],
}

# AR dle jednotlivých skupin
activity_rates = {
    'CAD': 0.98,
    'PCAD': 0.70,
    'KMNG': 0.66,
    'FEM': 0.98,
    'PFEM': 0.70,
    'BKMNG': 0.70,
    'CFD': 0.98,
    'PCFD': 0.70,
    'IT': 0.98,
    'PIT': 0.70,
    'BHMNG': 0.45,
    'H2': 0.98,
    'PH2': 0.70,
    'H2MN': 0.70,
    'HR': 0.98,
    'EE': 0.98,
    'PEE': 0.70
}

# filtrování
holidays = ['Urlaub', 'Urlaub?']
holidays_0_5 = ['1/2Urlaub', '1/2Urlaub?']
sholidays = ['SUrlaub', 'Surlaub']
sholidays_0_5 = ['1/2SUrlaub', '1/2Surlaub']
courses = ['X-S/H', 'X-Intro']
xhours = ['X-RFQ', 'X-Admin', 'X-ISO', 'X-Uni', 'X-Stud', 'X-Tren', 'X-INO', 'X-Kapaplan', 'X-Leadership', 'X-DEVEL', 'X-Train-Stud']
sick = ['Krank', 'Bebe', 'BeBe']
pun = ['PUN']
abwesend = ['Abwesend']
non_project = holidays + holidays_0_5 + sholidays + sholidays_0_5 + courses + xhours + sick + pun + abwesend

# svátky CZ - nutno vyhodit dny z KAPA
public_holidays_cz = [
    '2022-01-01',
    '2022-04-15',
    '2022-04-18',
    '2022-05-01',
    '2022-05-08',
    '2022-07-05',
    '2022-07-06',
    '2022-09-28',
    '2022-10-28',
    '2022-11-17',
    '2022-12-24',
    '2022-12-25',
    '2022-12-26'
]
# převod svátku do formátu date
public_holidays_cz = [datetime.date.fromisoformat(day) for day in public_holidays_cz]

## ---- POMOCNÉ FUNKCE ----

# ---- Tvorba DF ----
# základní DF vygenerovaný z Excel dokumentu
@st.cache
def get_df(excel_file, group):
    '''
    excel_file - pd.ExcelFile(KAPA_FOLDER+FILE)
    group - jedna ze čtyř možnosti: 'konstrukce', 'vypocty', 'h2', 'ets'
    '''
    excel = pd.ExcelFile(KAPA_FOLDER+excel_file)
    df_all = excel.parse("Mittelfristige Kapaplanung 2022", usecols="A,C,E:I,R:NR", skiprows=6)
    df_all = df_all.rename(
            columns={'Unnamed: 0': "par",
                     'Unnamed: 2': "Name",
                     'Unnamed: 4': "Rest Urlaub",
                     'Unnamed: 5': "Überwiesene Urlaub",
                     'Unnamed: 6': "Urlaub Anspruch",
                     'Unnamed: 7': "Gepl.Urlaub",
                     'Unnamed: 8': "Koeficient"
                    }).set_index('Name')

    df_all = df_all.loc[df_all.index.dropna()]
    df_all = df_all[df_all['par'].isin(groups[group])]

    df_par = df_all[["par", "Rest Urlaub", "Überwiesene Urlaub", "Urlaub Anspruch", "Gepl.Urlaub", "Koeficient"]]
    df = df_all.drop(["par", "Rest Urlaub", "Überwiesene Urlaub", "Urlaub Anspruch", "Gepl.Urlaub", "Koeficient"], axis="columns")

    df = df.T
    #df = df.astype('category') # nastavení údajů v DF na typ 'category' --> snad je výpočet a filtrování rychlejší
    df.index = pd.to_datetime(df.index) # nastavení indexu na datetime
    # vyfiltrování jen pracovních dnů a svátku z DF
    df = df[df.index.weekday < 5] # vyfiltrování jen pracovních dnů
    df = df[df.index.map((lambda x: x.date() not in public_holidays_cz))]

    return df, df_par

# Připravený DF pro určité období
def df_timestamp(df, start, end):
    '''
    start, end musí být v datetime.date formatu 
    '''
    return df.loc[pd.Timestamp(start) : pd.Timestamp(end)]

# Připravený DF pro vybrané kategorie (CAD, PCAD, KMNG...)
def df_category(df, df_par, fte_filter):
    '''
    fte_filter - seznam vybraných kategorii 
    '''
    return df[[name for name in df_par.index if df_par.loc[name]['par'] in fte_filter]]

# množina projektů a hodnot z DF
def get_projects(df):
    values = set()
    for name in df.columns:
        for value in df[name].unique():
            values.add(value)
    return values

# spočítej dannou hodnotu v DF
def count_values(df, values):
    return df.applymap((lambda x: x in values)).sum().sum()*8

# ---- Výpočty hodnot z DF ----
# výpočet disponibilní kapy
def dispo_hours(df):
    df_proj = df.applymap((lambda x: x not in holidays+holidays_0_5+abwesend))
    return df_proj.sum().sum()*8 # přepočet na hodiny

# výpočet dovolených
def urlaub_hours(df):
    df_urlaub = df.applymap((lambda x: x in holidays))
    df_urlaub_0_5 = df.applymap((lambda x: x in holidays_0_5))
    return df_urlaub.sum().sum()*8 + df_urlaub_0_5.sum().sum()*4 # přepočet na hodiny

# výpočet S-ek
def s_urlaub_hours(df):
    df_s_urlaub = df.applymap((lambda x: x in sholidays))
    df_s_urlaub_0_5 = df.applymap((lambda x: x in sholidays_0_5))
    return df_s_urlaub.sum().sum()*8 + df_s_urlaub_0_5.sum().sum()*4 # přepočet na hodiny

# výpočet Nemocí, BEBE
def krank_hours(df):
    df_krank = df.applymap((lambda x: x in sick))
    return df_krank.sum().sum()*8

# výpočet školení
def training_hours(df):
    df_trainig = df.applymap((lambda x: x in courses))
    return df_trainig.sum().sum()*8

# výpočet x-hodin
def x_hours(df):
    df_x = df.applymap((lambda x: x in xhours))
    return df_x.sum().sum()*8

# výpočet PUN
def pun_hours(df):
    df_pun = df.applymap((lambda x: x in pun))
    return df_pun.sum().sum()*8

# výpočet produktivních hodin včetně zahrnutí AR
def productive_hours(df, df_par):
    df_proj = df.applymap((lambda x: x not in non_project)) # namapování všechn hodnot, které nejsou v non_project
    for name in df_proj.columns:
        df_proj[name] = df_proj[name]*activity_rates[df_par.loc[name, 'par']] # přenásobení AR koeficientem
    return df_proj.sum().sum()*8 # přepočet na hodiny

# výpočet produktivních hodin včetně zahrnutí AR od těd do konce měsíce
def productive_hours_forecast(df, df_par):
    # vyselektování dnů od dnes do konce měsíce
    today_day = datetime.date.today()
    last_day = datetime.date.fromisoformat(f'{today_day.year}-{today_day.month}-{calendar.monthrange(today_day.year, today_day.month)[-1]}')
    return productive_hours((df_timestamp(df, today_day, last_day)), df_par)

# Všechny neproduktivní hodiny: disponibilní kapa "mínus" prodatelné hodiny
def unproductive_hours(df, df_par):
    return dispo_hours(df) - productive_hours(df, df_par)

# ---- statistiky po KW ----
# počet pracovních dní v jednotlivých KW
def work_days(df):
    work_days = df.index.isocalendar()
    work_days['count'] = 1
    work_days = work_days.groupby('week').sum()['count']
    return work_days

# disponibilní FTE v jednotlivých KW
def dispo_hours_week(df):
    week = df.index.isocalendar()['week']
    df_proj = df.applymap((lambda x: x not in holidays+holidays_0_5+abwesend)).groupby(week).sum().T.sum()
    df_proj /= work_days(df) 
    return df_proj

# zúčtovatelné FTE v jednotlivých KW
def productive_hours_week(df, df_par):
    week = df.index.isocalendar()['week']
    df_proj = df.applymap((lambda x: x not in non_project)) # namapování všechn hodnot, které nejsou v non_project
    for name in df_proj.columns:
        df_proj[name] = df_proj[name]*activity_rates[df_par.loc[name, 'par']] # přenásobení AR koeficientem
    df_proj = df_proj.groupby(week).sum().T.sum()
    df_proj /= work_days(df)
    return df_proj

# celkový DF se statistikami dle KW
def df_statistics(df, df_par):
    df_stat = pd.DataFrame(
        {
        'Disponibilní' : dispo_hours_week(df),
        'Produktivní' : productive_hours_week(df, df_par),
        'Neproduktivní': dispo_hours_week(df) - productive_hours_week(df, df_par),
        'AR': productive_hours_week(df, df_par) / dispo_hours_week(df)
        }
    )
    return df_stat

# ---- Kontrolní mechanismy ----
def check_duplicities(df):
    for name in df.columns:
        if list(df.columns).count(name) > 1:
            return name
    return False

## ---- Statické Proměnné ----
kapa_kon, kapa_kon_par = get_df(KAPA_KON_FILE, 'konstrukce')
kapa_ber, kapa_ber_par = get_df(KAPA_BER_FILE, 'vypocty')
kapa_h2, kapa_h2_par = get_df(KAPA_H2_FILE, 'h2')
kapa_ets, kapa_ets_par = get_df(KAPA_ETS_FILE, 'ets')

kapa = pd.concat([kapa_kon, kapa_ber, kapa_h2, kapa_ets], axis=1)
kapa_par = pd.concat([kapa_kon_par, kapa_ber_par, kapa_h2_par, kapa_ets_par], axis=0)

## ---- StreamLit ----
# ---- SIDEBAR ----
st.sidebar.image('akkodis.png')

cad_type = st.sidebar.multiselect(
    "Select the CAD - FTE Type:",
    options=groups['konstrukce'],
    default=groups['konstrukce'],
)

ber_type = st.sidebar.multiselect(
    "Select the BER - FTE Type:",
    options=groups['vypocty'],
    default=None,
)

h2_type = st.sidebar.multiselect(
    "Select the H2 - FTE Type:",
    options=groups['h2'],
    default=None,
)

ets_type = st.sidebar.multiselect(
    "Select the ETS - FTE Type:",
    options=groups['ets'],
    default=None,
)
fte_type = cad_type + ber_type + h2_type + ets_type

# ---- MAIN Page ----
day1, day2 = st.slider('Time Period:',
                        min_value= min(kapa.index).date(),
                        max_value= max(kapa.index).date(),
                        value=(min(kapa.index).date(), max(kapa.index).date())
                        )

## ---- Dynamické Proměnné ----
df_filtered_cat = df_category(kapa, kapa_par, fte_type)
df_filtered = df_timestamp(df_filtered_cat, day1, day2)

# ---- Chart's AREA ----
overall_data = {
    "Zúčtovatelný": productive_hours(df_filtered, kapa_par),
    "Dočasná nezaměstnanost": pun_hours(df_filtered),
    "Dovolená": urlaub_hours(df_filtered),
    "Nemoc": krank_hours(df_filtered),
    "Školení": training_hours(df_filtered),
    "Jiný": unproductive_hours(df_filtered, kapa_par)
}
overall_data = pd.Series(overall_data)

tab1, tab2, tab3, tab4 = st.tabs(["Overall", "KWs", "Projects", "Checks"])

with tab1:
    st.header("Overall")
    col1, col2 = st.columns(2)
    with col1:
        # ---- Koláčový graf - vše dohromady ----
        pie_overall = px.pie(
            values=overall_data.values,
            names=overall_data.index,
            title='Overall Data',
            width=200)
        st.plotly_chart(pie_overall, use_container_width=False)

    # ---- Koláčový graf - Activity Rate ----
    ar_data = {
        "Neproduktivní": unproductive_hours(df_filtered, kapa_par),
        "Zúčtovatelný": productive_hours(df_filtered, kapa_par),
    }
    ar = productive_hours(df_filtered, kapa_par) / dispo_hours(df_filtered)* 100
    ar_data = pd.Series(ar_data)

    with col2:
        pie_ar = px.pie(
            values=ar_data.values,
            names=ar_data.index,
            title=f'Activity Rate - {ar:.1f}%',
            width=200)
        st.plotly_chart(pie_ar)

with tab2:
    st.header("KWs")
    # ---- Statistiky po týdnech ----
    stat_chart = px.bar(
        df_statistics(df_filtered_cat, kapa_par)[['Produktivní', 'Neproduktivní']]
        )
    st.plotly_chart(stat_chart)

with tab3:
    st.header("Projects")
    project = st.multiselect(
    "Select Project / Value to count:",
    options=get_projects(df_filtered),
    default=None,
    )
    proj_hours = count_values(df_filtered, project)
    st.write(proj_hours, 'Hours in defined time period')

with tab4:
    st.header("Checks")
    # ---- Jen kontroly ----
    st.write(fte_type)
    st.write(df_filtered)


st.sidebar.subheader('Live Stats:')
forecast = f'{productive_hours_forecast(df_filtered_cat ,kapa_par):.1f}'
this_month = f'{productive_hours(df_filtered_cat[df_filtered_cat.index.month == datetime.date.today().month], kapa_par):.1f}'
st.sidebar.write(float(this_month), f'Total Valued Hours in this Month')
st.sidebar.write(float(forecast), f'Valued Hours till End of Month')