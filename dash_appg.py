# =========================================================
# POONAWALLA ENGINEERING GROUP - HRIS DASHBOARD
# FINAL COMPLETE VERSION
# PORT : 8054
# =========================================================

import os
import warnings
warnings.filterwarnings("ignore")

from matplotlib import style
import pandas as pd
import numpy as np
import base64
import requests

import dash
from dash import dcc, html, dash_table
from dash.dependencies import Input, Output

import plotly.express as px
import plotly.graph_objects as go

# =========================================================
# DASH APP
# =========================================================

app = dash.Dash(__name__)
server = app.server

# =========================================================
# COLORS
# =========================================================

BG = "#070B28"
CARD = "#11162A"

TEXT = "#F8FAFC"
SUBTEXT = "#94A3B8"

# =========================================================
# LOGO
# =========================================================

logo_path = "/home/sarthak/Desktop/ANNA work/logo.png"

encoded_logo = ""

if os.path.exists(logo_path):

    with open(logo_path, "rb") as image_file:

        encoded_logo = base64.b64encode(

            image_file.read()

        ).decode()

# =====================================================
# ONEDRIVE EXCEL DOWNLOAD
# =====================================================

DOWNLOAD_URL = "https://intervalvepoonawalla-my.sharepoint.com/:x:/g/personal/navnath_wani_poonawallagroup_com/IQBtS_E2NYhbRqXGf_wt-EqeAWqorVRI7eFtc84tY-saGj4?download=1"

TEMP_FILE = "latest_hr.xlsx"

# =========================================================
# LOAD DATA
# =========================================================

def load_data():

    try:

        # ==========================================
        # DOWNLOAD LATEST EXCEL
        # ==========================================

        try:

            response = requests.get(

                DOWNLOAD_URL,

                timeout=60

            )

        except Exception as e:

            print(

                "OneDrive Download Error:",

               e

           )

            return pd.DataFrame(),{}

        with open(
            TEMP_FILE,
            "wb"
        ) as f:

            f.write(
                response.content
            )

        # ==========================================
        # LOAD EXCEL
        # ==========================================

        df = pd.read_excel(
            TEMP_FILE
        )

        print(
            "Latest OneDrive File Loaded"
        )

    except Exception as e:

        print(e)

        return pd.DataFrame(),{}

    # ==========================================
    # REMOVE EMPTY HEADER ROW
    # ==========================================

    df=df.iloc[1:].reset_index(drop=True)

    # ==========================================
    # RENAME REAL PEG COLUMNS
    # ==========================================

    df.columns=[

        "sno",
        "bu",
        "active_status",
        "emp_no",
        "employee_name",
        "gender",
        "birth_date",
        "age",
        "joining_date",
        "status",
        "exit_date",
        "employee_type",
        "department",
        "grade",
        "designation",
        "location",
        "q1",
        "q1_yop",
        "q2",
        "q2_yop",
        "q3",
        "q3_yop",
        "qualification_group",
        "pre_exp_year",
        "pre_exp_month",
        "ext_exp_year",
        "ext_exp_month",
        "total_exp_year",
        "total_exp_month",
        "final_exp_year",
        "final_exp_month",
        "ctc"

    ]

    # ==========================================
    # DATE CLEANING
    # ==========================================

    for c in [

        "birth_date",
        "joining_date",
        "exit_date"

    ]:

        df[c]=pd.to_datetime(
            df[c],
            errors="coerce"
        )

    # ==========================================
    # GENDER
    # ==========================================

    df["gender"]=(
        df["gender"]
        .astype(str)
        .str.strip()
        .replace({

            "M":"Male",
            "F":"Female"

        })
    )

    # ==========================================
    # ACTIVE STATUS
    # ==========================================

    # ==========================================
    # ACTIVE STATUS CLEANING
    # ==========================================

    df["active_status"]=(

        df["active_status"]

        .astype(str)

        .str.strip()

        .str.upper()

    )

    df["active_status"]=df[
        "active_status"
    ].replace({

        "Y":"Active",
        "N":"Inactive",
        "YES":"Active",
        "NO":"Inactive",
        "ACTIVE":"Active",
        "INACTIVE":"Inactive"

    })

    df["active_status"]=df[
        "active_status"
    ].fillna("Inactive")

    # ==========================================
    # EXPERIENCE
    # ==========================================

    df["total_experience"]= (

        pd.to_numeric(
            df["final_exp_year"],
            errors="coerce"
        ).fillna(0)

        +

        (

            pd.to_numeric(
                df["final_exp_month"],
                errors="coerce"
            ).fillna(0)

            /12

        )

    )

    # ==========================================
    # AGE
    # ==========================================

    current=pd.Timestamp.now()

    df["calculated_age"]=(

        (current-df["birth_date"])
        .dt.days

        /365.25

    )

    df["calculated_age"]=(
        df["calculated_age"]
        .fillna(df["age"])
        .fillna(0)
    )
    # ==========================================
    # AGE GROUP
    # ==========================================

    df["calculated_age"]=pd.to_numeric(

       df["calculated_age"],

       errors="coerce"

    )   

    bins=[20,25,30,35,40,45,50,55,60,65,70,75]

    labels=[

        "21-25",
        "26-30",
        "31-35",
        "36-40",
        "41-45",
        "46-50",
        "51-55",
        "56-60",
        "61-65",
        "66-70",
        "71-75"

    ]

    df["age_group"]=pd.cut(

        df["calculated_age"],

        bins=bins,

        labels=labels,

        include_lowest=True

    )

    df["age_group"]=(
        df["age_group"]
        .astype(str)
    )

    df["age_group"]=df[
        "age_group"
    ].replace(

        "nan",
        "Unknown"

    )

    # ==========================================
    # CTC
    # ==========================================

    df["ctc"]=pd.to_numeric(

        df["ctc"],
        errors="coerce"

    ).fillna(0)

    def convert_ctc(x):

        x=float(x)

        if x>=10000000:

            return round(
                x/10000000,
                2
            )

        return round(
            x/100000,
            2
        )

    def ctc_label(x):

        x=float(x)

        if x>=10000000:

            return f"{round(x/10000000,2)} Cr"

        return f"{round(x/100000,2)} L"

    df["ctc_display"]=df[
        "ctc"
    ].apply(convert_ctc)

    df["ctc_label"]=df[
        "ctc"
    ].apply(ctc_label)

    # ==========================================
    # YEAR FILTER
    # ==========================================

    df["joining_year"]=(
        df["joining_date"]
        .dt.year
    )

    # ==========================================
    # DASHBOARD COLUMN MAP
    # ==========================================

    cols={

        "bu":"bu",

        "dept":"department",

        "grade":"grade",

        "qualification":
        "qualification_group",

        "gender":"gender",

        "status":"status",

        "emp_type":"employee_type",

        "active":"active_status",

        "designation":"designation",

        "joining":"joining_date",

        "birth":"birth_date",

        "ctc":"ctc",

        "experience":
        "total_experience"

    }

    return df,cols

# =========================================================
# KPI CARD
# =========================================================

def kpi_card(title, value, color):

    return html.Div([

        html.P(
            title,
            style={
                "color": SUBTEXT,
                "fontSize": "14px"
            }
        ),

        html.H2(
            str(value),
            style={
                "color": color
            }
        )

    ],
    style={

        "background": CARD,
        "padding": "20px",
        "borderRadius": "18px",
        "width": "24%",
        "textAlign": "center"

    })

# =========================================================
# GRAPH BOX
# =========================================================

def graph_box(graph_id):

    return html.Div([

        dcc.Graph(
            id=graph_id,
            config={
                "displaylogo": False
            }
        )

    ],
    style={

        "background": CARD,
        "padding": "10px",
        "borderRadius": "18px",
        "width": "32%"

    })

# =========================================================
# APP LAYOUT
# =========================================================

app.layout = html.Div([

    dcc.Interval(
        id="interval-component",
        interval=300000,
        n_intervals=0
    ),

    # =====================================================
    # HEADER
    # =====================================================

    html.Div([

        html.Div([

            html.Img(

               src=f"data:image/png;base64,{encoded_logo}",

               style={

                   "height":"80px",

                   "width":"auto",

                   "marginRight":"20px"

               }

           ),

           html.H1(

               "Poonawalla Engineering Group - HRIS Dashboard",

               style={

                   "color":TEXT,

                   "margin":"0"

               }

           )

       ],

        style={

            "display":"flex",

            "alignItems":"center",

            "justifyContent":"center"

        })

    ]),

    html.Br(),

    # =====================================================
    # FILTERS
    # =====================================================

    html.Div([

        dcc.Dropdown(
            id="bu-filter",
            multi=True,
            placeholder="Select BU"
        ),

        dcc.Dropdown(
            id="emp-filter",
            multi=True,
            placeholder="Select Emp Type"
        ),

        dcc.Dropdown(
            id="active-filter",
            multi=True,
            placeholder="Select Active Status"
        ),

        dcc.Dropdown(
            id="dept-filter",
            multi=True,
            placeholder="Select Department"
        ),

        dcc.Dropdown(
            id="grade-filter",
            multi=True,
            placeholder="Select Grade"
        ),

        dcc.Dropdown(
            id="qualification-filter",
            multi=True,
            placeholder="Select Qualification"
        ),

        dcc.Dropdown(
            id="gender-filter",
            multi=True,
            placeholder="Select Gender"
        ),

        dcc.Dropdown(
            id="year-filter",
            multi=True,
            placeholder="Select JoiningYear"
        ),

        dcc.Dropdown(
            id="location-filter",
            multi=True,
            placeholder="Select Location"
        ),

        dcc.Dropdown(
            id="age-filter",
            multi=True,
            placeholder="Select Age Group"
        )

    ],

    style={

        "display":"grid",

        "gridTemplateColumns":
        "repeat(5,1fr)",

        "gap":"15px"

    }),

    html.Br(),

    # =====================================================
    # KPI
    # =====================================================

    html.Div(
        id="kpi-section",
        style={
            "display": "flex",
            "justifyContent": "space-between",
            "gap": "10px"
        }
    ),

    html.Br(),

    # =====================================================
    # SUMMARY
    # =====================================================

    html.Div([

        html.H3(
            "AI Executive Summary",
            style={"color": "#00E5C3"}
        ),

        html.Div(
            id="summary-div",
            style={"color": TEXT}
        )

    ],
    style={

        "background": CARD,
        "padding": "20px",
        "borderRadius": "18px"

    }),

    html.Br(),

    # =====================================================
    # ROW 1
    # =====================================================

    html.Div([

        graph_box("emp-chart"),
        graph_box("qualification-chart"),
        graph_box("active-chart")

    ],
    style={
        "display": "flex",
        "justifyContent": "space-between"
    }),

    html.Br(),

    # =====================================================
    # ROW 2
    # =====================================================

    html.Div([

        graph_box("dept-chart"),
        graph_box("designation-chart"),
        graph_box("age-chart")

    ],
    style={
        "display": "flex",
        "justifyContent": "space-between"
    }),

    html.Br(),

    # =====================================================
    # ROW 3
    # =====================================================

    html.Div([

        graph_box("ctc-chart"),
        graph_box("experience-chart"),
        graph_box("joining-chart")

    ],
    style={
        "display": "flex",
        "justifyContent": "space-between"
    }),

    html.Br(),

    # =====================================================
    # DYNAMIC ANALYZER
    # =====================================================

    html.Div([

        html.H2(
            "Dynamic Relationship Analyzer",
            style={"color": "#7C3AED"}
        ),

        html.Div([

            dcc.Dropdown(
                id="x-axis",
                placeholder="Select X Axis"
            ),

            dcc.Dropdown(
                id="y-axis",
                placeholder="Select Y Axis"
            ),

            dcc.Dropdown(
                id="color-axis",
                placeholder="Select Color Axis"
            )

        ],
        style={

            "display": "grid",
            "gridTemplateColumns": "repeat(3,1fr)",
            "gap": "15px",
            "marginBottom": "20px"

        }),

        dcc.Graph(id="dynamic-chart")

    ],
    style={

        "background": CARD,
        "padding": "20px",
        "borderRadius": "18px"

    }),

    html.Br(),

    # =====================================================
    # TABLE
    # =====================================================

    dash_table.DataTable(

        id="table",

        page_size=10,

        filter_action="native",
        sort_action="native",

        style_header={

            "backgroundColor": "#1E293B",
            "color": "white",
            "fontWeight": "bold"

        },

        style_cell={

            "backgroundColor": CARD,
            "color": TEXT,
            "padding": "10px"

        }

    )

],
style={

    "backgroundColor": BG,
    "padding": "20px",
    "fontFamily": "Arial",
    "minHeight": "100vh"

})

# =========================================================
# CALLBACK
# =========================================================

@app.callback(

    [
        Output("bu-filter","options"),

        Output("emp-filter","options"),

        Output("active-filter","options"),

        Output("dept-filter","options"),

        Output("grade-filter","options"),

        Output("qualification-filter","options"),

        Output("gender-filter","options"),

        Output("year-filter","options"),

        Output("location-filter","options"),

        Output("age-filter","options"),

        Output("kpi-section","children"),

        Output("summary-div","children"),

        Output("emp-chart","figure"),
        Output("qualification-chart","figure"),
        Output("active-chart","figure"),

        Output("dept-chart","figure"),
        Output("designation-chart","figure"),
        Output("age-chart","figure"),

        Output("ctc-chart","figure"),
        Output("experience-chart","figure"),
        Output("joining-chart","figure"),

        Output("dynamic-chart","figure"),

        Output("x-axis","options"),
        Output("y-axis","options"),
        Output("color-axis","options"),

        Output("table","data"),
        Output("table","columns")
    ],

    [
        Input("interval-component","n_intervals"),

        Input("bu-filter","value"),

        Input("emp-filter","value"),

        Input("active-filter","value"),

        Input("dept-filter","value"),

        Input("grade-filter","value"),

        Input("qualification-filter","value"),

        Input("gender-filter","value"),

        Input("year-filter","value"),

        Input("location-filter","value"),

        Input("age-filter","value"),

        Input("x-axis","value"),
        Input("y-axis","value"),
        Input("color-axis","value")
    ]

)

def update_dashboard(

    n,

    bu,
    emp,
    active,
    dept,
    grade,
    qualification,
    gender,
    year,
    location,
    age,

    x_axis,
    y_axis,
    color_axis

):

    try:

        df, cols = load_data()
        full_df = df.copy()

        # =================================================
        # DEFAULT ACTIVE EMPLOYEES
        # =================================================

        dff = df[

            df[cols["active"]]

            .fillna("")

            .astype(str)

            .str.strip()

            .str.lower()

            == "active"

        ].copy()

        # =================================================
        # FILTERS
        # =================================================

        def apply_filter(frame, column, values):

            if column and values:

                return frame[
                    frame[column]
                    .astype(str)
                    .isin(values)
                ]

            return frame

        dff = apply_filter(dff, cols["bu"], bu)
        dff = apply_filter(dff, cols["dept"], dept)
        dff = apply_filter(dff, cols["grade"], grade)
        dff = apply_filter(dff, cols["qualification"], qualification)
        dff = apply_filter(dff, cols["gender"], gender)
        if year:

            dff = dff[
                dff["joining_year"]
                .isin(year)
            ]
        dff=apply_filter(
            dff,
            cols["emp_type"],
            emp
        )

        # =================================================
        # ACTIVE FILTER
        # =================================================

        if active:

            dff=apply_filter(
                full_df,
                cols["active"],
                active
            )

        dff=apply_filter(
            dff,
            "location",
            location
        )

        dff=apply_filter(
            dff,
            "age_group",
            age
        )
        # =================================================
        # FILTER OPTIONS
        # =================================================

        # =================================================
        # SAFE DROPDOWN OPTIONS
        # =================================================

        def options(column):

            if not column:
                return []

            try:

                vals = (

                    full_df[column]

                    .fillna("Unknown")

                    .astype(str)

                    .replace("nan","Unknown")

                    .unique()

                )

                vals = sorted(

                    vals,

                    key=lambda x:str(x)

                )

                return [

                    {

                        "label":str(i),

                        "value":str(i)

                    }

                    for i in vals

                ]

            except:

                return []

        # =================================================
        # DROPDOWN OPTIONS
        # =================================================

        bu_options=options(
            cols["bu"]
        )

        emp_options=options(
            cols["emp_type"]
        )

        active_options=options(
            cols["active"]
        )

        dept_options=options(
            cols["dept"]
        )

        grade_options=options(
            cols["grade"]
        )

        qualification_options=options(
            cols["qualification"]
        )

        gender_options=options(
            cols["gender"]
        )

        location_options=options(
            "location"
        )

        age_options=options(
            "age_group"
        )

        year_options=[

            {

                "label":str(int(i)),
                "value":int(i)

            }

            for i in sorted(

                full_df["joining_year"]

                .dropna()

                .astype(int)

                .unique()

            )

        ]

        # =================================================
        # KPI
        # =================================================

        active_df = dff[

           dff[cols["active"]]

            .fillna("")

            .astype(str)

            .str.strip()

            .str.lower()

            == "active"

        ]

        # ================================================
        # ACTIVE EMPLOYEE
        # ================================================

        active_emp = len(
            active_df
        )

        # ================================================
        # MALE ACTIVE
        # ================================================

        male_active = active_df[

            active_df[cols["gender"]]

            .astype(str)

            .str.lower()

            == "male"

        ].shape[0]

        # ================================================
        # FEMALE ACTIVE
        # ================================================

        female_active = active_df[

            active_df[cols["gender"]]

            .astype(str)

            .str.lower()

            == "female"

        ].shape[0]

        # ================================================
        # TOTAL CTC
        # ================================================

        total_ctc = dff[
            cols["ctc"]
        ].sum()

        if total_ctc >= 10000000:

            total_ctc_label = (

                f"{round(total_ctc/10000000,2)} Cr"

            )

        else:

            total_ctc_label=(

                f"{round(total_ctc/100000,2)} L"

            )

        # ================================================
        # KPI CARDS
        # ================================================

        kpis=[

            kpi_card(

                "Active Employees",

                active_emp,

                "#7C3AED"

            ),

            kpi_card(
                    
                "Male Active",

                male_active,

                "#3B82F6"

            ),

            kpi_card(

                "Female Active",

                female_active,

                "#EC4899"

            ),

            kpi_card(

                "Total CTC",

                total_ctc_label,

                "#F59E0B"

            )

        ]

        # =================================================
        # SUMMARY
        # =================================================

        summary=f"""
    
        Active Employees : {active_emp}

        Male Active : {male_active}

        Female Active : {female_active}

        Total Employee CTC : {total_ctc_label}

        PEG HR workforce analytics updated successfully.

        """

        # =================================================
        # FIGURE STYLE
        # =================================================

        def style(fig, title):

            fig.update_layout(

                title=title,

                paper_bgcolor=CARD,
                plot_bgcolor=CARD,

                font=dict(color=TEXT)

            )

            return fig

        # =================================================
        # CHARTS
        # =================================================

        # =====================================================
        # EMPLOYEE TYPE
        # =====================================================

        emp_fig = px.histogram(

            dff,

            x=cols["emp_type"],

            text_auto=True,

            color=cols["emp_type"]

        )

        emp_fig.update_traces(

            textposition="outside"

        )

        emp_fig = style(

            emp_fig,

            "Employee Type"

        )

        qualification_fig = px.pie(
            dff,
            names=cols["qualification"],
            hole=0.4
        )

        qualification_fig.update_traces(
            textinfo="label+value"
        )

        qualification_fig = style(
            qualification_fig,
            "Qualification Analysis"
        )

        # =====================================================
        # EMPLOYEE STATUS DISTRIBUTION
        # =====================================================

        status_data=(

            full_df[
                cols["status"]
            ]

            .fillna("Unknown")

            .astype(str)

            .value_counts()

            .reset_index()

        )

        status_data.columns=[

            "Status",
            "Count"

        ]

        active_fig=px.pie(

            status_data,

            names="Status",

            values="Count",

            hole=0.4,

            color_discrete_sequence=px.colors.qualitative.Bold

        )

        active_fig.update_traces(

            textinfo="label+value"

        )

        active_fig=style(

            active_fig,

            "Employee Status Analysis"

        )

        # =====================================================
        # TOP DEPARTMENT DISTRIBUTION
        #        =====================================================

        dept_data=(

            dff[
                    cols["dept"]
            ]

            .value_counts()

            .head(15)

            .reset_index()

        )

        dept_data.columns=[

            "Department",
            "Count"

        ]
        
        dept_fig=px.bar(

            dept_data,

            y="Department",

            x="Count",

            orientation="h",

            text="Count",

            color="Count",

            color_continuous_scale="viridis"

        )

        dept_fig.update_traces(
            textposition="outside"
        )

        dept_fig.update_layout(

            height=700,

            yaxis=dict(
                categoryorder="total ascending"
            )

        )

        dept_fig=style(
            dept_fig,
            "Department Analysis"
        )

        

        # =================================================
        # DEPARTMENT WISE GRADE COUNT
        # =================================================

        # =====================================================
        # DEPARTMENT VS GRADE HEATMAP
        # =====================================================

        heat_df=(

           dff

            .groupby(

                [

                    cols["dept"],
                    cols["grade"]

                ]

            )

           .size()

           .reset_index(

                name="count"

            )

        )

        designation_fig=px.density_heatmap(

            heat_df,

            y=cols["dept"],

            x=cols["grade"],

            z="count",

            text_auto=True,

            color_continuous_scale="Turbo"

        )

        designation_fig.update_layout(

            height=900

        )

        designation_fig=style(

            designation_fig,

            "Department vs Grade Employee Count"

        )

        # =====================================================
        # AGE DISTRIBUTION
        # =====================================================

        bins=[20,25,30,35,40,45,50,55,60,65,70,75]

        labels=[

            "20-25",
            "26-30",
            "31-35",
            "36-40",
            "41-45",
            "46-50",
            "51-55",
            "56-60",
            "61-65",
            "66-70",
            "71-75"

        ]

        dff["Age_Group"]=pd.cut(

            dff["calculated_age"],

            bins=bins,

            labels=labels,

            include_lowest=True

        )

        age_df=(

            dff.groupby(

                ["Age_Group",cols["gender"]]

            )

            .size()

            .reset_index(

                name="Count"

            )

        )

        age_fig=px.bar(

            age_df,

            x="Age_Group",

            y="Count",

            color=cols["gender"],

            barmode="group",

            text="Count",

            color_discrete_sequence=[
                "#7C3AED",
                "#F97316"
            ]

        )

        age_fig.update_traces(

            textposition="outside"

        )

        age_fig.update_layout(

            xaxis_title="Age Group",

            yaxis_title="Employees",

            height=650

        )

        age_fig=style(

            age_fig,

            "Age Analysis (Male vs Female)"

        )

        # =====================================================
        # GRADE-WISE CTC TABLE HEATMAP
        # =====================================================

        ctc_data=(

            dff.groupby(

                cols["grade"]

            )[

                "ctc_display"

            ]

            .agg(

                ["min","mean","max","count"]

            )

            .reset_index()

        )

        ctc_data.columns=[

            "Grade",
            "Min",
            "Avg",
            "Max",
            "Employees"

        ]

        ctc_data["Min"]=ctc_data["Min"].round(2)

        ctc_data["Avg"]=ctc_data["Avg"].round(2)

        ctc_data["Max"]=ctc_data["Max"].round(2)

        ctc_fig=px.imshow(

            ctc_data[

                [

                    "Min",
                    "Avg",
                    "Max"

                ]

            ],

            text_auto=True,

            y=ctc_data["Grade"],

            x=[

                "Min",
                "Average",
                "Max"

            ],

            aspect="auto",

            color_continuous_scale="Turbo"

        )

        ctc_fig.update_layout(

            height=max(
                800,
               len(ctc_data)*35
           )

        )

        ctc_fig=style(

           ctc_fig,

            "Grade-wise CTC (Min / Avg / Max)"

        )

        

        # =====================================================
        # EXPERIENCE HEATMAP
        # =====================================================

        exp_stats=(

           dff.groupby(

                cols["grade"]

            )[

                cols["experience"]

            ]

            .agg(

                ["min","mean","max"]

            )

            .reset_index()

        )

        exp_stats.columns=[

            "Grade",
            "Min",
            "Average",
            "Max"

        ]

        # ================================================
        # ROUND VALUES
        # ================================================

        exp_stats["Min"]=(
            exp_stats["Min"]
            .round(1)
        )

        exp_stats["Average"]=(
            exp_stats["Average"]
            .round(1)
        )

        exp_stats["Max"]=(
            exp_stats["Max"]
            .round(1)
        )

        # ================================================
        # YEAR MONTH LABEL
        # ================================================

        def ym(x):

            years=int(x)

            months=int(

                round(
                    (x-years)*12
                )
            )

            return f"{years}Y {months}M"

        text_data=[

            [

                ym(v)

                for v in row

            ]

            for row in exp_stats[
                [

                    "Min",
                    "Average",
                    "Max"

                ]

            ].values

        ]

        # ================================================
        # HEATMAP
        # ================================================

        exp_fig=go.Figure(

            data=go.Heatmap(

                z=exp_stats[
                    [

                        "Min",
                        "Average",
                        "Max"

                    ]

                ].values,

                x=[

                    "Min",
                    "Average",
                    "Max"

                ],

                y=exp_stats["Grade"],

                text=text_data,

                texttemplate="%{text}",

                colorscale="Turbo"

            )

        )

        exp_fig.update_layout(

            height=max(
                900,
                len(exp_stats)*35
            ),

            xaxis_title="Experience Type",

            yaxis_title="Grade"

        )

        exp_fig=style(

            exp_fig,

            "Grade-wise Experience (Min / Avg / Max)"

        )

        # =====================================================
        # JOINING TREND
        # =====================================================

        join_year=(

            dff["joining_year"]

            .dropna()

            .astype(int)

            .value_counts()

           .sort_index()

            .reset_index()

        )

        join_year.columns=[

            "Year",
            "Count"

        ]

        join_year=join_year.sort_values(
            "Year"
        )

        join_fig=px.bar(

            join_year,

            y="Year",

            x="Count",

            orientation="h",

            text="Count",

            color="Count",

            color_continuous_scale="Turbo"

        )

        join_fig.update_traces(

            textposition="outside"

        )

        join_fig.update_layout(

            height=max(
                700,
                len(join_year)*35
        ),

            yaxis_title="Joining Year",

            xaxis_title="Employee Count"

        )

        join_fig=style(

            join_fig,

            "Joining Trend"

        )

        # =================================================
        # DYNAMIC DROPDOWNS
        # =================================================

        numeric_cols = []

        for c in dff.columns:

            try:

                if pd.api.types.is_numeric_dtype(
                    dff[c]
                ):
                    numeric_cols.append(c)

            except:
                pass

        x_options = [
            {"label": i, "value": i}
            for i in numeric_cols
        ]

        y_options = x_options

        color_options = [
            {"label": i, "value": i}
            for i in dff.columns
        ]

        if not x_axis:
            x_axis = "calculated_age"

        if not y_axis:
            y_axis = "ctc_display"

        if not color_axis:
            color_axis = cols["dept"]

        # =================================================
        # DYNAMIC GRAPH
        # =================================================

        dynamic_fig = px.scatter(

            dff,

            x=x_axis,

            y=y_axis,

            color=color_axis,

            hover_data=dff.columns

        )

        dynamic_fig.update_traces(
            marker=dict(size=12)
        )

        dynamic_fig = style(
            dynamic_fig,
            "Dynamic Relationship Analyzer"
        )

        # =================================================
        # TABLE
        # =================================================

        table_data = dff.head(200).to_dict("records")

        table_columns = [

            {
                "name": i,
                "id": i
            }

            for i in dff.columns

        ]

        return (

            bu_options,

            emp_options,

            active_options,

            dept_options,

            grade_options,

            qualification_options,

            gender_options,

            year_options,

            location_options,

            age_options,

            kpis,
            summary,

            emp_fig,
            qualification_fig,
            active_fig,

            dept_fig,
            designation_fig,
            age_fig,

            ctc_fig,
            exp_fig,
            join_fig,

            dynamic_fig,

            x_options,
            y_options,
            color_options,

            table_data,
            table_columns

        )

    except Exception as e:

        print("CALLBACK ERROR :", e)

        empty = go.Figure()

        return (

            [], [], [], [], [], [],

            [],

            str(e),

            empty,
            empty,
            empty,

            empty,
            empty,
            empty,

            empty,
            empty,
            empty,

            empty,

            [],
            [],
            [],

            [],
            []

        )

# =========================================================
# RUN APP
# =========================================================

if __name__ == "__main__":

    print("=" * 60)
    print("Poonawalla Engineering Group - HRIS Dashboard")
    print("URL : http://127.0.0.1:8054")
    print("=" * 60)

    app.run(
        debug=False,
        port=8054
    )