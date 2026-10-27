# Databricks notebook source
# MAGIC %md
# MAGIC # Leiden QA - OCP Attendance & Task Assignment Automation
# MAGIC
# MAGIC - This notebook is designed to ingest, transform, and prepare data required for automating the QA OCP daily task assignment process.
# MAGIC
# MAGIC #### Objective
# MAGIC
# MAGIC The goal is to replace manual Excel-based scheduling by building a data-driven solution in Databricks that:
# MAGIC
# MAGIC - Identifies employee availability on each working day
# MAGIC - Applies role-based eligibility and constraints
# MAGIC - Supports automated assignment of operational tasks (e.g., QOTF, QA Exp Contact, Triager, etc.)
# MAGIC
# MAGIC This notebook serves as the foundation for building a scalable and maintainable QA task automation pipeline.
# MAGIC
# MAGIC Inshort these note book - 
# MAGIC -  Replace manual Excel-based scheduling with a scalable assignment engine that:
# MAGIC -  Identifies employee availability by day
# MAGIC - Applies role/subteam/eligibility rules
# MAGIC - Prevents over-assignment using task count logic
# MAGIC -  Supports daily / weekly / specific-day tasks
# MAGIC -  Produces a dashboard-ready output table

# COMMAND ----------

# MAGIC %md
# MAGIC #### Input Data Sources
# MAGIC
# MAGIC 1. Attendance_Schedule
# MAGIC   - Contains daily employee attendance/availability in a matrix format
# MAGIC   - Each row represents a date
# MAGIC   - Each column represents an employee
# MAGIC   - Cell values represent attendance codes (W, L, WFH, SL, etc.)
# MAGIC 2. Roles_QA_Master
# MAGIC   - Contains employee details such as:
# MAGIC     - Team
# MAGIC     - Role
# MAGIC     - FTE
# MAGIC     - Active status
# MAGIC     - Task eligibility flags (QOTF, QA Exp Contact, etc.)
# MAGIC 3. Codes
# MAGIC   - Defines the meaning of attendance codes
# MAGIC   - Used to determine employee availability (Available / Not Available)
# MAGIC 4. Task schedule 
# MAGIC   - used for the validation with th output 

# COMMAND ----------

# MAGIC %pip install Office365-REST-Python-Client==2.4.1 sqlalchemy tqdm
# MAGIC %pip install xlrd
# MAGIC dbutils.library.restartPython()

# COMMAND ----------

# MAGIC %md
# MAGIC # Load input files

# COMMAND ----------

# MAGIC %md
# MAGIC ## Share point connection - CDP certificate authentication --New Method

# COMMAND ----------


import pyspark.sql.functions as F
from pyspark.sql import Window
import pandas as pd
import io

from office365.sharepoint.client_context import ClientContext
from office365.sharepoint.files.file import File

TENANT_ID = dbutils.secrets.get(scope="cdp-sharepoint-scope",key="tenantid")
CLIENT_ID = dbutils.secrets.get(scope="cdp-sharepoint-scope",key="clientid")
THUMBPRINT = dbutils.secrets.get(scope="cdp-sharepoint-scope",key="thumbprint")
PRIVATE_KEY_PEM = dbutils.secrets.get(scope="cdp-sharepoint-scope",key="privatekey")

cert_settings = {"client_id": CLIENT_ID,"thumbprint": THUMBPRINT,"private_key": PRIVATE_KEY_PEM}

site_url = "https://jnj.sharepoint.com/teams/QualityEveryonesBusiness"

file_url = (
    "teams/QualityEveryonesBusiness/"
    "Shared Documents/General/2. Leave schedules/"
    "Leave Schedule QA OCP-NPD.xlsx"
)

ctx = ClientContext(site_url).with_client_certificate(
    TENANT_ID,
    **cert_settings
)

web = ctx.web
ctx.load(web)
ctx.execute_query()

print(f" Connected to: {web.url}")

file_response = File.open_binary(
    ctx,
    "/" + file_url.lstrip("/")
)

file_bytes = io.BytesIO(file_response.content)

# COMMAND ----------

#--------------------------------------------------------------
# Read Sheets - Leave Schedule /Roles_QA/ Task Schedule / Codes
# --------------------------------------------------------------

# Leave Schedule
# leave_pd = pd.read_excel(
#     file_bytes,
#     sheet_name="Leave Schedule",
#     engine="openpyxl",
#     skiprows=1,
#     dtype=str
# )
leave_pd = pd.read_excel(
    file_bytes,
    sheet_name="Leave Schedule",
    engine="openpyxl",
    skiprows=1,
    dtype=str,
    keep_default_na=False
)
# Reset pointer before next read
file_bytes.seek(0)

# Task Schedule
# task_pd = pd.read_excel(
#     file_bytes,
#     sheet_name="Task Schedule",
#     engine="openpyxl",
#     skiprows=0,
#     dtype=str
# )
task_pd = pd.read_excel(
    file_bytes,
    sheet_name="Task Schedule",
    engine="openpyxl",
    skiprows=0,
    dtype=str,
    keep_default_na=False
)
# Reset pointer
file_bytes.seek(0)

# Roles sheet
# roles_pd = pd.read_excel(
#     file_bytes,
#     sheet_name="Roles_QA",
#     engine="openpyxl",
#     dtype=str
# )
roles_pd = pd.read_excel(
    file_bytes,
    sheet_name="Roles_QA",
    engine="openpyxl",
    dtype=str,
    keep_default_na=False
)
# Reset pointer
file_bytes.seek(0)

#Codes sheet
# codes_pd = pd.read_excel(
#     file_bytes,
#     sheet_name="Codes",
#     engine="openpyxl",
#     dtype=str
# )
codes_pd = pd.read_excel(
    file_bytes,
    sheet_name="Codes",
    engine="openpyxl",
    dtype=str,
    keep_default_na=False
)

# Clean leave sheet

leave_pd.columns = (
    leave_pd.columns
    .astype(str)
    .str.strip()
)

leave_pd = leave_pd.replace(
    {
        "": pd.NA,
        " ": pd.NA,
        "None": pd.NA,
        "nan": pd.NA
    }
)

leave_pd = leave_pd.dropna(axis=1, how="all")
leave_pd = leave_pd.dropna(how="all")

# --------------------------------------------------
# Fix Date + remove invalid rows
# --------------------------------------------------

leave_pd = leave_pd[leave_pd["Date"].notna()]

leave_pd["Date"] = pd.to_datetime(
    leave_pd["Date"], errors="coerce"
)

leave_pd = leave_pd[leave_pd["Date"].notna()]

# # --------------------------------------------------
# # Remove empty rows (weekends)
# # --------------------------------------------------
employee_cols = [c for c in leave_pd.columns if c != "Date"]

# Detect site closed days before converting blanks to W
def is_site_closed_row(row):
    values = row.dropna().astype(str).str.strip().str.upper()
    if len(values) == 0:
        return False
    
    c_count = (values == "C").sum()
    return c_count >= max(1, len(values) * 0.5)

leave_pd["Is_Site_Closed"] = leave_pd[employee_cols].apply(
    is_site_closed_row,
    axis=1
)

employee_cols = [
    c for c in leave_pd.columns
    if c not in ["Date", "Is_Site_Closed"]
]

# --------------------------------------------------
# Normalize codes
# --------------------------------------------------

def normalize_code(x):

    if pd.isna(x):
        return "W"

    x = str(x).strip().upper()

    if x in ["", "NONE", "NAN"]:
        return "W"

    mapping = {
        "W": "W",
        "WFH": "WFH",
        "WEW": "WEW",
        "L": "L",
        "P": "P",
        "P2": "P",
        "T": "T",
        "R": "R",
        "C": "C",
        "A": "A",
        "SL": "SL",

        # Not Available
        "NAV": "NAV",

        # backward compatibility
        "NA": "NAV",
        "N/A": "NAV",
        "NOT AVAILABLE": "NAV"
    }

    return mapping.get(x, "W") # unknown → Working

for col in employee_cols:
    leave_pd[col] = leave_pd[col].apply(normalize_code)

#Validation
nav_total = 0

for col in employee_cols:

    nav_count = (
        leave_pd[col]
        .astype(str)
        .str.upper()
        .eq("NAV")
        .sum()
    )

    if nav_count > 0:

        print(
            f"{col}: {nav_count}"
        )

        nav_total += nav_count

print(
    "Total NAV entries:",
    nav_total
)
#-------------------------------------
# Final clean
leave_pd = leave_pd.sort_values("Date")

# display(leave_pd.head(10))

# Clean Task
task_pd = task_pd.dropna(axis=1, how="all")
task_pd.columns = task_pd.columns.astype(str).str.strip()

# --------------------------------------------------
# Prepare existing Tier 1 history from Task Schedule
# --------------------------------------------------

tier_history_pd = task_pd[
    [
        "Date",
        "Tier 1 Lead",
        "Tier 1 Back-up"
    ]
].copy()

tier_history_pd["Date"] = pd.to_datetime(
    tier_history_pd["Date"],
    errors="coerce"
)

# Remove Allowed Codes row and other non-date rows
tier_history_pd = tier_history_pd[
    tier_history_pd["Date"].notna()
].copy()

# Clean employee values
for col in ["Tier 1 Lead", "Tier 1 Back-up"]:

    tier_history_pd[col] = (
        tier_history_pd[col]
        .astype("string")
        .str.strip()
    )

    tier_history_pd[col] = tier_history_pd[col].replace(
        {
            "": pd.NA,
            "nan": pd.NA,
            "None": pd.NA,
            "<NA>": pd.NA
        }
    )

# Monday-based week
tier_history_pd["WeekStart"] = (
    tier_history_pd["Date"]
    - pd.to_timedelta(
        tier_history_pd["Date"].dt.weekday,
        unit="D"
    )
)

# Exact date/task lookup
tier_history_lookup = {}

for _, history_row in tier_history_pd.iterrows():

    history_date = pd.Timestamp(
        history_row["Date"]
    ).normalize()

    for history_task in [
        "Tier 1 Lead",
        "Tier 1 Back-up"
    ]:

        history_employee = history_row[history_task]

        if pd.notna(history_employee):

            tier_history_lookup[
                (history_date, history_task)
            ] = str(history_employee).strip()

print(
    "Historical Tier 1 assignments loaded:",
    len(tier_history_lookup)
)
#---------------------------------------------------------------
# Clean Roles
roles_pd.columns = roles_pd.columns.astype(str).str.strip()

# Clean Codes
codes_pd.columns = codes_pd.columns.astype(str).str.strip()

#  Convert to Spark
leave_df = spark.createDataFrame(leave_pd)
task_config_df = spark.createDataFrame(task_pd)
roles_df = spark.createDataFrame(roles_pd)
codes_df = spark.createDataFrame(codes_pd)

# Convert the date to more simplefiled terms
leave_df = leave_df.withColumn(
    "Date",
    F.to_date("Date")
)

# Validate
print("Leave rows:", leave_df.count()) 
print("Roles rows:", roles_df.count())
print("Codes rows:", codes_df.count())

print("Task rows:", task_config_df.count()) #This sheet only use for the validation purpose not for the actual logic creation

display(leave_df)
display(roles_df)
display(codes_df)
#This sheet only use for the validation purpose not for the actual logic creation
display(task_config_df)



# COMMAND ----------

print(codes_pd)

# COMMAND ----------

# MAGIC %md
# MAGIC ###  Unpivot leave_df - making the Leave sheet as long format for the future logic easiness.
# MAGIC

# COMMAND ----------

from pyspark.sql.functions import expr

# Unpivot leave_df

# Identify columns
# fixed_cols = ["Date"]
fixed_cols = ["Date", "Is_Site_Closed"]
employee_cols = [c for c in leave_df.columns if c not in fixed_cols]

# Create stack expression
stack_expr = "stack({0}, {1}) as (Employee, Code)".format(
    len(employee_cols),
    ", ".join([f"'{col}', `{col}`" for col in employee_cols])
)

# Convert wide → long
# attendance_long_df = leave_df.select("Date", expr(stack_expr))
attendance_long_df = leave_df.select(
    "Date",
    "Is_Site_Closed",
    expr(stack_expr)
)

# Remove NULL codes (unknown availability)
attendance_long_df = attendance_long_df.filter(F.col("Code").isNotNull())

display(attendance_long_df)

# Filter only weekday data

attendance_long_df = attendance_long_df.withColumn(
    "DayOfWeek",
    F.date_format("Date", "EEEE")
)

# attendance_long_df = attendance_long_df.filter(
#     F.col("DayOfWeek").isin("Monday", "Tuesday", "Wednesday", "Thursday", "Friday")
# )

# COMMAND ----------

# Clean names
roles_df = roles_df.withColumn("Employee", F.trim(F.col("Employee")))


# COMMAND ----------

# MAGIC %md
# MAGIC ### Assign one person per day (rotation)

# COMMAND ----------

from collections import defaultdict, deque
import pandas as pd
import pyspark.sql.functions as F

# ---------------------------------------------------------
# Build candidate base from attendance_long_df + roles_df
# ---------------------------------------------------------

# Make sure Date is date type
attendance_long_df = attendance_long_df.withColumn(
    "Date",
    F.to_date("Date")
)

# Add DayOfWeek and WeekStart
attendance_long_df = attendance_long_df.withColumn(
    "DayOfWeek",
    F.date_format("Date", "EEEE")
)

# attendance_long_df = attendance_long_df.withColumn(
#     "WeekStart",
#     F.date_sub(F.col("Date"), F.dayofweek(F.col("Date")) - 2)
# )

attendance_long_df = attendance_long_df.withColumn(
    "WeekStart",
    F.date_sub(
        F.col("Date"),
        F.weekday(F.col("Date"))
    )
)

# Clean employee names
attendance_long_df = attendance_long_df.withColumn(
    "Employee",
    F.trim(F.col("Employee"))
)

roles_df = roles_df.withColumn(
    "Employee",
    F.trim(F.col("Employee"))
)

# Join attendance with roles
candidate_base_df = attendance_long_df.join(
    roles_df,
    on="Employee",
    how="left"
)

# Keep only employees available in Roles_QA and Present = Yes
candidate_base_df = candidate_base_df.filter(
    (F.col("Role").isNotNull()) &
    (F.upper(F.col("Present")) == "YES")
)

display(candidate_base_df)

# COMMAND ----------

# MAGIC %md
# MAGIC ### Task  configuration

# COMMAND ----------

# --------------------------------------------------
# Task rule configuration
# --------------------------------------------------

task_rules = [
    # DAILY ROTATION TASKS
    {
        "Task": "QA Exp Contact",
        "Assignment_Type": "DAILY",
        # "Allowed_Codes": ["W", "WFH"],
        "Allowed_Codes": ["W"],   # Onsite only. Blank cells are already converted to W.
        "Eligibility_Flag": "QA_Exp_Contact_Eligible",
        "Active": True
    },
    {
        "Task": "Triager of the day",
        "Assignment_Type": "DAILY",
        "Allowed_Codes": ["W", "WFH"],
        "Eligibility_Flag": "Triager_Eligible",
        "Active": True
    },
    {
        "Task": "SAP notifications",
        "Assignment_Type": "DAILY",
        "Allowed_Codes": ["W", "WFH"],
        "Eligibility_Flag": "SAP_Notification_Eligible",
        "Active": True
    },
    {
        "Task": "QA-ITP",
        "Assignment_Type": "DAILY",
        "Allowed_Codes": ["W"],
        "Eligibility_Flag": "QA_ITP_Eligible",
        "Active": True
    },
    {
        "Task": "Release Inbox",
        "Assignment_Type": "DAILY",
        "Allowed_Codes": ["W", "WFH"],
        "Eligibility_Flag": "Release_Inbox_Eligible",
        "Active": True
    },
    {
        "Task": "Mailbox OCP Review & release",
        "Assignment_Type": "DAILY",
        "Allowed_Codes": ["W", "WFH"],
        "Eligibility_Flag": "QA_OCP_Mailbox_Rev_Release_Eligible",
        "Active": True
    },
    {
        "Task": "Daily Exception Review DSP 10:30",
        "Assignment_Type": "DAILY",
        "Allowed_Codes": ["W", "WFH"],
        "Eligibility_Flag": "Daily_Exception_Review_DSP_1030_Eligible",
        "Active": True
    },
    {
        "Task": "Daily Exception Review PR 13:00",
        "Assignment_Type": "DAILY",
        "Allowed_Codes": ["W", "WFH"],
        "Eligibility_Flag": "Daily_Exception_Review_PR_1300_Eligible",
        "Active": True
    },
    {
        "Task": "Daily Exception Review FERM 13:00",
        "Assignment_Type": "DAILY",
        "Allowed_Codes": ["W", "WFH"],
        "Eligibility_Flag": "Daily_Exception_Review_FERM_1300_Eligible",
        "Active": True
    },

    # WEEKLY FIXED MONDAY-FRIDAY TASKS
    {
        "Task": "Tier 1 Lead",
        "Assignment_Type": "WEEKLY_FIXED",
        "Allowed_Codes": ["W", "WFH"],
        "Eligibility_Flag": "Tier_1_Lead_Eligible",
        "Active": True
    },
    {
        "Task": "Tier 1 Back-up",
        "Assignment_Type": "WEEKLY_FIXED",
        "Allowed_Codes": ["W", "WFH"],
        "Eligibility_Flag": "Tier_1_Back_Up_Eligible",
        "Active": True
    },

    # # ACHTERWACHT MONDAY-SUNDAY
    # {
    #     "Task": "Achterwacht (W)",
    #     "Assignment_Type": "WEEKLY_FIXED_WITH_WEEKEND",
    #     "Allowed_Codes": ["W", "WFH"],
    #     "Eligibility_Flag": "Achterwacht_Eligible",
    #     "Active": True
    # },

    {
    "Task": "QA-in-QC",
    "Assignment_Type": "SPECIFIC_DAY",
    "Specific_Day": "Tuesday",
    "Allowed_Codes": ["W"],
    "Eligibility_Flag": "QA_in_QC_Eligible",
    "Active": True
    },

    # DEPARTMENT ONLY
    # {
    #     "Task": "QA-in-WH",
    #     "Assignment_Type": "DEPARTMENT_ONLY",
    #     "Allowed_Codes": ["ALL"],
    #     "Eligibility_Flag": None,
    #     "Active": True
    # },

    {
    "Task": "QA-in-WH",
    "Assignment_Type": "SPECIFIC_DAY",
    "Specific_Day": "Wednesday",
    "Allowed_Codes": ["W"],
    "Eligibility_Flag": "QA_in_WH_Eligible",
    "Active": True
    },

    # SPECIFIC DAY
    {
        "Task": "Weekly Exception Review MP Wed - 11:00",
        "Assignment_Type": "SPECIFIC_DAY",
        "Specific_Day": "Wednesday",
        "Allowed_Codes": ["W", "WFH"],
        "Eligibility_Flag": "Weekly_Exception_Review_MP_Wed_1100_Eligible",
        "Active": True
    },

    #Weekly Exception Review PREC 14:00
    {
    "Task": "Weekly Exception Review PREC 14:00",
    "Assignment_Type": "SPECIFIC_DAY",
    "Specific_Day": "Monday",
    "Allowed_Codes": ["W", "WFH"],
    "Eligibility_Flag":
        "Weekly_Exception_Review_PREC_1400_Eligible",
    "Active": True
}
]

task_rules_pd = pd.DataFrame(task_rules)
task_rules_pd = task_rules_pd[task_rules_pd["Active"] == True].copy()

display(task_rules_pd)

# COMMAND ----------

# --------------------------------------------------
# QOTF department logic
# --------------------------------------------------

qotf_department_rules = {
    "Monday": ["DSP"],
    "Tuesday": ["EP", "DC"],
    "Wednesday": ["FERM"],
    "Thursday": ["PREC", "BP", "MP"],
    "Friday": ["PR"]
}

def get_qotf_department(day_name, week_number):
    departments = qotf_department_rules.get(day_name, [])

    if len(departments) == 0:
        return None

    selected_index = (week_number - 1) % len(departments)
    return departments[selected_index]

# COMMAND ----------

# MAGIC %md
# MAGIC Prepare Candidate Dataset 

# COMMAND ----------

# --------------------------------------------------
# Convert candidate base to pandas for sequential assignment
# --------------------------------------------------

candidate_base_pd = candidate_base_df.toPandas()

candidate_base_pd["Date"] = pd.to_datetime(candidate_base_pd["Date"])
candidate_base_pd["Employee"] = candidate_base_pd["Employee"].astype(str).str.strip()
candidate_base_pd["Code"] = candidate_base_pd["Code"].astype(str).str.strip().str.upper()
candidate_base_pd["DayOfWeek"] = candidate_base_pd["DayOfWeek"].astype(str).str.strip()

# WeekStart as pandas date
candidate_base_pd["WeekStart"] = pd.to_datetime(candidate_base_pd["WeekStart"])

display(candidate_base_pd.head())

# COMMAND ----------

# --------------------------------------------------
# Candidate filtering function
# --------------------------------------------------

def get_task_candidates(
    base_pd,
    task_rule,
    date_value=None,
    week_start=None
):
    task_name = task_rule["Task"]
    allowed_codes = task_rule["Allowed_Codes"]
    eligibility_flag = task_rule["Eligibility_Flag"]

    temp = base_pd.copy()

    # --------------------------------------------------
    # Filter by a specific date for daily tasks
    # --------------------------------------------------
    if date_value is not None:
        temp = temp[
            temp["Date"] == date_value
        ]

    # --------------------------------------------------
    # Filter by week for weekly tasks
    # --------------------------------------------------
    if week_start is not None:
        temp = temp[
            temp["WeekStart"] == week_start
        ]

    # --------------------------------------------------
    # NA hard exclusion
    # NA means no task can be assigned on that date
    # --------------------------------------------------
    temp = temp[
        temp["Code"]
        .astype(str)
        .str.strip()
        .str.upper()
        != "NA"
    ]

    # --------------------------------------------------
    # Allowed code filter
    # Example:
    # W/WFH tasks → only W or WFH
    # W-only tasks → only W
    # --------------------------------------------------
    if allowed_codes != ["ALL"]:
        temp = temp[
            temp["Code"].isin(allowed_codes)
        ]

 
    # --------------------------------------------------
    # Eligibility flag filter
    # --------------------------------------------------
    if eligibility_flag is not None:

        if eligibility_flag in temp.columns:

            temp = temp[
                temp[eligibility_flag]
                .astype(str)
                .str.strip()
                .str.upper()
                == "YES"
            ]

        else:
            # If the eligibility column does not exist,
            # return an empty candidate dataframe
            temp = temp.iloc[0:0]

    # --------------------------------------------------
    # Not Available (NAV)
    #
    # An employee marked as NAV on a date must not
    # receive any task on that date.
    # --------------------------------------------------

    temp = temp[
        temp["Code"]
        .astype(str)
        .str.strip()
        .str.upper()
        != "NAV"
    ]

    return temp.copy()

# COMMAND ----------

# --------------------------------------------------
# Multi-task assignment engine
# --------------------------------------------------

results = []

# Track assignment count per task and employee
task_assignment_count = defaultdict(lambda: defaultdict(int))

# Track recent assignees per task
recent_assignees_by_task = defaultdict(lambda: deque(maxlen=15))

# Track previous employee per daily task
previous_employee_by_task = {}

#---------------------------------------------------------------------
# Track the last employee assigned to each task
previous_employee_by_task_weekday = {}

#---------------------------------------------------------------------

# Avoid same employee getting multiple tasks on same date if possible
assigned_employees_by_date = defaultdict(set)

# --------------------------------------------------
# Daily Exception Review unique-assignment rule(MES)
# --------------------------------------------------

exception_review_tasks = [
    "Daily Exception Review DSP 10:30",
    "Daily Exception Review PR 13:00",
    "Daily Exception Review FERM 13:00"
]

# Tracks people already assigned to one of the three
# Exception Review tasks on each date
exception_review_assigned_by_date = defaultdict(set)

#--------------------------------------------------------
# Weekly fixed cache
weekly_task_cache = {}

# previous_tier_pair = {
#     "Tier 1 Lead": last_tier1_lead,
#     "Tier 1 Back-up": last_tier1_backup
# }

# All dates
# all_dates = sorted(candidate_base_pd["Date"].dropna().unique())
start_date = candidate_base_pd["Date"].min()
end_date = candidate_base_pd["Date"].max()

all_dates = pd.date_range(
    start=start_date,
    end=end_date,
    freq="D"
)

# Week list
all_weeks = sorted(candidate_base_pd["WeekStart"].dropna().unique())


# COMMAND ----------

# --------------------------------------------------
# Weekly fixed Tier 1 tasks
# --------------------------------------------------
weekly_rules_pd = task_rules_pd[
    (
        task_rules_pd["Assignment_Type"]
        == "WEEKLY_FIXED"
    )
    &
    (
        task_rules_pd["Task"].isin([
            "Tier 1 Lead",
            "Tier 1 Back-up"
        ])
    )
].copy()

# Always process Lead before Back-up
weekly_task_order = {
    "Tier 1 Lead": 1,
    "Tier 1 Back-up": 2
}

weekly_rules_pd["Task_Order"] = (
    weekly_rules_pd["Task"]
    .map(weekly_task_order)
)

weekly_rules_pd = (
    weekly_rules_pd
    .sort_values("Task_Order")
    .reset_index(drop=True)
)

# --------------------------------------------------
# Previous 15 completed Tier 1 weeks
# --------------------------------------------------
# --------------------------------------------------
# Rolling 15-week Tier 1 history
# This is populated chronologically while all_weeks
# is processed. It contains both historical and
# newly automated Tier 1 assignments.
# --------------------------------------------------

recent_tier_week_history = deque(maxlen=15)

print(
    "Starting Tier 1 15-week history:",
    list(recent_tier_week_history)
)

# --------------------------------------------------
# Process one complete Monday-Friday week at a time
# --------------------------------------------------

for week_start in all_weeks:

    week_start = pd.Timestamp(
        week_start
    ).normalize()

    week_data = candidate_base_pd[
        candidate_base_pd["WeekStart"]
        == week_start
    ].copy()

    if week_data.empty:
        continue

    week_data["Date"] = pd.to_datetime(
        week_data["Date"]
    ).dt.normalize()

    # Selected Lead and Back-up for this week
    selected_tier_people_this_week = {}

    # All people actually used as Tier 1 Lead or Back-up
    # in the current week, including historical assignments
    tier_people_used_this_week = set()

    # Dates where the selected Lead is unavailable
    tier1_lead_unavailable_dates = set()

    # --------------------------------------------------
    # Process Lead first, then Back-up
    # --------------------------------------------------

    for _, rule in weekly_rules_pd.iterrows():

        task_name = rule["Task"]

        required_days = [
            "Monday",
            "Tuesday",
            "Wednesday",
            "Thursday",
            "Friday"
        ]

        # --------------------------------------------------
        # Monday-Friday dates for the current week
        # --------------------------------------------------

        required_dates = (
            week_data[
                week_data["DayOfWeek"]
                .isin(required_days)
            ][
                [
                    "Date",
                    "DayOfWeek"
                ]
            ]
            .drop_duplicates()
            .sort_values("Date")
        )

        if required_dates.empty:
            continue

        required_dates["Date"] = pd.to_datetime(
            required_dates["Date"]
        ).dt.normalize()

        # --------------------------------------------------
        # Existing Task Schedule history for this week
        # --------------------------------------------------

        task_week_history = tier_history_pd[
            (
                tier_history_pd["WeekStart"]
                == week_start
            )
            &
            (
                tier_history_pd[task_name]
                .notna()
            )
        ].sort_values("Date")

        historical_week_person = None

        if not task_week_history.empty:

            historical_week_person = str(
                task_week_history.iloc[-1][
                    task_name
                ]
            ).strip()

        # ---------------------------------------------------------------
        # Get eligible and available candidates
        # get_task_candidates should already apply:
        # 1. Allowed codes W / WFH
        # 2. Roles_QA eligibility
        # 3. Present = Yes
        # 4. NAV exclusion 
        # (If the person coming from an long vacation or training then are going to edit the LEave schedule as
        # 'NAV' and those day that person are not going to assign any task - it will give them an come back time 
        # so that they can assign next day for the task from the end and assign the automation next day task as normal)
        # -----------------------------------------------------------------

        weekly_candidates = get_task_candidates(
            candidate_base_pd,
            rule,
            week_start=week_start
        )

        if not weekly_candidates.empty:

            weekly_candidates = (
                weekly_candidates.copy()
            )

            weekly_candidates["Date"] = (
                pd.to_datetime(
                    weekly_candidates["Date"]
                ).dt.normalize()
            )

        # --------------------------------------------------
        # If no candidate exists, preserve historical
        # assignments where available
        # --------------------------------------------------

        if weekly_candidates.empty:

            for _, date_row in required_dates.iterrows():

                assignment_date = pd.Timestamp(
                    date_row["Date"]
                ).normalize()

                historical_employee = (
                    tier_history_lookup.get(
                        (
                            assignment_date,
                            task_name
                        )
                    )
                )

                if historical_employee is not None:

                    tier_people_used_this_week.add(
                        historical_employee
                    )

                results.append({
                    "Date": assignment_date,
                    "Task": task_name,
                    "Employee":
                        historical_employee,
                    "Code": None,
                    "Assignment_Type":
                        rule["Assignment_Type"],
                    "Assignment_Status":
                        (
                            "Assigned"
                            if historical_employee
                            is not None
                            else
                            "No weekly candidate available"
                        ),
                    "Assignment_Source":
                        (
                            "Existing Task Schedule history"
                            if historical_employee
                            is not None
                            else
                            "Not assigned"
                        )
                })

            continue

        # --------------------------------------------------
        # Determine required open days
        # Site-closed dates are not included in the
        # full-week availability requirement.
        # --------------------------------------------------

        if "Is_Site_Closed" in week_data.columns:

            open_required_dates = (
                week_data[
                    (
                        week_data["DayOfWeek"]
                        .isin(required_days)
                    )
                    &
                    (
                        week_data[
                            "Is_Site_Closed"
                        ]
                        .fillna(False)
                        == False
                    )
                ][["Date"]]
                .drop_duplicates()
            )

        else:

            open_required_dates = (
                week_data[
                    week_data["DayOfWeek"]
                    .isin(required_days)
                ][["Date"]]
                .drop_duplicates()
            )

        needed_days = len(
            open_required_dates
        )

        if needed_days == 0:

            needed_days = (
                required_dates["Date"]
                .nunique()
            )

        # --------------------------------------------------
        # Count available days for each eligible employee
        # P, L, T, NA, SL, A, etc. are not counted as
        # available days because they are excluded from
        # weekly_candidates.
        # --------------------------------------------------

        candidate_selection_pool = (
            weekly_candidates[
                weekly_candidates["DayOfWeek"]
                .isin(required_days)
            ]
            .copy()
        )

        if (
            "Is_Site_Closed"
            in candidate_selection_pool.columns
        ):

            candidate_selection_pool = (
                candidate_selection_pool[
                    candidate_selection_pool[
                        "Is_Site_Closed"
                    ]
                    .fillna(False)
                    == False
                ]
                .copy()
            )

        availability_count = (
            candidate_selection_pool
            .groupby("Employee")["Date"]
            .nunique()
            .reset_index(
                name="Available_Required_Days"
            )
        )

        # --------------------------------------------------
        # If no availability remains, preserve history
        # --------------------------------------------------

        if availability_count.empty:

            for _, date_row in required_dates.iterrows():

                assignment_date = pd.Timestamp(
                    date_row["Date"]
                ).normalize()

                historical_employee = (
                    tier_history_lookup.get(
                        (
                            assignment_date,
                            task_name
                        )
                    )
                )

                if historical_employee is not None:

                    tier_people_used_this_week.add(
                        historical_employee
                    )

                results.append({
                    "Date": assignment_date,
                    "Task": task_name,
                    "Employee":
                        historical_employee,
                    "Code": None,
                    "Assignment_Type":
                        rule["Assignment_Type"],
                    "Assignment_Status":
                        (
                            "Assigned"
                            if historical_employee
                            is not None
                            else
                            "No available weekly candidate"
                        ),
                    "Assignment_Source":
                        (
                            "Existing Task Schedule history"
                            if historical_employee
                            is not None
                            else
                            "Not assigned"
                        )
                })

            continue

        # --------------------------------------------------
        # Prefer fully available employees
        # If no fully available employee exists, use
        # partially available employees as fallback.
        # --------------------------------------------------
        # Include all eligible people
        # Full-time people will be preferred later
        # through sorting

        candidate_pool = (
            availability_count.copy()
        )
        # full_week_candidates = (
        #     availability_count[
        #         availability_count[
        #             "Available_Required_Days"
        #         ]
        #         >= needed_days
        #     ]
        #     .copy()
        # )

        # if full_week_candidates.empty:

        #     candidate_pool = (
        #         availability_count.copy()
        #     )

        # else:

        #     candidate_pool = (
        #         full_week_candidates.copy()
        #     )

        # Count previous automated assignments for this role
        candidate_pool["task_count"] = (
            candidate_pool["Employee"]
            .map(
                lambda employee:
                task_assignment_count[
                    task_name
                ].get(employee, 0)
            )
        )

        # --------------------------------------------------
        # Tier 1 Back-up must differ from current Lead
        # --------------------------------------------------

        if task_name == "Tier 1 Back-up":

            current_week_lead = (
                selected_tier_people_this_week
                .get("Tier 1 Lead")
            )

            if current_week_lead is not None:

                different_backup_pool = (
                    candidate_pool[
                        candidate_pool["Employee"]
                        != current_week_lead
                    ]
                    .copy()
                )

                if not different_backup_pool.empty:

                    candidate_pool = (
                        different_backup_pool
                    )

        # --------------------------------------------------
        # Calculate Back-up coverage of dates where
        # the selected Lead is unavailable
        # --------------------------------------------------

        if (
            task_name == "Tier 1 Back-up"
            and len(
                tier1_lead_unavailable_dates
            ) > 0
        ):

            backup_coverage_pd = (
                weekly_candidates[
                    (
                        weekly_candidates["Date"]
                        .isin(
                            tier1_lead_unavailable_dates
                        )
                    )
                    &
                    (
                        weekly_candidates["Code"]
                        .isin([
                            "W",
                            "WFH"
                        ])
                    )
                ]
                .groupby("Employee")["Date"]
                .nunique()
                .reset_index(
                    name=
                    "Lead_Gap_Coverage_Days"
                )
            )

            candidate_pool = (
                candidate_pool.merge(
                    backup_coverage_pd,
                    on="Employee",
                    how="left"
                )
            )

            candidate_pool[
                "Lead_Gap_Coverage_Days"
            ] = (
                candidate_pool[
                    "Lead_Gap_Coverage_Days"
                ]
                .fillna(0)
                .astype(int)
            )

        else:

            candidate_pool[
                "Lead_Gap_Coverage_Days"
            ] = 0

        chosen_employee = None

        # --------------------------------------------------
        # Build combined Tier 1 history from previous
        # 15 completed weeks
        # --------------------------------------------------

        recent_tier_people = set()

        for week_people in recent_tier_week_history:
            recent_tier_people.update(week_people)

        # --------------------------------------------------
        # Preserve existing Tier 1 Task Schedule history
        # The 15-week rule must not overwrite assignments
        # that already exist in the Task Schedule.
        # --------------------------------------------------

        if historical_week_person is not None:

            chosen_employee = historical_week_person
        # --------------------------------------------------
        # Automated selection for blank/future dates
        # --------------------------------------------------

        if chosen_employee is None:

            # Build combined history of both Tier 1 roles
            # from the previous 15 completed weeks
            recent_tier_people = set()

            for week_people in recent_tier_week_history:

                recent_tier_people.update(
                    week_people
                )

            # Include full-time and part-time candidates,
            # but exclude recently used Tier 1 people
            fresh_candidate_pool = (
                candidate_pool[
                    ~candidate_pool["Employee"].isin(
                        recent_tier_people
                    )
                ]
                .copy()
            )

            # Use fresh candidates whenever available
            if not fresh_candidate_pool.empty:

                candidate_pool = (
                    fresh_candidate_pool
                )

            else:

                # Only relax the rule when all eligible and
                # available candidates appear in the history
                print(
                    f"WARNING: 15-week rotation relaxed for "
                    f"{task_name}, week {week_start.date()}. "
                    f"No fresh eligible candidate available."
                )

            # --------------------------------------------------
            # Sort candidates
            # --------------------------------------------------

            if task_name == "Tier 1 Back-up":

                candidate_pool = (
                    candidate_pool
                    .sort_values(
                        [
                            "Lead_Gap_Coverage_Days",
                            "Available_Required_Days",
                            "task_count",
                            "Employee"
                        ],
                        ascending=[
                            False,
                            False,
                            True,
                            True
                        ]
                    )
                )

            else:

                candidate_pool = (
                    candidate_pool
                    .sort_values(
                        [
                            "Available_Required_Days",
                            "task_count",
                            "Employee"
                        ],
                        ascending=[
                            False,
                            True,
                            True
                        ]
                    )
                )

            if candidate_pool.empty:

                for _, date_row in required_dates.iterrows():

                    assignment_date = pd.Timestamp(
                        date_row["Date"]
                    ).normalize()

                    results.append({
                        "Date": assignment_date,
                        "Task": task_name,
                        "Employee": None,
                        "Code": None,
                        "Assignment_Type":
                            rule["Assignment_Type"],
                        "Assignment_Status":
                            "No candidate available",
                        "Assignment_Source":
                            "Not assigned"
                    })

                continue

            chosen_employee = (
                candidate_pool.iloc[0]["Employee"]
            )
        # Store current week selection
        selected_tier_people_this_week[
            task_name
        ] = chosen_employee

        tier_people_used_this_week.add(
            chosen_employee
        )

        # --------------------------------------------------
        # Identify dates when selected Lead is unavailable
        # --------------------------------------------------

        if task_name == "Tier 1 Lead":

            selected_lead_week_rows = (
                week_data[
                    week_data["Employee"]
                    == chosen_employee
                ][
                    [
                        "Date",
                        "Code"
                    ]
                ]
                .copy()
            )

            selected_lead_week_rows[
                "Date"
            ] = pd.to_datetime(
                selected_lead_week_rows[
                    "Date"
                ]
            ).dt.normalize()

            selected_lead_week_rows[
                "Code"
            ] = (
                selected_lead_week_rows[
                    "Code"
                ]
                .astype(str)
                .str.strip()
                .str.upper()
            )

            required_date_values = set(
                pd.to_datetime(
                    required_dates["Date"]
                ).dt.normalize()
            )

            lead_available_dates = set(
                selected_lead_week_rows[
                    selected_lead_week_rows[
                        "Code"
                    ].isin([
                        "W",
                        "WFH"
                    ])
                ]["Date"]
            )

            tier1_lead_unavailable_dates = (
                required_date_values
                - lead_available_dates
            )

        # --------------------------------------------------
        # Preserve existing history and automate blank dates
        # --------------------------------------------------

        for _, date_row in required_dates.iterrows():

            assignment_date = pd.Timestamp(
                date_row["Date"]
            ).normalize()

            historical_employee = (
                tier_history_lookup.get(
                    (
                        assignment_date,
                        task_name
                    )
                )
            )

            if historical_employee is not None:

                employee_for_date = (
                    historical_employee
                )

                assignment_source = (
                    "Existing Task Schedule history"
                )

            else:

                employee_for_date = (
                    chosen_employee
                )

                assignment_source = (
                    "Automation"
                )

            tier_people_used_this_week.add(
                employee_for_date
            )

            day_row = weekly_candidates[
                (
                    weekly_candidates["Date"]
                    == assignment_date
                )
                &
                (
                    weekly_candidates["Employee"]
                    == employee_for_date
                )
            ]

            code_value = None

            if not day_row.empty:

                code_value = (
                    day_row.iloc[0]["Code"]
                )

            results.append({
                "Date": assignment_date,
                "Task": task_name,
                "Employee":
                    employee_for_date,
                "Code": code_value,
                "Assignment_Type":
                    rule["Assignment_Type"],
                "Assignment_Status":
                    "Assigned",
                "Assignment_Source":
                    assignment_source
            })

        # --------------------------------------------------
        # Update weekly tracking
        # --------------------------------------------------

        task_assignment_count[
            task_name
        ][chosen_employee] += 1

        weekly_task_cache[
            (
                week_start,
                task_name
            )
        ] = chosen_employee

        # Do not add Tier 1 people to:
        # assigned_employees_by_date
        # Tier 1 holders may still receive other
        # eligible daily operational tasks.

    # --------------------------------------------------
    # Update rolling history after completing the week
    # --------------------------------------------------

    if tier_people_used_this_week:

        recent_tier_week_history.append(
            tier_people_used_this_week.copy()
        )

# COMMAND ----------

# --------------------------------------------------
# Daily / Specific day / department-only tasks
# --------------------------------------------------

daily_rules_pd = task_rules_pd[
    task_rules_pd["Assignment_Type"].isin([
        "DAILY",
        "SPECIFIC_DAY",
        "SPECIFIC_DAY_MULTI",
        "DEPARTMENT_ONLY"
    ])
].copy()

# --------------------------------------------------
# Diagnostic: current task processing order
# --------------------------------------------------

print("Current daily task processing order:")

for order_number, task_name in enumerate(
    daily_rules_pd["Task"].tolist(),
    start=1
):
    print(
        f"{order_number}. {task_name}"
    )
#---------------------------------------------------
# Daily task priority
# Lower number = processed earlier
# --------------------------------------------------

daily_task_priority = {
    "Release Inbox": 1,
    "Mailbox OCP Review & release": 2,
    "Daily Exception Review DSP 10:30": 3,
    "Daily Exception Review PR 13:00": 4,
    "Daily Exception Review FERM 13:00": 5,
    "Weekly Exception Review PREC 14:00": 6,
    "Weekly Exception Review MP Wed - 11:00": 7,
    "QA Exp Contact": 8,
    "Triager of the day": 9,
    "SAP notifications": 10,
    "QA-ITP": 11,
    "QA-in-QC": 12,
    "QA-in-WH": 13,

}

daily_rules_pd["Task_Priority"] = (
    daily_rules_pd["Task"]
    .map(daily_task_priority)
    .fillna(99)
)

daily_rules_pd = (
    daily_rules_pd
    .sort_values(
        [
            "Task_Priority",
            "Task"
        ],
        ascending=[
            True,
            True
        ]
    )
    .reset_index(drop=True)
)

print("\nUpdated daily task processing order:")

for _, priority_row in daily_rules_pd.iterrows():

    print(
        f'{int(priority_row["Task_Priority"])}. '
        f'{priority_row["Task"]}'
    )
#----------------------------------------------------

from collections import defaultdict
weekly_task_person_count = defaultdict(int)

#-----------------------------------------------------
# Maximum two out of three weekly task rule
# Protected tasks:
# 1. QA Exp Contact
# 2. Triager of the day
# 3. SAP notifications
#
# Tier 1 Lead and Tier 1 Back-up are NOT connected.
# Other tasks (Review & Release Task) are NOT connected.
# -------------------------------------------------------

protected_weekly_tasks = [
    "QA Exp Contact",
    "Triager of the day",
    "SAP notifications"
]

# Key:
# (WeekStart, Employee)
# Value:
# Set of protected task types assigned in that week
protected_tasks_by_employee_week = defaultdict(set)

#-------------------------------------------------------


for dt in all_dates:

    # day_name = candidate_base_pd[
    #     candidate_base_pd["Date"] == dt
    # ]["DayOfWeek"].iloc[0]

    
    day_rows = candidate_base_pd[
        candidate_base_pd["Date"] == dt
    ]

    if day_rows.empty:

        day_name = pd.Timestamp(dt).day_name()

    else:

        day_name = day_rows["DayOfWeek"].iloc[0]

    week_number = pd.Timestamp(dt).isocalendar().week

    for _, rule in daily_rules_pd.iterrows():

        task_name = rule["Task"]
        assignment_type = rule["Assignment_Type"]

        # --------------------------------------------------
        # Weekend handling
        # --------------------------------------------------

        if day_name in ["Saturday", "Sunday"]:

            results.append({
                "Date": dt,
                "Task": task_name,
                "Employee": None,
                "Code": None,
                "Assignment_Type": assignment_type,
                "Assignment_Status": "Weekend"
            })

            continue

        # --------------------------------------------------
        # Department only
        # --------------------------------------------------

        if assignment_type == "DEPARTMENT_ONLY":

            department_value = get_qotf_department(
                day_name,
                week_number
            )

            results.append({
                "Date": dt,
                "Task": task_name,
                "Employee": department_value,
                "Code": "ALL",
                "Assignment_Type": "DEPARTMENT_ONLY",
                "Assignment_Status": "Assigned"
            })

            continue
        
        # --------------------------------------------------
        # Multi specific day logic (QITP)
        # --------------------------------------------------

        if assignment_type == "SPECIFIC_DAY_MULTI":

            allowed_days = rule.get("Specific_Days", [])

            if day_name not in allowed_days:

              continue

        # --------------------------------------------------
        # Specific day logic
        # --------------------------------------------------

        if assignment_type == "SPECIFIC_DAY":

            if day_name != rule.get("Specific_Day", None):
                continue

        # --------------------------------------------------
        # Get eligible candidates
        # --------------------------------------------------

        day_candidates = get_task_candidates(
            candidate_base_pd,
            rule,
            date_value=dt
        )

        if day_candidates.empty:

            results.append({
                "Date": dt,
                "Task": task_name,
                "Employee": None,
                "Code": None,
                "Assignment_Type": assignment_type,
                "Assignment_Status": "No candidate available"
            })

            continue
        
        # --------------------------------------------------
        # Hard rule for Daily Exception Review tasks
        # One employee can receive only one of the three
        # Exception Review tasks on the same date
        # --------------------------------------------------

        if task_name in exception_review_tasks:

            day_candidates = day_candidates[
                ~day_candidates["Employee"].isin(
                    exception_review_assigned_by_date[dt]
                )
            ].copy()

            # If everybody remaining already has an Exception Review,
            # leave this task unassigned rather than duplicate a person
            if day_candidates.empty:

                results.append({
                    "Date": dt,
                    "Task": task_name,
                    "Employee": None,
                    "Code": None,
                    "Assignment_Type": assignment_type,
                    "Assignment_Status":
                        "No unique candidate available for Exception Review"
                })

                continue
        
        # --------------------------------------------------
        # Tracking columns
        # --------------------------------------------------

        day_candidates["task_assignment_count"] = (
            day_candidates["Employee"].map(
                lambda x:
                task_assignment_count[task_name].get(x, 0)
            )
        )
        #--------------------------------------------------------------
        week_start = (
            pd.Timestamp(dt)
            - pd.to_timedelta(
                pd.Timestamp(dt).weekday(),
                unit="D"
            )
        )

        day_candidates["weekly_task_count"] = (
            day_candidates["Employee"].map(
                lambda x:
                weekly_task_person_count.get(
                    (
                        week_start,
                        task_name,
                        x
                    ),
                    0
                )
            )
        )
        # --------------------------------------------------
        # Count how many different protected tasks each
        # employee already has during the current week
        # --------------------------------------------------

        if task_name in protected_weekly_tasks:

            day_candidates["protected_task_type_count"] = (
                day_candidates["Employee"].map(
                    lambda employee:
                    len(
                        protected_tasks_by_employee_week.get(
                            (
                                week_start,
                                employee
                            ),
                            set()
                        )
                    )
                )
            )

        else:

            # All other tasks are unaffected
            day_candidates["protected_task_type_count"] = 0

        #---------------------------------------------
        day_candidates["recently_assigned"] = (
            day_candidates["Employee"].isin(
                list(recent_assignees_by_task[task_name])
            )
        )

        previous_employee = previous_employee_by_task.get(
            task_name,
            None
        )

        day_candidates["same_as_previous"] = (
            day_candidates["Employee"] == previous_employee
        )
        # --------------------------------------------------
        # Check who received this task on the previous
        # occurrence of the same weekday
        # --------------------------------------------------

        previous_weekday_employee = (
            previous_employee_by_task_weekday.get(
                (
                    task_name,
                    day_name
                ),
                None
            )
        )

        day_candidates["same_as_previous_weekday"] = (
            day_candidates["Employee"]
            == previous_weekday_employee
        )


    #-----------------------------------------------------------
        day_candidates["already_assigned_today"] = (
            day_candidates["Employee"].isin(
                assigned_employees_by_date[dt]
            )
        )

        # Avoid same as previous
        pool = day_candidates[
            day_candidates["same_as_previous"] == False
        ].copy()

        if pool.empty:
            pool = day_candidates.copy()

        # Avoid already assigned today
        not_assigned_today_pool = pool[
            pool["already_assigned_today"] == False
        ].copy()

        if not not_assigned_today_pool.empty:
            pool = not_assigned_today_pool.copy()
       
        # ---------------------***-----------------------------
        # Hard rule: the QA Exp Contact person cannot take
        # QA-ITP on the same day
        # ----------------------***----------------------------

        if task_name == "QA-ITP":

            qa_exp_person_today = [
                r["Employee"]
                for r in results
                if r["Task"] == "QA Exp Contact"
                and pd.Timestamp(r["Date"]).normalize() == pd.Timestamp(dt).normalize()
                and r["Employee"] is not None
            ]

            pool = pool[
                ~pool["Employee"].isin(qa_exp_person_today)
            ].copy()

            if pool.empty:

                results.append({
                    "Date": dt,
                    "Task": task_name,
                    "Employee": None,
                    "Code": None,
                    "Assignment_Type": assignment_type,
                    "Assignment_Status":
                        "No candidate: only QA Exp Contact person available"
                })

                continue
        #----------------------------------------------
        # Maximum two DIFFERENT tasks out of:
        # QA Exp Contact / Triager / SAP notifications
        # during the same Monday-Friday week
        # ------------------------------------------------

        if task_name in protected_weekly_tasks:

            allowed_protected_pool = pool[
                pool["Employee"].map(
                    lambda employee:
                    (
                        # Employee currently has fewer than
                        # two different protected tasks
                        len(
                            protected_tasks_by_employee_week.get(
                                (
                                    week_start,
                                    employee
                                ),
                                set()
                            )
                        ) < 2
                    )
                    or
                    (
                        # Current task is already one of the
                        # person's protected task types.
                        # Therefore it does not create a third type.
                        task_name
                        in protected_tasks_by_employee_week.get(
                            (
                                week_start,
                                employee
                            ),
                            set()
                        )
                    )
                )
            ].copy()

            # Hard business rule:
            # Do not assign the third protected task.
            if allowed_protected_pool.empty:

                results.append({
                    "Date": dt,
                    "Task": task_name,
                    "Employee": None,
                    "Code": None,
                    "Assignment_Type": assignment_type,
                    "Assignment_Status":
                        "No candidate available under maximum "
                        "two-of-three weekly task rule"
                })

                continue

            pool = allowed_protected_pool.copy()
        #---------------------------------------------------
        #---------------------------------------------------
        # Avoid the employee who received this same task
        # on the previous occurrence of this weekday
        # Soft rule:
        # only apply when another candidate exists
        # --------------------------------------------------

        different_weekday_employee_pool = pool[
            pool["same_as_previous_weekday"] == False
        ].copy()

        if not different_weekday_employee_pool.empty:
            pool = different_weekday_employee_pool.copy()

        #------------------------------------------------------   

        # Avoid recently assigned
        non_recent_pool = pool[
            pool["recently_assigned"] == False
        ].copy()

        if not non_recent_pool.empty:
            pool = non_recent_pool.copy()
       
        #---------------------------------------- 
        # Selection logic
        #----------------------------------------
        pool = pool.sort_values(
            [
                "protected_task_type_count",
                "weekly_task_count",
                "task_assignment_count",
                "recently_assigned",
                "already_assigned_today",
                "Employee"
            ],
            ascending=[
                True,
                True,
                True,
                True,
                True,
                True
            ]
        )
        #---------------------------------------------------
        chosen = pool.iloc[0]
        chosen_employee = chosen["Employee"]
        
        weekly_task_person_count[
            (
                week_start,
                task_name,
                chosen_employee
            )
        ] += 1
        # --------------------------------------------------
        # Update protected-task weekly tracker
        # --------------------------------------------------

        if task_name in protected_weekly_tasks:

            protected_tasks_by_employee_week[
                (
                    week_start,
                    chosen_employee
                )
            ].add(
                task_name
            )
       #-----------------------------------------------
        results.append({
            "Date": dt,
            "Task": task_name,
            "Employee": chosen_employee,
            "Code": chosen["Code"],
            "Assignment_Type": assignment_type,
            "Assignment_Status": "Assigned"
        })

        # Update tracking
        task_assignment_count[task_name][chosen_employee] += 1

        recent_assignees_by_task[task_name].append(
            chosen_employee
        )

        previous_employee_by_task[task_name] = (
            chosen_employee
        )
        # Remember who received this task on this weekday
        previous_employee_by_task_weekday[
            (
                task_name,
                day_name
            )
        ] = chosen_employee
        #----------------------------------------------------
        #----------------------------------------------------
        assigned_employees_by_date[dt].add(
            chosen_employee
        )

        # --------------------------------------------------
        # Update Exception Review tracker
        # --------------------------------------------------
        # Record the person as already assigned to an Exception Review
        # for this date
        if task_name in exception_review_tasks:

            exception_review_assigned_by_date[dt].add(
                chosen_employee
            )



# COMMAND ----------

# --------------------------------------------------
# Final output
# --------------------------------------------------

all_assignments_pd = pd.DataFrame(results)

# Keep real date for sorting
all_assignments_pd["Date_Sort"] = pd.to_datetime(
    all_assignments_pd["Date"]
)

# Create display date format
all_assignments_pd["Date"] = (
    all_assignments_pd["Date_Sort"]
    .dt.strftime("%A, %B %d, %Y")
    .str.replace(" 0", " ", regex=False)
)

# Sort using real date, not formatted text
all_assignments_pd = all_assignments_pd.sort_values(
    ["Date_Sort", "Task"]
).reset_index(drop=True)

# Convert to Spark
all_assignments_df = spark.createDataFrame(all_assignments_pd)

# --------------------------------------------------
# Create final task schedule
# IMPORTANT:
# Do NOT filter only Assigned, otherwise weekend blank rows may disappear
# --------------------------------------------------

task_schedule_output_df = (
    all_assignments_df
    .groupBy("Date_Sort", "Date")
    .pivot("Task")
    .agg(F.first("Employee"))
    .orderBy("Date_Sort")
)



# COMMAND ----------

# --------------------------------------------------
# Final task order
# --------------------------------------------------
final_task_order = [
    "QA Exp Contact",
    "Tier 1 Lead",
    "Tier 1 Back-up",
    "Triager of the day",
    # "Achterwacht (W)", -> Removed the Achterwacht
    "SAP notifications",
    "QA-ITP",
    "QA-in-QC",
    "QA-in-WH",
    "Release Inbox",
    "Mailbox OCP Review & release",
    "Daily Exception Review DSP 10:30",
    "Daily Exception Review PR 13:00",
    "Daily Exception Review FERM 13:00",
    "Weekly Exception Review PREC 14:00",
    "Weekly Exception Review MP Wed - 11:00"
]

existing_cols = task_schedule_output_df.columns

task_schedule_output_df = task_schedule_output_df.select(
    "Date",
    *[
        c
        for c in final_task_order
        if c in existing_cols
    ]
)

display(task_schedule_output_df)

# COMMAND ----------

# DBTITLE 1,Save task schedule locally
task_schedule_output_pd = task_schedule_output_df.toPandas()

#Locally saving the file 
local_file_path = f"/tmp/Automated_Task_Schedule_{pd.Timestamp.now().strftime('%Y%m%d%H%M%S%f')}.xlsx"

task_schedule_output_pd.to_excel(
    local_file_path,
    index=False
)
print("Excel created:", local_file_path)


# COMMAND ----------

# Upload to SharePoint

with open(local_file_path, "rb") as f:
    file_content = f.read()

target_folder_url = (
    "/teams/QualityEveryonesBusiness/"
    "Shared Documents/General/"
    "2. Leave schedules/"
    "NEW Leave Schedule"
)

# Connect using certificate authentication
ctx = ClientContext(site_url).with_client_certificate(
    TENANT_ID,
    **cert_settings
)

web = ctx.web
ctx.load(web)
ctx.execute_query()

# **lease any existing checkout/lock on the target file before uploading
try:
    existing_file = ctx.web.get_file_by_server_relative_url(
        target_folder_url + "/Automated_Task_Schedule.xlsx"
    )
    existing_file.check_out()
    ctx.execute_query()
except Exception:
    pass

folder = ctx.web.get_folder_by_server_relative_url(
    target_folder_url
)

file = folder.upload_file(
    f"Automated_Task_Schedule_"
    f"{pd.Timestamp.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
    file_content
).execute_query()

print("File uploaded successfully")

# COMMAND ----------

# MAGIC %md
# MAGIC