#!/usr/bin/env python
# coding: utf-8

# In[ ]:


import pandas as pd

file_path = r"C:\Users\Isha\Downloads\CFL_External Data Pack_Phase1 (1).xlsx"

# Read with 2 header rows
df_actuals = pd.read_excel(file_path, sheet_name="Data Pack - Actual Bookings", header=[0,1])

# Flatten multi-level columns
df_actuals.columns = [
    (str(col[0]).strip() + "_" + str(col[1]).strip()).replace("Unnamed", "").strip("_")
    for col in df_actuals.columns
]

# Clean column names
df_actuals.columns = df_actuals.columns.str.replace("  ", " ").str.strip()

# Drop completely empty columns
df_actuals = df_actuals.dropna(axis=1, how='all')

# Preview
df_actuals.head()


# In[ ]:


df_actuals.shape


# In[ ]:


#dropping the forecast cols from df_actuals 


# In[ ]:


# ==============================
# DROP FY26 Q2 FORECAST BLOCK
# ==============================

cols_to_drop = [col for col in df_actuals.columns if "forecasted units" in col.lower()]

df_actuals = df_actuals.drop(columns=cols_to_drop)

print("Dropped columns:", cols_to_drop)

df_actuals.head()


# In[ ]:


#making df_forecasts from the data pack sheet 


# In[ ]:


import pandas as pd

file_path = r"C:\Users\Isha\Downloads\CFL_External Data Pack_Phase1 (1).xlsx"

# ==============================
# 1. READ + FLATTEN
# ==============================

df = pd.read_excel(file_path, sheet_name="Data Pack - Actual Bookings", header=[0,1])

df.columns = [
    (str(col[0]).strip() + "_" + str(col[1]).strip()).replace("Unnamed", "").strip("_")
    for col in df.columns
]

df.columns = df.columns.str.strip()

# ==============================
# 2. FIND REQUIRED COLUMNS
# ==============================

cost_col = [col for col in df.columns if "cost rank" in col.lower()][0]
product_col = [col for col in df.columns if "product name" in col.lower()][0]
lifecycle_col = [col for col in df.columns if "life cycle" in col.lower()][0]

demand_col = [col for col in df.columns if "demand" in col.lower() and "forecast" in col.lower()][0]
marketing_col = [col for col in df.columns if "marketing" in col.lower()][0]
ds_col = [col for col in df.columns if "data science" in col.lower()][0]

# ==============================
# 3. CREATE CLEAN DATAFRAME
# ==============================

df_forecasts = df[[cost_col, product_col, lifecycle_col, demand_col, marketing_col, ds_col]].copy()

# ==============================
# 4. RENAME COLUMNS
# ==============================

df_forecasts.columns = [
    "cost_rank",
    "product",
    "life_cycle",
    "demand_forecast",
    "marketing_forecast",
    "data_science_forecast"
]

# ==============================
# 5. REMOVE EMPTY / HEADER ROWS
# ==============================

df_forecasts = df_forecasts[df_forecasts["cost_rank"].notna()]

# ==============================
# 6. FINAL OUTPUT
# ==============================

df_forecasts.head()
df_forecasts.shape


# In[ ]:


df_forecasts.head()


# In[ ]:


#making the big deals df 


# In[ ]:


import pandas as pd

file_path =r"C:\Users\Isha\Downloads\CFL_External Data Pack_Phase1 (1).xlsx"

# ==============================
# 1. READ WITH MULTI-HEADER
# ==============================

df_big = pd.read_excel(file_path, sheet_name="Big Deal", header=[0,1])

# ==============================
# 2. FLATTEN COLUMNS
# ==============================

df_big.columns = [
    (str(col[0]).strip() + "_" + str(col[1]).strip()).replace("Unnamed", "").strip("_")
    for col in df_big.columns
]

df_big.columns = df_big.columns.str.strip()

# ==============================
# 3. IDENTIFY KEY COLUMNS
# ==============================

cost_col = [col for col in df_big.columns if "cost rank" in col.lower()][0]
product_col = [col for col in df_big.columns if "plid" in col.lower()][0]

# ==============================
# 4. KEEP ONLY BIG DEAL COLUMNS
# ==============================

bigdeal_cols = [
    col for col in df_big.columns 
    if "big deals" in col.lower()
]

# Final subset
df_big = df_big[[cost_col, product_col] + bigdeal_cols].copy()

# Rename keys
df_big = df_big.rename(columns={
    cost_col: "cost_rank",
    product_col: "product"
})

# ==============================
# 5. REMOVE INVALID ROWS
# ==============================

df_big = df_big[df_big["cost_rank"].notna()]

# ==============================
# 6. CONVERT WIDE → LONG
# ==============================

df_big_long = df_big.melt(
    id_vars=["cost_rank", "product"],
    var_name="quarter",
    value_name="big_deal"
)

# ==============================
# 7. CLEAN QUARTER COLUMN
# ==============================

df_big_long["quarter"] = df_big_long["quarter"].str.extract(r'(\d{4}Q\d)')

# ==============================
# 8. HANDLE MISSING VALUES
# ==============================

df_big_long["big_deal"] = pd.to_numeric(df_big_long["big_deal"], errors="coerce")
df_big_long["big_deal"] = df_big_long["big_deal"].fillna(0)

# ==============================
# 9. FINAL CLEANUP
# ==============================

df_big_long["cost_rank"] = df_big_long["cost_rank"].astype(int)

df_big_long = df_big_long.sort_values(by=["cost_rank", "quarter"])

# ==============================
# 10. OUTPUT
# ==============================

print(df_big_long.head())
print("Shape:", df_big_long.shape)


# In[ ]:


df_big_long.head()


# In[ ]:


df_big_long.shape


# In[ ]:





# In[ ]:


#making the scms df 


# In[ ]:


import pandas as pd

file_path = r"C:\Users\Isha\Downloads\CFL_External Data Pack_Phase1 (1).xlsx"

# ==============================
# 1. READ + FLATTEN
# ==============================

df = pd.read_excel(file_path, sheet_name="SCMS", header=[0,1])

df.columns = [
    (str(col[0]).strip() + "_" + str(col[1]).strip()).replace("Unnamed", "").strip("_")
    for col in df.columns
]

df.columns = df.columns.str.strip()

# ==============================
# 2. IDENTIFY REQUIRED COLUMNS
# ==============================

cost_col = [col for col in df.columns if "cost rank" in col.lower()][0]
product_col = [col for col in df.columns if "plid" in col.lower()][0]
coverage_col = [col for col in df.columns if "coverage" in col.lower()][0]

# 🔥 IMPORTANT FIX: pick SCMS Units columns
scms_cols = [
    col for col in df.columns 
    if "scms units" in col.lower()
]

print("SCMS columns:", scms_cols)

# ==============================
# 3. CREATE DATAFRAME
# ==============================

df_scms = df[[cost_col, product_col, coverage_col] + scms_cols].copy()

# Rename
df_scms = df_scms.rename(columns={
    cost_col: "cost_rank",
    product_col: "product",
    coverage_col: "coverage"
})

# ==============================
# 4. REMOVE INVALID ROWS
# ==============================

df_scms = df_scms[df_scms["cost_rank"].notna()]

# ==============================
# 5. OUTPUT
# ==============================

print(df_scms.head())
print("Shape:", df_scms.shape)


# In[ ]:


df_scms.head(10)


# In[ ]:


#making the vms df


# In[ ]:


#VMS Cleaning 


# In[ ]:


import pandas as pd

file_path = r"C:\Users\Isha\Downloads\CFL_External Data Pack_Phase1 (1).xlsx"

# ==============================
# 1. READ + FLATTEN
# ==============================

df = pd.read_excel(file_path, sheet_name="VMS", header=[0,1])

df.columns = [
    (str(col[0]).strip() + "_" + str(col[1]).strip()).replace("Unnamed", "").strip("_")
    for col in df.columns
]

df.columns = df.columns.str.strip()

# ==============================
# 2. IDENTIFY REQUIRED COLUMNS
# ==============================

cost_col = [col for col in df.columns if "cost rank" in col.lower()][0]
product_col = [col for col in df.columns if "plid" in col.lower()][0]
vms_name_col = [col for col in df.columns if "vms top name" in col.lower()][0]

# 🔥 IMPORTANT: pick VMS Units columns
vms_cols = [
    col for col in df.columns 
    if "vms units" in col.lower()
]

print("VMS columns:", vms_cols)

# ==============================
# 3. CREATE DATAFRAME
# ==============================

df_vms = df[[cost_col, product_col, vms_name_col] + vms_cols].copy()

# Rename
df_vms = df_vms.rename(columns={
    cost_col: "cost_rank",
    product_col: "product",
    vms_name_col: "vms_category"
})

# ==============================
# 4. REMOVE INVALID ROWS
# ==============================

df_vms = df_vms[df_vms["cost_rank"].notna()]

# ==============================
# 5. OUTPUT
# ==============================

print(df_vms.head())
print("Shape:", df_vms.shape)


# In[ ]:


df_vms.head()


# In[ ]:


#converting the actual bookings to long format 


# In[ ]:


df_actuals.head()


# In[ ]:


#dropping the forecasting col from df_actuals and storing it in df_actuals_clean


# In[ ]:


# ==============================
# DROP ALL FORECAST COLUMNS
# ==============================

df_actuals_clean = df_actuals[
    [col for col in df_actuals.columns if "forecast" not in col.lower()]
].copy()

# ==============================
# CHECK RESULT
# ==============================

print("Remaining columns:")
print(df_actuals_clean.columns)

df_actuals_clean.head()


# In[ ]:


df_actuals_clean.head()


# In[ ]:


#remooving the rows 32- 64 from the df_actuals_clean 


# In[ ]:


df_actuals_clean = df_actuals_clean.drop(df_actuals_clean.index[32:64])


# In[ ]:


df_actuals_clean.tail()


# In[ ]:


df_actuals_clean.tail()


# In[ ]:


#converting the df_actuals_clean to df_actuals_long 


# In[ ]:


df_actuals_clean.head()


# In[ ]:


# ==============================
# STEP 1: EXTRACT QUARTER NAMES
# ==============================

# Take first row (contains FY values)
quarter_row = df_actuals_clean.iloc[0]

# Convert to list
quarter_values = quarter_row.tolist()

# ------------------------------
# STEP 2: BUILD NEW COLUMN NAMES
# ------------------------------

new_columns = []

for i in range(len(df_actuals_clean.columns)):
    if i < 3:
        # Keep first 3 columns as is
        new_columns.append(df_actuals_clean.columns[i])
    else:
        # Replace with actual FY values
        new_columns.append(quarter_values[i])

# Assign new columns
df_actuals_clean.columns = new_columns

# Drop first row (it was header)
df_actuals_clean = df_actuals_clean[1:].reset_index(drop=True)

# ------------------------------
# STEP 3: MELT USING POSITIONS
# ------------------------------

df_actuals_long = df_actuals_clean.melt(
    id_vars=df_actuals_clean.columns[:3],   # first 3 columns
    value_vars=df_actuals_clean.columns[3:], # rest are quarters
    var_name="quarter",
    value_name="bookings"
)

# ------------------------------
# STEP 4: RENAME CLEANLY
# ------------------------------

df_actuals_long = df_actuals_long.rename(columns={
    df_actuals_long.columns[0]: "cost_rank",
    df_actuals_long.columns[1]: "product",
    df_actuals_long.columns[2]: "life_cycle"
})

# ------------------------------
# STEP 5: CLEAN VALUES
# ------------------------------

df_actuals_long["quarter"] = df_actuals_long["quarter"].astype(str).str.replace(" ", "")
df_actuals_long["bookings"] = pd.to_numeric(df_actuals_long["bookings"], errors="coerce")
df_actuals_long["cost_rank"] = pd.to_numeric(df_actuals_long["cost_rank"], errors="coerce")

# Drop empty rows
df_actuals_long = df_actuals_long.dropna(subset=["bookings"])

# Sort
df_actuals_long = df_actuals_long.sort_values(by=["cost_rank", "quarter"])



# ------------------------------
# FINAL OUTPUT
# ------------------------------

print(df_actuals_long.head())
print("Shape:", df_actuals_long.shape)


# In[ ]:


# ==============================
# FIX QUARTER COLUMN COMPLETELY
# ==============================

# Sort first (IMPORTANT)
df_actuals_long = df_actuals_long.sort_values(by=["cost_rank"]).reset_index(drop=True)

# Number of unique quarters per product
n_quarters = df_actuals_long.groupby("product")["quarter"].transform("count")

# Create correct quarter sequence (based on your data)
quarter_sequence = [
    "2023Q2","2023Q3","2023Q4",
    "2024Q1","2024Q2","2024Q3","2024Q4",
    "2025Q1","2025Q2","2025Q3","2025Q4",
    "2026Q1"
]

# Assign quarters per product
df_actuals_long["quarter"] = (
    df_actuals_long.groupby("product")
    .cumcount()
    .map(lambda x: quarter_sequence[x] if x < len(quarter_sequence) else None)
)

print(df_actuals_long.head())


# In[ ]:


df_actuals_long.shape


# In[ ]:


df_actuals_long.head(100)


# In[ ]:


df_actuals_long = df_actuals_long.sort_values(by=["cost_rank", "product", "quarter"])


# In[ ]:


df_actuals_long.head()


# In[ ]:


df_big_long.head()


# In[ ]:


# df_actuals_long has been made 
#df_big_long is already in long format 


# In[ ]:


# making long format of df_scms


# In[ ]:


df_scms.head()


# In[ ]:


# ==============================
# STEP 2: SCMS → LONG (FIXED)
# ==============================

# 1. Identify ID columns
id_cols = ["cost_rank", "product", "coverage"]

# 2. Value columns
value_cols = [col for col in df_scms.columns if col not in id_cols]

# 🔥 3. Create quarter mapping manually
quarters = [
    "2023Q1","2023Q2","2023Q3","2023Q4",
    "2024Q1","2024Q2","2024Q3","2024Q4",
    "2025Q1","2025Q2","2025Q3","2025Q4",
    "2026Q1"
]

# 4. MELT
df_scms_long = df_scms.melt(
    id_vars=id_cols,
    value_vars=value_cols,
    var_name="temp",
    value_name="scms"
)

# 5. Assign correct quarter using order
df_scms_long["quarter"] = df_scms_long.groupby(
    ["cost_rank", "product", "coverage"]
).cumcount().map(lambda x: quarters[x])

# ==============================
# 6. CLEAN VALUES
# ==============================

df_scms_long["scms"] = pd.to_numeric(df_scms_long["scms"], errors="coerce").fillna(0)

# ==============================
# 7. AGGREGATE
# ==============================

df_scms_long = df_scms_long.groupby(
    ["cost_rank", "product", "quarter"],
    as_index=False
)["scms"].sum()

# ==============================
# 8. OUTPUT
# ==============================

print(df_scms_long.head())
print("Shape:", df_scms_long.shape)


# In[ ]:


df_scms_long.head()


# In[ ]:


#make the long format of df_long and aggregate it 


# In[ ]:


# ==============================
# STEP 3: VMS → LONG FORMAT
# ==============================

# 1. Identify ID columns
id_cols = ["cost_rank", "product", "vms_category"]

# 2. Value columns
value_cols = [col for col in df_vms.columns if col not in id_cols]

# 🔥 3. Quarter mapping (same as SCMS)
quarters = [
    "2023Q1","2023Q2","2023Q3","2023Q4",
    "2024Q1","2024Q2","2024Q3","2024Q4",
    "2025Q1","2025Q2","2025Q3","2025Q4",
    "2026Q1"
]

# 4. MELT
df_vms_long = df_vms.melt(
    id_vars=id_cols,
    value_vars=value_cols,
    var_name="temp",
    value_name="vms"
)

# 5. Assign correct quarter using order
df_vms_long["quarter"] = df_vms_long.groupby(
    ["cost_rank", "product", "vms_category"]
).cumcount().map(lambda x: quarters[x])

# ==============================
# 6. CLEAN VALUES
# ==============================

df_vms_long["vms"] = pd.to_numeric(df_vms_long["vms"], errors="coerce").fillna(0)

# ==============================
# 7. AGGREGATE (IMPORTANT)
# ==============================

df_vms_long = df_vms_long.groupby(
    ["cost_rank", "product", "quarter"],
    as_index=False
)["vms"].sum()

# ==============================
# 8. SORT
# ==============================

df_vms_long = df_vms_long.sort_values(by=["cost_rank", "quarter"])

# ==============================
# 9. OUTPUT
# ==============================

print(df_vms_long.head())
print("Shape:", df_vms_long.shape)


# In[ ]:


df_vms_long.head()


# In[ ]:


df_actuals_long.head()


# In[ ]:


print(df_actuals_long["quarter"].unique())
print(df_scms_long["quarter"].unique())


# In[ ]:


df_actuals_long.head()


# In[ ]:


df_big_long.head()


# In[ ]:


df_scms_long.head()


# In[ ]:


df_vms_long.head()


# In[ ]:


df_actuals_long.shape


# In[ ]:


df_big_long.shape


# In[ ]:


df_scms_long.shape


# In[ ]:


df_vms_long.shape


# In[ ]:


#merging for master table 


# In[ ]:


# ==============================
# STEP 1: CLEAN KEYS (IMPORTANT)
# ==============================

for df in [df_actuals_long, df_big_long, df_scms_long, df_vms_long]:
    df["product"] = df["product"].astype(str).str.strip()
    df["quarter"] = df["quarter"].astype(str).str.strip()
    df["cost_rank"] = pd.to_numeric(df["cost_rank"], errors="coerce")

# ==============================
# STEP 2: START WITH ACTUALS
# ==============================

df_master = df_actuals_long.copy()

# ==============================
# STEP 3: MERGE BIG DEAL
# ==============================

df_master = df_master.merge(
    df_big_long,
    on=["cost_rank", "product", "quarter"],
    how="left"
)

# ==============================
# STEP 4: MERGE SCMS
# ==============================

df_master = df_master.merge(
    df_scms_long,
    on=["cost_rank", "product", "quarter"],
    how="left"
)

# ==============================
# STEP 5: MERGE VMS
# ==============================

df_master = df_master.merge(
    df_vms_long,
    on=["cost_rank", "product", "quarter"],
    how="left"
)

# ==============================
# STEP 6: FILL MISSING VALUES
# ==============================

df_master["big_deal"] = df_master["big_deal"].fillna(0)
df_master["scms"] = df_master["scms"].fillna(0)
df_master["vms"] = df_master["vms"].fillna(0)

# ==============================
# STEP 7: FINAL SORT
# ==============================

df_master = df_master.sort_values(by=["product", "quarter"]).reset_index(drop=True)

# ==============================
# OUTPUT
# ==============================

print(df_master.head())
print(df_master.shape)


# In[ ]:


df_master.head(30)


# In[ ]:


df_master.head()


# In[ ]:


df_master.shape


# In[ ]:


#everything is okay till here 


# In[ ]:


#time based features - quarter number 


# In[ ]:


df_master["quarter_num"] = (
    df_master["quarter"]
    .astype(str)
    .str.extract(r'Q(\d)', expand=False)
    .astype(float)
    .astype("Int64")   # safe integer (handles NaN)
)


# In[ ]:


df_master.head()


# In[ ]:


# time trend 


# In[ ]:


df_master = df_master.sort_values(["cost_rank", "quarter"])
df_master["time_index"] = df_master.groupby("cost_rank").cumcount()


# In[ ]:


# big deal indicator - buisness feature 


# In[ ]:


df_master["has_bigdeal"] = (df_master["big_deal"] > 0).astype(int)


# In[ ]:


df_master["rolling_mean_3"] = (
    df_master.groupby("cost_rank")["bookings"]
    .transform(lambda x: x.rolling(3, min_periods=1).mean())
)

df_master["rolling_std_3"] = (
    df_master.groupby("cost_rank")["bookings"]
    .transform(lambda x: x.rolling(3, min_periods=1).std())
)


# In[ ]:


df_master.tail(30)


# In[ ]:


df_master.shape


# In[ ]:


df_master.shape


# In[ ]:


df_master = df_master.sort_values(by=["cost_rank", "product", "quarter"])


# In[ ]:


df_master["rolling_std_3"] = df_master["rolling_std_3"].fillna(0)


# In[ ]:


df_master.head(10)


# In[ ]:


df_master.tail(5)


# In[ ]:


df_master.shape


# In[ ]:


#scms ratio 


# In[ ]:


df_master["scms_ratio"] = df_master["scms"] / (df_master["bookings"] + 1)


# In[ ]:


#vms ratio 


# In[ ]:


df_master["vms_ratio"] = df_master["vms"] / (df_master["bookings"] + 1)


# In[ ]:


df_master["bigdeal_scms"] = df_master["big_deal"] * df_master["scms"]


# In[ ]:


df_master["bigdeal_ratio"] = df_master["big_deal"] / (df_master["bookings"] + 1) 


# In[ ]:


df_master["scms_pressure"] = df_master["scms"] / (df_master["bookings"] + 1)  


# In[ ]:


df_master["vms_intensity"] = df_master["vms"] / (df_master["bookings"] + 1)  


# In[ ]:


#product behaviour 


# In[ ]:


df_master["product_avg"] = df_master.groupby("cost_rank")["bookings"].transform("mean") 


# In[ ]:


df_master["deviation"] = df_master["bookings"] - df_master["product_avg"]


# In[ ]:


df_master["scms_vms"] = df_master["scms"] * df_master["vms"]


# In[ ]:


df_master.shape


# In[ ]:


# base forecasting - time series 


# In[ ]:


pip install statsmodels prophet


# In[ ]:


from statsmodels.tsa.statespace.sarimax import SARIMAX

def sarima_forecast(df, product_id):
    
    # Filter one product
    temp = df[df["cost_rank"] == product_id].copy()
    temp = temp.sort_values("quarter")

    y = temp["bookings"].values

    # SARIMA model (simple version)
    model = SARIMAX(y, order=(1,1,0), seasonal_order=(0,0,0,0))
    model_fit = model.fit(disp=False)

    forecast = model_fit.forecast(steps=1)

    return forecast[0]


# In[ ]:


print(sarima_forecast(df_master,1))


# In[ ]:


# basline model 2 - moving average 


# In[ ]:


def moving_avg_forecast(df, product_id):
    
    temp = df[df["cost_rank"] == product_id].copy()
    temp = temp.sort_values("quarter")

    return temp["bookings"].tail(4).mean()


# In[ ]:


print(moving_avg_forecast(df_master,1))


# In[ ]:


# baseline model 3- naive forecast 


# In[ ]:


def naive_forecast(df, product_id):
    
    temp = df[df["cost_rank"] == product_id].copy()
    temp = temp.sort_values("quarter")

    return temp["bookings"].iloc[-1]


# In[ ]:


print(naive_forecast(df_master,1))


# In[ ]:


# combining Sarima, Moving Average and Naive forecast for all products (baseline forecast of all products) 


# In[ ]:


results = []

products = df_master["cost_rank"].unique()

for p in products:
    
    # SARIMA (with error handling)
    try:
        sarima_pred = sarima_forecast(df_master, p)
    except:
        sarima_pred = None

    # Moving Average
    try:
        moving_avg_pred = moving_avg_forecast(df_master, p)
    except:
        moving_avg_pred = None

    # Naive (last value)
    try:
        naive_pred = naive_forecast(df_master, p)
    except:
        naive_pred = None

    results.append({
        "cost_rank": p,
        "sarima": sarima_pred,
        "moving_avg": moving_avg_pred,
        "naive": naive_pred
    })

df_predictions = pd.DataFrame(results)

print(df_predictions.head())


# In[ ]:


df_predictions = df_predictions.round(2)


# In[ ]:


print(df_predictions.head())


# In[ ]:


df_predictions = df_predictions.sort_values("cost_rank")


# In[ ]:


df_predictions.head()


# In[ ]:


# ML model - preparation 


# In[ ]:


df_master.head()


# In[ ]:


# Drop non-feature columns
df_ml = df_master.copy()

# Remove columns not needed for ML
df_ml = df_ml.drop(columns=["product", "quarter"])

# Target
y = df_ml["bookings"]

# Features
X = df_ml.drop(columns=["bookings"])


# In[ ]:


df_ml.head()


# In[ ]:


# train test split time based 


# In[ ]:


# keep last quarter for testing 

df_ml = df_master.copy()
df_ml = df_ml.sort_values(["cost_rank", "quarter"])


# In[ ]:


df_ml["quarter"].isna().sum()


# In[ ]:


df_ml[df_ml["quarter"].isna()]


# In[ ]:


df_ml["quarter"].value_counts(dropna=False)


# In[ ]:


import numpy as np

# Step 1: Convert to string and clean spaces
df_ml["quarter"] = df_ml["quarter"].astype(str).str.strip()

# Step 2: Replace "None" (string) with actual NaN
df_ml["quarter"] = df_ml["quarter"].replace("None", np.nan)

# Step 3: Drop those rows
df_ml = df_ml.dropna(subset=["quarter"]).reset_index(drop=True)


# In[ ]:


last_quarter = df_ml["quarter"].max()
print("Last quarter:", last_quarter)


# In[ ]:


df_ml.shape


# In[ ]:


X.head()


# In[ ]:


df_master.head(60)


# In[ ]:


df_ml.isna().sum()


# In[ ]:


invalid = df_ml[~df_ml["quarter"].str.contains(r"\d{4}Q\d", na=False)]
print(invalid["quarter"].unique())


# In[ ]:


# ==============================
# STEP 5: EXTRACT YEAR & QUARTER NUMBER
# ==============================

df_ml["year"] = df_ml["quarter"].str[:4].astype(int)
df_ml["q_num"] = df_ml["quarter"].str.extract(r'Q(\d)').astype(int)


# In[ ]:


# ==============================
# STEP 6: FIND LAST QUARTER
# ==============================

last_year = df_ml["year"].max()
last_q = df_ml[df_ml["year"] == last_year]["q_num"].max()

last_quarter = f"{last_year}Q{last_q}"

print("Last quarter:", last_quarter)


# In[ ]:


# ==============================
# STEP 7: TRAIN-TEST SPLIT
# ==============================

train_df = df_ml[df_ml["quarter"] != last_quarter]
test_df = df_ml[df_ml["quarter"] == last_quarter]


# In[ ]:


# ==============================
# STEP 8: FEATURES & TARGET
# ==============================

X_train = train_df.drop(columns=["bookings"])
y_train = train_df["bookings"]

X_test = test_df.drop(columns=["bookings"])
y_test = test_df["bookings"]


# In[ ]:


# ==============================
# STEP 9: DROP NON-NUMERIC COLUMNS
# ==============================

cols_to_drop = ["product", "quarter", "life_cycle"]

X_train = X_train.drop(columns=cols_to_drop)
X_test = X_test.drop(columns=cols_to_drop)


# In[ ]:


# ==============================
# STEP 10: FINAL CLEAN (VERY IMPORTANT)
# ==============================

X_train = X_train.fillna(0)
X_test = X_test.fillna(0)


# In[ ]:


# ==============================
# FINAL CHECK
# ==============================

print("Train shape:", X_train.shape)
print("Test shape:", X_test.shape)
print("Test quarter:", test_df["quarter"].unique())


# In[ ]:


# linear regression 


# In[ ]:


from sklearn.linear_model import LinearRegression
from sklearn.metrics import mean_absolute_error

lr = LinearRegression()
lr.fit(X_train, y_train)

y_pred_lr = lr.predict(X_test)

print("Linear Regression MAE:", mean_absolute_error(y_test, y_pred_lr))


# In[ ]:


# figuring out the reason for low MAE 


# In[ ]:


from sklearn.linear_model import LinearRegression
from sklearn.metrics import mean_absolute_error

# ==============================
# STEP 1: DEFINE MODEL
# ==============================

model = LinearRegression()

# ==============================
# STEP 2: TRAIN MODEL
# ==============================

model.fit(X_train, y_train)

# ==============================
# STEP 3: PREDICT
# ==============================

# Train predictions
y_train_pred = model.predict(X_train)

# Test predictions
y_test_pred = model.predict(X_test)

# ==============================
# STEP 4: CALCULATE MAE
# ==============================

train_mae = mean_absolute_error(y_train, y_train_pred)
test_mae = mean_absolute_error(y_test, y_test_pred)

# ==============================
# OUTPUT
# ==============================

print("Train MAE:", train_mae)
print("Test MAE:", test_mae)


# In[ ]:


print(X_train.columns.tolist())


# In[ ]:


from sklearn.linear_model import LinearRegression
from sklearn.metrics import mean_absolute_error

# ==============================
# STEP 1: REMOVE LEAKAGE FEATURES
# ==============================

cols_to_remove = [
    "rolling_mean_3",
    "rolling_std_3",
    "product_avg",
    "deviation"
]

X_train_clean = X_train.drop(columns=cols_to_remove, errors="ignore")
X_test_clean = X_test.drop(columns=cols_to_remove, errors="ignore")

# ==============================
# STEP 2: FILL NaNs (SAFE)
# ==============================

X_train_clean = X_train_clean.fillna(0)
X_test_clean = X_test_clean.fillna(0)

# ==============================
# STEP 3: TRAIN MODEL
# ==============================

model = LinearRegression()
model.fit(X_train_clean, y_train)

# ==============================
# STEP 4: PREDICT
# ==============================

y_train_pred = model.predict(X_train_clean)
y_test_pred = model.predict(X_test_clean)

# ==============================
# STEP 5: EVALUATE
# ==============================

train_mae = mean_absolute_error(y_train, y_train_pred)
test_mae = mean_absolute_error(y_test, y_test_pred)

print("Train MAE:", train_mae)
print("Test MAE:", test_mae)


# In[ ]:


y_test_pred = np.maximum(y_test_pred, 0)


# In[ ]:


import pandas as pd
import numpy as np

# Create results dataframe
results = pd.DataFrame({
    "Actual": y_test,
    "Predicted": y_test_pred
})

# Absolute Error
results["Error"] = abs(results["Actual"] - results["Predicted"])

# Accuracy (%)
results["Accuracy (%)"] = 100 * (1 - results["Error"] / (results["Actual"] + 1))

# Clip negative accuracy (optional)
results["Accuracy (%)"] = results["Accuracy (%)"].clip(lower=0)

print(results.head(300))


# In[ ]:


print("Average Accuracy:", results["Accuracy (%)"].mean())


# In[ ]:


df_master["bookings"].mean


# In[ ]:


# XG boost 


# In[ ]:


get_ipython().system('pip install xgboost')


# In[ ]:


from xgboost import XGBRegressor

xgb = XGBRegressor(
    n_estimators=100,
    max_depth=4,
    learning_rate=0.1,
    random_state=42
)

xgb.fit(X_train, y_train)

y_pred_xgb = xgb.predict(X_test)

print("XGBoost MAE:", mean_absolute_error(y_test, y_pred_xgb))


# In[ ]:


import pandas as pd
import numpy as np
from sklearn.metrics import mean_absolute_error

# ==============================
# STEP 1: CREATE RESULTS TABLE
# ==============================

results_xgb = pd.DataFrame({
    "Actual": y_test,
    "Predicted": y_pred_xgb
})

# ==============================
# STEP 2: ERROR
# ==============================

results_xgb["Error"] = abs(results_xgb["Actual"] - results_xgb["Predicted"])

# ==============================
# STEP 3: ACCURACY (%)
# ==============================

results_xgb["Accuracy (%)"] = 100 * (
    1 - results_xgb["Error"] / (results_xgb["Actual"] + 1)
)

# Prevent negative accuracy
results_xgb["Accuracy (%)"] = results_xgb["Accuracy (%)"].clip(lower=0)

# ==============================
# STEP 4: AVERAGE ACCURACY
# ==============================

avg_accuracy_xgb = results_xgb["Accuracy (%)"].mean()

# ==============================
# OUTPUT
# ==============================

print(results_xgb.head(10))
print("\nXGBoost Average Accuracy:", avg_accuracy_xgb)


# In[ ]:


from sklearn.metrics import mean_absolute_error

def evaluate_model(df, forecast_function):
    
    errors = []

    products = df["cost_rank"].unique()

    for p in products:
        
        temp = df[df["cost_rank"] == p].copy()
        temp = temp.sort_values("quarter")

        if len(temp) < 5:
            continue

        train = temp.iloc[:-1]
        test = temp.iloc[-1:]

        try:
            pred = forecast_function(train, p)
            actual = test["bookings"].values[0]

            error = abs(actual - pred)
            errors.append(error)

        except:
            continue

    return sum(errors) / len(errors)


# In[ ]:


# sarima , naive and moving avg codes 


# In[ ]:


y_pred_naive = []

for product in test_df["cost_rank"].unique():
    
    train_prod = train_df[train_df["cost_rank"] == product]
    test_prod = test_df[test_df["cost_rank"] == product]
    
    if len(train_prod) > 0:
        last_value = train_prod["bookings"].iloc[-1]
        preds = [last_value] * len(test_prod)
    else:
        preds = [0] * len(test_prod)
    
    y_pred_naive.extend(preds)

y_pred_naive = np.array(y_pred_naive)


# In[ ]:


#moving avg 


# In[ ]:


y_pred_ma = []

for product in test_df["cost_rank"].unique():
    
    train_prod = train_df[train_df["cost_rank"] == product]
    test_prod = test_df[test_df["cost_rank"] == product]
    
    if len(train_prod) >= 3:
        avg = train_prod["bookings"].iloc[-3:].mean()
    elif len(train_prod) > 0:
        avg = train_prod["bookings"].mean()
    else:
        avg = 0
    
    preds = [avg] * len(test_prod)
    y_pred_ma.extend(preds)

y_pred_ma = np.array(y_pred_ma)


# In[ ]:





# In[ ]:


import pandas as pd
import numpy as np

def calculate_accuracy(y_true, y_pred, model_name):
    
    results = pd.DataFrame({
        "Actual": y_true,
        "Predicted": y_pred
    })
    
    # Error
    results["Error"] = abs(results["Actual"] - results["Predicted"])
    
    # Accuracy %
    results["Accuracy (%)"] = 100 * (
        1 - results["Error"] / (results["Actual"] + 1)
    )
    
    # Clip negative values
    results["Accuracy (%)"] = results["Accuracy (%)"].clip(lower=0)
    
    avg_accuracy = results["Accuracy (%)"].mean()
    
    print(f"\n{model_name} Average Accuracy:", avg_accuracy)
    
    return results, avg_accuracy


# In[ ]:


results_ma, avg_ma = calculate_accuracy(
    y_test, y_pred_ma, "Moving Average"
)

print(results_ma.head())


# In[ ]:


results_naive, avg_naive = calculate_accuracy(
    y_test, y_pred_naive, "Naive"
)

print(results_naive.head())


# In[ ]:


print(df_master["bookings"].describe())


# In[ ]:


def evaluate_mape(df, forecast_function):
    
    errors = []

    products = df["cost_rank"].unique()

    for p in products:
        temp = df[df["cost_rank"] == p].copy()
        temp = temp.sort_values("quarter")

        if len(temp) < 5:
            continue

        train = temp.iloc[:-1]
        test = temp.iloc[-1:]

        try:
            pred = forecast_function(train, p)
            actual = test["bookings"].values[0]

            if actual != 0:
                error = abs((actual - pred) / actual)
                errors.append(error)

        except:
            continue

    return sum(errors) / len(errors) * 100


# In[ ]:


def weighted_moving_avg(df, product_id):
    
    temp = df[df["cost_rank"] == product_id].copy()
    temp = temp.sort_values("quarter")

    values = temp["bookings"].tail(4).values
    weights = [0.1, 0.2, 0.3, 0.4]

    return sum(v*w for v, w in zip(values, weights))


# In[ ]:


weighted_mape = evaluate_mape(df_master, weighted_moving_avg)

print("Weighted Moving Avg MAPE:", weighted_mape)


# In[ ]:


df_master.head()


# In[ ]:


df_master.describe()


# In[ ]:


get_ipython().system('pip install lightgbm')


# In[ ]:


from lightgbm import LGBMRegressor
import numpy as np
from sklearn.metrics import mean_absolute_error

df_lgb = df_ml.copy()

# Remove non-numeric columns
cols_to_drop = ["product", "quarter", "life_cycle"]
df_lgb = df_lgb.drop(columns=cols_to_drop, errors="ignore")

# Remove leakage features
leak_cols = ["rolling_mean_3", "rolling_std_3", "product_avg", "deviation"]
df_lgb = df_lgb.drop(columns=leak_cols, errors="ignore")

# Fill NaNs
df_lgb = df_lgb.fillna(0)


# In[ ]:


y = np.log1p(df_lgb["bookings"])
X = df_lgb.drop(columns=["bookings"])


# In[ ]:


# Get last quarter
last_quarter = df_ml["quarter"].max()

train_idx = df_ml["quarter"] != last_quarter
test_idx = df_ml["quarter"] == last_quarter

X_train = X[train_idx]
X_test = X[test_idx]

y_train = y[train_idx]
y_test = y[test_idx]


# In[ ]:


lgb = LGBMRegressor(
    n_estimators=200,
    learning_rate=0.05,
    max_depth=5,
    random_state=42
)

lgb.fit(X_train, y_train)


# In[ ]:


y_pred_log = lgb.predict(X_test)

y_pred = np.expm1(y_pred_log)
y_test_actual = np.expm1(y_test)


# In[ ]:


mae = mean_absolute_error(y_test_actual, y_pred)

print("LightGBM MAE:", mae)


# In[ ]:


print(y_test_actual.describe())


# In[ ]:


#accuracy of light bgm 


# In[ ]:


import pandas as pd
import numpy as np

# ==============================
# STEP 1: CREATE RESULTS TABLE
# ==============================

results_lgb = pd.DataFrame({
    "Actual": y_test_actual,
    "Predicted": y_pred
})

# ==============================
# STEP 2: CALCULATE ERROR
# ==============================

results_lgb["Error"] = abs(results_lgb["Actual"] - results_lgb["Predicted"])

# ==============================
# STEP 3: CALCULATE ACCURACY (%)
# ==============================

results_lgb["Accuracy (%)"] = 100 * (
    1 - results_lgb["Error"] / (results_lgb["Actual"] + 1)
)

# Avoid negative accuracy
results_lgb["Accuracy (%)"] = results_lgb["Accuracy (%)"].clip(lower=0)

# ==============================
# STEP 4: AVERAGE ACCURACY
# ==============================

avg_accuracy_lgb = results_lgb["Accuracy (%)"].mean()

# ==============================
# OUTPUT
# ==============================

print(results_lgb.head(10))
print("\nLightGBM Average Accuracy:", avg_accuracy_lgb)


# In[ ]:


#comparison 


# In[ ]:


print("Linear Accuracy:", results["Accuracy (%)"].mean())
print("XGBoost Accuracy:", avg_accuracy_xgb)
print("LightGBM Accuracy:", avg_accuracy_lgb)


# In[ ]:


print(y_pred[:10])


# In[ ]:


df_ml = df_ml.drop(columns=[
    "product_avg",
    "deviation"
])


# In[ ]:


mape = (82.63 / df_master["bookings"].mean()) * 100
print("Approx MAPE:", mape)


# In[ ]:


#creating df_predictions


# In[ ]:


feature_cols = [
    'cost_rank', 'big_deal', 'scms', 'vms', 'quarter_num', 'has_bigdeal',
    'scms_ratio', 'vms_ratio', 'bigdeal_scms', 'bigdeal_ratio',
    'scms_pressure', 'vms_intensity', 'scms_vms', 'year', 'q_num'
]


# In[ ]:


print("\n===== DF_ML COLUMNS =====")
print(df_ml.columns.tolist())

print("\n===== FEATURE COLS =====")
print(feature_cols)

print("\n===== X_train COLUMNS =====")
print(X_train.columns.tolist())

print("\n===== X_test COLUMNS =====")
print(X_test.columns.tolist())

print("\n===== MODELS =====")
print("Linear model:", type(model))
print("XGB model:", type(xgb))
print("LGB model:", type(lgb))

print("\n===== SAMPLE LAST ROW =====")
print(df_ml.tail(1))

print("\n===== CHECK MISSING FEATURE COLS IN DF_ML =====")
missing_cols = [col for col in feature_cols if col not in df_ml.columns]
print("Missing:", missing_cols)

print("\n===== CHECK EXTRA COLS IN DF_ML =====")
extra_cols = [col for col in df_ml.columns if col not in feature_cols]
print("Extra:", extra_cols)


# In[ ]:


#finding df_predictions_ml


# In[ ]:


# XGBoost features (used during training)
xgb_features = xgb.get_booster().feature_names

# LightGBM features (used during training)
lgb_features = X_train.columns   # (this is your clean dataset)


# In[ ]:


ml_results = []

for p in df_ml["cost_rank"].unique():
    
    temp = df_ml[df_ml["cost_rank"] == p].sort_values("quarter")
    last_row = temp.tail(1)
    
    # ==============================
    # XGBOOST INPUT
    # ==============================
    X_xgb = last_row.reindex(columns=xgb_features, fill_value=0)
    xgb_pred = xgb.predict(X_xgb)[0]
    
    # ==============================
    # LIGHTGBM INPUT
    # ==============================
    X_lgb = last_row.reindex(columns=lgb_features, fill_value=0)
    
    lgb_pred_log = lgb.predict(X_lgb)[0]
    lgb_pred = np.expm1(lgb_pred_log)
    
    # ==============================
    # STORE
    # ==============================
    ml_results.append({
        "cost_rank": p,
        "xgb_pred": xgb_pred,
        "lgb_pred": lgb_pred
    })

df_predictions_ml = pd.DataFrame(ml_results)

print(df_predictions_ml.head())


# In[ ]:


# creating df_predictions_ml


# In[ ]:


import pandas as pd
import numpy as np

ts_results = []

for p in df_ml["cost_rank"].unique():
    
    # Get product-wise data (sorted properly)
    temp = df_ml[df_ml["cost_rank"] == p].sort_values("quarter")
    
    # ==============================
    # SAFE SERIES EXTRACTION
    # ==============================
    
    series = temp["bookings"].fillna(0).values
    
    # ==============================
    # NAIVE PREDICTION
    # ==============================
    
    if len(series) > 0:
        naive_pred = series[-1]   # last value
    else:
        naive_pred = 0
    
    # ==============================
    # MOVING AVERAGE (LAST 3)
    # ==============================
    
    if len(series) >= 3:
        ma_pred = np.mean(series[-3:])
    elif len(series) > 0:
        ma_pred = np.mean(series)
    else:
        ma_pred = 0
    
    # ==============================
    # STORE RESULTS
    # ==============================
    
    ts_results.append({
        "cost_rank": p,
        "naive_pred": naive_pred,
        "ma_pred": ma_pred
    })

# Final dataframe
df_predictions_ts = pd.DataFrame(ts_results)

print(df_predictions_ts.head())


# In[ ]:


df_ml.columns


# In[ ]:


# merging df_predictions and df_predictions_ml 


# In[ ]:


df_final = df_predictions.merge(
    df_predictions_ml,
    on="cost_rank"
)


# In[ ]:


df_final.head()


# In[ ]:


# emsemble 


# In[ ]:


acc_xgb = 88.744
acc_ma = 75.59
acc_naive = 71.02
acc_lgb = 74.67


# In[ ]:


total = acc_xgb + acc_ma + acc_naive + acc_lgb

w_xgb = acc_xgb / total
w_ma = acc_ma / total
w_naive = acc_naive / total
w_lgb = acc_lgb / total

print("Weights:")
print("XGB:", w_xgb)
print("MA:", w_ma)
print("Naive:", w_naive)
print("LGB:", w_lgb)


# In[ ]:


df_final.columns


# In[ ]:


df_final["final_forecast"] = (
    w_xgb   * df_final["xgb_pred"] +
    w_lgb   * df_final["lgb_pred"] +
    w_ma    * df_final["moving_avg"] +
    w_naive * df_final["naive"]
)


# In[ ]:


df_final.head()


# In[ ]:


df_final.head()


# In[ ]:


df_predictions.head()


# In[ ]:


df_master.head()


# In[ ]:


df_master.columns


# In[ ]:


df_final.columns


# In[ ]:


# big deal effect 


# In[ ]:


bigdeal_prob = df_master.groupby("cost_rank")["big_deal"].apply(lambda x: (x > 0).mean()).reset_index()
bigdeal_prob.columns = ["cost_rank", "bigdeal_prob"]


# In[ ]:


bigdeal_avg = df_master[df_master["big_deal"] > 0].groupby("cost_rank")["big_deal"].mean().reset_index()
bigdeal_avg.columns = ["cost_rank", "avg_bigdeal"]


# In[ ]:


df_final = df_final.merge(bigdeal_prob, on="cost_rank", how="left")
df_final = df_final.merge(bigdeal_avg, on="cost_rank", how="left")

df_final[["bigdeal_prob", "avg_bigdeal"]] = df_final[["bigdeal_prob", "avg_bigdeal"]].fillna(0)


# In[ ]:


df_final["expected_bigdeal"] = df_final["bigdeal_prob"] * df_final["avg_bigdeal"]


# In[ ]:


df_final["forecast_after_bigdeal"] = (
    df_final["final_forecast"] + df_final["expected_bigdeal"]
)


# In[ ]:


df_final.head()


# In[ ]:


df_master.columns


# In[ ]:


#final forecast code 


# In[ ]:


df_ml = df_ml.sort_values(["cost_rank", "year", "q_num"])


# In[ ]:


xgb_features = xgb.get_booster().feature_names
lgb_features = X_train.columns


# In[ ]:


final_results = []

for p in df_ml["cost_rank"].unique():
    
    temp = df_ml[df_ml["cost_rank"] == p].sort_values(["year", "q_num"])
    
    # Latest available row (FY26Q1)
    last_row = temp.tail(1)
    
    # ==============================
    # ML MODELS
    # ==============================
    
    X_xgb = last_row.reindex(columns=xgb_features, fill_value=0)
    xgb_pred = xgb.predict(X_xgb)[0]
    
    X_lgb = last_row.reindex(columns=lgb_features, fill_value=0)
    lgb_pred = np.expm1(lgb.predict(X_lgb)[0])
    
    # ==============================
    # TIME SERIES BASELINES
    # ==============================
    
    series = temp["bookings"].fillna(0).values
    
    # Naive
    naive_pred = series[-1] if len(series) > 0 else 0
    
    # Moving Average
    if len(series) >= 3:
        ma_pred = np.mean(series[-3:])
    else:
        ma_pred = np.mean(series) if len(series) > 0 else 0
    
    # ==============================
    # ENSEMBLE (USE YOUR WEIGHTS)
    # ==============================
    
    final_pred = (
        w_xgb * xgb_pred +
        w_lgb * lgb_pred +
        w_ma * ma_pred +
        w_naive * naive_pred
    )
    
    # ==============================
    # STORE RESULT
    # ==============================
    
    final_results.append({
        "cost_rank": p,
        "forecast_quarter": "FY26Q2",
        "xgb_pred": xgb_pred,
        "lgb_pred": lgb_pred,
        "ma_pred": ma_pred,
        "naive_pred": naive_pred,
        "final_forecast": final_pred
    })

# Final dataframe
df_fy26q2_forecast = pd.DataFrame(final_results)

print(df_fy26q2_forecast.head())


# In[ ]:


df_forecasts.columns



# In[ ]:


# comparing with df_forecasts 


# In[ ]:


df_compare = df_fy26q2_forecast.merge(
    df_forecasts,
    on="cost_rank",
    how="inner"
)

print(df_compare.head())


# In[ ]:


df_compare["diff_demand"] = abs(df_compare["final_forecast"] - df_compare["demand_forecast"])

df_compare["diff_marketing"] = abs(df_compare["final_forecast"] - df_compare["marketing_forecast"])

df_compare["diff_ds"] = abs(df_compare["final_forecast"] - df_compare["data_science_forecast"])


# In[ ]:


print("Avg diff vs Demand:", df_compare["diff_demand"].mean())
print("Avg diff vs Marketing:", df_compare["diff_marketing"].mean())
print("Avg diff vs Data Science:", df_compare["diff_ds"].mean())


# In[ ]:


import numpy as np

# ==============================
# SAFE PERCENTAGE DIFFERENCE
# ==============================

df_compare["pct_diff_demand"] = (
    abs(df_compare["final_forecast"] - df_compare["demand_forecast"]) 
    / (np.maximum(df_compare["demand_forecast"], df_compare["final_forecast"]) + 1)
) * 100

df_compare["pct_diff_marketing"] = (
    abs(df_compare["final_forecast"] - df_compare["marketing_forecast"]) 
    / (np.maximum(df_compare["marketing_forecast"], df_compare["final_forecast"]) + 1)
) * 100

df_compare["pct_diff_ds"] = (
    abs(df_compare["final_forecast"] - df_compare["data_science_forecast"]) 
    / (np.maximum(df_compare["data_science_forecast"], df_compare["final_forecast"]) + 1)
) * 100

# ==============================
# AVERAGE % DIFFERENCE
# ==============================

avg_demand = df_compare["pct_diff_demand"].mean()
avg_marketing = df_compare["pct_diff_marketing"].mean()
avg_ds = df_compare["pct_diff_ds"].mean()

print("Avg % Diff vs Demand:", avg_demand)
print("Avg % Diff vs Marketing:", avg_marketing)
print("Avg % Diff vs Data Science:", avg_ds)

# ==============================
# OPTIONAL: CLEAN TABLE OUTPUT
# ==============================

comparison = pd.DataFrame({
    "Comparison": ["Demand", "Marketing", "Data Science"],
    "Avg % Difference": [avg_demand, avg_marketing, avg_ds]
})

print("\nFinal Comparison Table:")
print(comparison)


# In[ ]:


df_fy26q2_forecast = df_fy26q2_forecast.merge(
    df_ml[["cost_rank", "product"]].drop_duplicates(),
    on="cost_rank",
    how="left"
)

df_fy26q2_forecast[["cost_rank", "product", "final_forecast"]].sort_values("cost_rank")


# In[1]:


import os
os.getcwd()


# In[ ]:




