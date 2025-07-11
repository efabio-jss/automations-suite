# 🧠 Exploratory Data Analysis (EDA) with Pandas – Cheatsheet

This video provides a visual and practical reference for performing **Exploratory Data Analysis (EDA)** using **Pandas** in Python. The image included here summarizes the most essential commands for **loading**, **previewing**, **cleaning**, and **inspecting** datasets.

---

## 📂 1. Data Loading

Use Pandas to import datasets from various file formats:

```python
df = pd.read_csv("file.csv")          # Load CSV file  
df = pd.read_excel("file.xlsx")       # Load Excel file  
df = pd.read_csv("https://...")       # Load from URL  
df = pd.read_json("file.json")        # Load JSON  
df = pd.read_parquet("file.parquet")  # Load Parquet  
```

---

## 👀 2. Data Preview

Quickly inspect the contents of your DataFrame:

```python
df.head(n)        # First N rows  
df.tail(n)        # Last N rows  
df.sample(n)      # Random sample  
df.shape          # Shape of the data (rows, columns)  
df.columns        # Column names  
df.index          # Index information  
```

---

## 🧼 3. Data Cleaning

Clean your dataset for analysis:

```python
df.dropna(inplace=True)                   # Drop missing values  
df['col'] = df['col'].fillna(df['col'].mean())     # Fill missing with mean  
df['col'] = df['col'].fillna(df['col'].mode()[0])  # Fill missing with mode  
df.drop_duplicates(inplace=True)         # Drop duplicates  
df.rename(columns={'Old': 'New'}, inplace=True)    # Rename columns  
df['Date'] = pd.to_datetime(df['Date'])            # Convert to datetime  
df['col'] = df['col'].astype('int')      # Change data type  
```

---

## ℹ️ 4. Data Info

Extract useful metadata from your dataset:

```python
df.info()                    # Data types & non-null info  
df.describe()                # Summary statistics (numeric)  
df.describe(include='object')  # Summary statistics (categorical)  
df.isnull().sum()            # Null counts per column  
(df.isnull().mean() * 100).round(2)  # % missing per column  
```

---

## 📌 Summary

This cheat sheet is ideal for beginners and professionals working with data in Python. It provides a structured workflow for:

- Importing data
- Understanding data structure
- Handling missing values and duplicates
- Getting a statistical overview of the dataset

---

## 💡 Tips

- Use `.copy()` when manipulating slices of a DataFrame to avoid `SettingWithCopyWarning`.
- Combine this EDA with visual tools like `matplotlib` or `seaborn` for deeper insights.
- Use `pip install pandas` if not already installed.

---


