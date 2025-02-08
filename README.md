# 📊 **Excel Sports Event Management Data Analysis**  

This project involves analyzing sports event data for **XYZ Co Pvt Ltd**, focusing on **data cleaning, analysis, and report generation**. The dataset includes **athlete details, country information, and sports metadata**.  

---

## **📌 Project Overview**  
- **🎯 Objective**: Systematize membership rosters and generate reports for international sports events.  
- **📂 Data Sources**:  
  - **`SPORTSMEN`**: Athlete profiles (names, birthdates, countries, languages, salaries, etc.).  
  - **`LOCATION`**: Maps country codes to names and languages.  
  - **`SPORT`**: Associates sports with indoor/outdoor locations.  

---

## **🛠 Tasks & Solutions**  

### **🔹 Stage 1: Data Cleaning (Question 1)**  

#### **1️⃣ Standardize Names**  
Combine `PREFIX`, `FIRSTNAME`, and `LASTNAME` into `FULLNAME` (uppercase):  
```excel
=UPPER(CONCAT(C2, " ", D2, " ", F2))
```  

#### **2️⃣ Fetch Country & Language**  
Use **HLOOKUP** and **XLOOKUP** to populate `COUNTRY NAME` and `LANGUAGE` from the `LOCATION` sheet:  
```excel
=HLOOKUP(J2, LOCATION!$A$2:$M$3, 2, 0)
=XLOOKUP(K2, LOCATION!$B$3:$M$3, LOCATION!$B$1:$M$1)
```  

#### **3️⃣ Generate Email**  
Create emails based on language (`.org` for English, `.com` otherwise):  
```excel
=LOWER(F2 & "." & D2 & IF(L2="English", "@xyz.org", "@xyz.com"))
```  

#### **4️⃣ Format Data**  
- **MEMBER ID** as a 3-digit format:  
  ```excel
  =TEXT(A2, "000")
  ```  
- **Birthdate** as `dd mmm' yyyy` format:  
  ```excel
  (Custom format: dd mmm' yyyy)
  ```  
- **Salary** in thousands with conditional decimals:  
  ```excel
  =IF(S2<100000, TEXT(S2/1000, "0.00") & " k", TEXT(S2/1000, "0.0") & " k")
  ```  

---

### **🔹 Stage 2: Data Analysis (Question 2)**  

#### **1️⃣ Pivot Table**  
Summarize athlete counts by **COUNTRY** and **GENDER** (cell `B3` in `ANALYSIS` sheet):  

- **Rows**: `COUNTRY NAME`  
- **Columns**: `GENDER`  
- **Values**: Count of `MEMBER ID`  
- **Grand Totals** removed  

#### **2️⃣ Summary Table with Functions**  

- **Extract distinct genders** using **Remove Duplicates + Transpose**.  
- **Count athletes per country/gender** using `COUNTIFS`:  
  ```excel
  =COUNTIFS(SPORTSMEN!$I$2:$I$51, $H$4, SPORTSMEN!$K$2:$K$51, $G5)
  ```  

---

### **🔹 Stage 3: Report Generation (Question 3)**  

#### **1️⃣ Pivot Table Report**  

- **Fields**: `MEMBER ID`, `FULL NAME`, `EMAIL`, `GENDER`, `YEAR OF BIRTH`, `COUNTRY NAME`, `LANGUAGE`, `SPORTS`  
- **Layout**: Tabular form, **collapse buttons removed**  
- **Filter**: Add **SPORT LOCATION** slicer at `A1`  

---

## **📈 Key Insights**  

✔ **Data Relationships**: Athlete demographics and sports preferences vary by country (e.g., Germany dominates Alpine Skiing).  
✔ **Formatting**: Ensured consistency in **IDs, dates, and financial metrics** for readability.  
✔ **Automation**: Used **dynamic formulas (XLOOKUP, COUNTIFS)** to reduce manual updates.  
✔ **Tools Used**: **Excel Formulas, Pivot Tables, Conditional Formatting**  

---

🔗 **This project highlights advanced Excel data analytics skills, focusing on automation and insightful reporting.** 🚀
