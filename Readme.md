```markdown
# From Raw Data to Insights — Interactive Sales Performance Dashboard in Excel

GitHub only renders markdown when the file name and extension are **exactly** `README.md`.

---

### ⚠️ 3. You Uploaded It as a Code Snippet or Inside Another File
You must upload it as the main **README file** in your repo root folder — not inside another markdown code block or file.

---

Here’s the **final plain markdown content** again — ready to copy-paste directly (no code fences, no extra formatting):  

---

# From Raw Data to Insights — Interactive Sales Performance Dashboard in Excel

**Built with:** Excel 365 (Web) · Superstore Sales Dataset ([Kaggle](https://www.kaggle.com/datasets/vivek468/superstore-dataset-final))

---

## 📊 Overview
This project demonstrates an **end-to-end analytics workflow** that transforms raw sales data into actionable business insights using **Microsoft Excel 365 (Web)**. The workflow includes **data cleaning**, **analysis**, **forecasting**, and **interactive dashboard creation**.

**Key Highlights:**
- Data ingestion & cleaning  
- Advanced Excel formulas: `XLOOKUP`, `SUMIFS`, `INDEX-MATCH`, `IF`, `AND`, `OR`  
- PivotTables, PivotCharts, and KPI visualizations  
- Forecasting using `FORECAST.LINEAR`  
- Interactive slicers and dynamic filtering  
- Automation via recorded macros  
- ChatGPT-assisted formula creation and data summaries  

---

## 🗂️ Repository Structure

| Folder | Description |
|--------|-------------|
| `data/` | Dataset files and download instructions |
| `excel/` | Main Excel workbook for cleaning, analysis, and dashboard |
| `media/` | Demo videos and screenshots of the dashboard |
| `docs/` | Documentation: walkthroughs, formula explanations, project notes |
| `scripts/` | Optional scripts for automation or data cleaning (Python/Power Query) |

---

## 📁 Files of Interest
- `excel/Superstore_Sales_Workbook.xlsx` — Main workbook with raw, cleaned, and dashboard sheets  
- `media/demo_video.mp4` — Demo video showing raw data → insights  
- `docs/dashboard_walkthrough.md` — Detailed explanation of dashboard logic and formulas  
- `data/kaggle_instructions.md` — Guide to downloading the dataset from Kaggle  

---

## 🔧 How to Reproduce

1. **Download Dataset**  
   - Visit [Superstore Dataset – Final](https://www.kaggle.com/datasets/vivek468/superstore-dataset-final)  
   - Download and save as `Superstore_Sales.csv` in the `data/` folder

2. **Open Workbook**  
   - Open `excel/Superstore_Sales_Workbook.xlsx` in **Excel 365 (Web)**

3. **Review Sheets**  
   - `Raw_Data` — original data  
   - `Cleaned_Data` — formatted and cleaned data  
   - `Analytics` — calculations and formulas  
   - `Dashboard` — final visualization

4. **Optional**  
   - Review `scripts/data_cleaning/clean_data.py` (if available)  
   - Explore macros for auto-refresh and automation  

---

## 🧮 Key Excel Formulas

| Purpose | Formula Example | Description |
|---------|----------------|-------------|
| Total Sales | `=SUMIFS(Sales, Region, "West")` | Total sales for the West region |
| Profit Margin | `=Profit / Sales` | Calculates profit margin |
| Customer Segmentation | `=IF(Sales>1000, "High Value", "Regular")` | Categorizes customers by purchase value |
| Forecasting | `=FORECAST.LINEAR(Month, Known_Ys, Known_Xs)` | Predicts next period sales or profit |
| Dynamic Lookup | `=XLOOKUP(Customer_ID, Customer_Table[ID], Customer_Table[Name])` | Retrieves customer names |
| Error Handling | `=IFERROR(Formula, "Not Found")` | Replaces errors with readable text |
| Text Cleanup | `=TRIM(PROPER(TEXT(A2,"@")))` | Cleans inconsistent text values |
| Conditional Flag | `=IF(AND(Profit<0,Discount>0),"Loss Due to Discount","Profit")` | Flags transactions affecting profitability |

---

## 🎯 Key Insights
- Sales and profit trends visualized by **region, category, and time**  
- Forecast models predict short-term performance  
- KPIs: Total Sales, Profit, Average Margin, Forecast Accuracy  
- Slicers enable interactive exploration by **Region**, **Segment**, and **Category**  

---

## 📝 Key Learnings
- **End-to-end workflow:** Raw → Cleaning → Analysis → Forecasting → Dashboard  
- **Advanced Excel:** Use of `XLOOKUP`, `INDEX-MATCH`, `SUMIFS`, and dynamic referencing  
- **Forecasting & Automation:** Leveraging Excel functions and macros for time-series insights  
- **Data Storytelling:** Using KPI visuals and slicers for intuitive insights  
- **AI in Analytics:** ChatGPT-assisted formula generation and data summary creation  

---

## 👩‍💻 Author
**Name:** Shriraksha  

**Dataset Source:** [Superstore Dataset – Final (Kaggle)](https://www.kaggle.com/datasets/vivek468/superstore-dataset-final)

---

## 🔖 Hashtags
#DataAnalytics #ExcelDashboard #BusinessIntelligence #DataVisualization #SuperstoreDataset #Forecasting #KPIAnalysis #Automation #ExcelForBusiness #AIinAnalytics

---

✅ After you paste this:
- Make sure your file name is **README.md**
- Upload it to the **root folder** of your repository
- Refresh the GitHub page  

It will now display exactly like a **formatted project README** (headings, emojis, bold text, tables, etc.).  

Would you like me to add a **header banner or badges (Excel | Kaggle | BI)** to make it visually more professional?
