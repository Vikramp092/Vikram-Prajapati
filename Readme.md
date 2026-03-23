# 💰 Australian Tax Depreciation Calculator

A web app to calculate depreciation deductions for Australian tax purposes.

## 🚀 Features
- Diminishing Value and Prime Cost methods
- ATO effective life guidelines
- Single asset and batch processing
- Export to Excel with formatted schedules
- Clean UI using Streamlit

## 📊 Depreciation Methods

### Diminishing Value Method (Default)
- Depreciation rate applied to remaining asset value
- Front-loads deductions in early years
- Generally provides higher tax benefits initially

### Prime Cost Method
- Depreciation rate applied to original asset cost
- Provides consistent deductions over time
- Better for assets with stable value

## 🎯 Asset Categories & Effective Lives

The calculator includes ATO-recommended effective lives:
- **Computers/Office Equipment**: 4 years
- **Motor Vehicles**: 8 years
- **Plant & Equipment**: 10 years
- **Buildings**: 40 years
- **Custom Rate**: Enter your own depreciation rate

## ▶️ How to Run

```bash
pip install -r requirements.txt
streamlit run app.py
```

## 📋 Usage

### Single Asset Mode
1. Enter asset details (name, cost, purchase date)
2. Select depreciation method
3. Choose asset type or enter custom rate
4. Click "Calculate Depreciation"
5. View schedule and export to Excel

### Asset Register Mode
1. Upload Excel/CSV file with columns:
   - `Asset_Name`: Name/description of asset
   - `Cost`: Original purchase cost
   - `Purchase_Date`: Date of purchase (YYYY-MM-DD format)
   - `Method`: 'diminishing_value' or 'prime_cost'
   - `Rate`: Annual depreciation rate (as percentage, e.g., 25 for 25%)
2. Click "Calculate All Depreciation"
3. View summary and export detailed schedules

### Sample Data
A sample asset register file (`sample_assets.csv`) is included for testing the batch processing feature.
