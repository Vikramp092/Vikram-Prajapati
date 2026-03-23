# 📊 Comprehensive Depreciation Calculator

A professional Excel-based tool for calculating depreciation under both **Companies Act, 2013** (Schedule II - SLM method) and **Income Tax Act, 1961** (WDV method). Perfect for Chartered Accountants, Tax Professionals, and Finance Teams.

## 🎯 Key Features

### ✅ **Dual Compliance**
- **Companies Act, 2013**: Straight Line Method (SLM) with Schedule II rates
- **Income Tax Act, 1961**: Written Down Value (WDV) method with block concept

### ✅ **Advanced Calculations**
- Automatic pro-rata depreciation for assets used <180 days
- Block-wise depreciation for Income Tax compliance
- Additional depreciation calculations
- Variance analysis between book and tax depreciation

### ✅ **Professional Features**
- Excel formulas (not just values) for dynamic calculations
- Data validation with dropdown menus
- Professional formatting with borders and headers
- Error handling and input validation
- Scalable for unlimited assets

## 📋 Sheet Structure

### 1. **Input Sheet**
Enter all asset details with the following fields:
- **Asset Name**: Description of the asset
- **Asset Category**: Plant & Machinery, Furniture, Building, etc. (dropdown)
- **Date of Purchase**: When asset was acquired
- **Date Put to Use**: When asset was first used
- **Original Cost**: Purchase price of asset
- **Residual Value**: Scrap value (for Companies Act)
- **Useful Life**: As per Schedule II (for Companies Act)
- **Depreciation Rate**: As per Income Tax Rules
- **Block of Asset**: Block A, B, C, D (dropdown)
- **Opening WDV**: For existing assets
- **Additions During Year**: New purchases
- **Deletions During Year**: Assets sold/disposed

### 2. **Companies Act Sheet**
Calculates depreciation under Companies Act, 2013:
- **SLM Method**: (Cost - Residual Value) / Useful Life
- **Pro-rata Calculation**: For assets used <180 days
- **Automatic Days Calculation**: Based on dates entered
- **Total Summary**: With grand totals

### 3. **Income Tax Act Sheet**
Calculates depreciation under Income Tax Act, 1961:
- **WDV Method**: Applied to block of assets
- **Block Concept**: Assets grouped by depreciation rates
- **Additional Depreciation**: 50% for assets used <180 days
- **Closing WDV**: Opening WDV + Additions - Deletions - Depreciation

### 4. **Dashboard Sheet**
- **Summary Overview**: Total depreciation under both acts
- **Variance Analysis**: Difference between book and tax depreciation
- **Instructions**: Step-by-step usage guide

## 🚀 How to Use

### Step 1: Enter Asset Data
1. Open the `Depreciation_Calculator.xlsx` file
2. Go to the **"Input"** sheet
3. Enter asset details in the provided columns
4. Use dropdown menus for Asset Category and Block of Asset

### Step 2: Review Calculations
1. Go to **"Companies Act"** sheet to see SLM calculations
2. Go to **"Income Tax Act"** sheet to see WDV calculations
3. All calculations update automatically when you change input data

### Step 3: Generate Reports
1. Use the **"Dashboard"** sheet for summary reports
2. Copy-paste data to your tax working papers
3. Export sheets as needed for audit purposes

## 📊 Calculation Logic

### Companies Act, 2013 (SLM)
```
Annual Depreciation = (Cost - Residual Value) / Useful Life
Pro-rata Depreciation = Annual Depreciation × (Days Used / 365)
If Days Used < 180: Apply pro-rata
If Days Used ≥ 180: Apply full annual depreciation
```

### Income Tax Act, 1961 (WDV)
```
Block WDV = Opening WDV + Additions - Deletions
Normal Depreciation = Block WDV × Depreciation Rate
Additional Depreciation = Normal Depreciation × 50% (if used <180 days)
Total Depreciation = Normal + Additional
Closing WDV = Block WDV - Total Depreciation
```

## 🎨 Sample Data Included

The Excel file comes pre-loaded with sample assets:
- Office Computer (Plant & Machinery)
- Office Furniture (Furniture & Fittings)
- Factory Machine (Plant & Machinery)
- Building (Building)
- Motor Vehicle (Motor Cars)

## 📈 Advanced Features

### ✅ **Auto-Detection**
- Automatically detects assets used <180 days
- Calculates exact number of days used
- Applies appropriate depreciation rates

### ✅ **Error Handling**
- Data validation for required fields
- Dropdown menus prevent invalid entries
- Formula error checking

### ✅ **Scalability**
- Add unlimited assets by copying rows
- Formulas automatically extend
- Block calculations update dynamically

### ✅ **Professional Output**
- Currency formatting (₹)
- Percentage formatting
- Date formatting (DD/MM/YYYY)
- Professional borders and headers

## 🔧 Technical Details

### Excel Formulas Used
- `IF()` - Conditional calculations
- `DAYS()` - Date calculations
- `SUM()` - Total calculations
- `COUNTA()` - Asset counting
- `AVERAGE()` - Rate calculations

### Data Validation
- Asset Category dropdown
- Block of Asset dropdown
- Date format validation
- Numeric input validation

## 📋 Compliance Standards

### Companies Act, 2013
- Schedule II useful lives
- SLM methodology
- Pro-rata depreciation rules
- Residual value considerations

### Income Tax Act, 1961
- Section 32 depreciation
- Block of assets concept
- Additional depreciation rules
- WDV methodology

## 🎯 Use Cases

- **Chartered Accountants**: Tax depreciation working papers
- **Finance Teams**: Book vs tax depreciation reconciliation
- **Auditors**: Depreciation compliance checking
- **Business Owners**: Tax planning and compliance
- **Tax Consultants**: Client depreciation calculations

## 📞 Support

The Excel file is self-contained with:
- Built-in instructions
- Formula explanations
- Error handling
- Sample data for reference

## ⚠️ Important Notes

1. **Backup Data**: Always backup your input data
2. **Verify Calculations**: Cross-check with manual calculations for first use
3. **Update Rates**: Depreciation rates may change - update as per latest rules
4. **Professional Use**: Consult with tax experts for complex scenarios
5. **Regular Updates**: Keep the file updated with new assets and disposals

---

**Created for professional use in audit and tax compliance.** 🇮🇳