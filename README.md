# 🇮🇳 Professional Depreciation Calculator

[![Python](https://img.shields.io/badge/Python-3.8+-blue.svg)](https://www.python.org/)
[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](https://opensource.org/licenses/MIT)
[![Open In Colab](https://colab.research.google.com/assets/colab-badge.svg)](https://colab.research.google.com/)

A comprehensive, professional-grade Excel-based depreciation calculator designed specifically for **Indian Chartered Accountants** and **Tax Professionals**. Calculates depreciation under both:

- 🏢 **Companies Act, 2013** (Schedule II - Straight Line Method)
- 💰 **Income Tax Act, 1961** (Written Down Value Method)

## 🎯 Perfect For

- ✅ Chartered Accountants preparing depreciation schedules
- ✅ Tax consultants calculating WDV for Income Tax returns
- ✅ Finance teams doing book-tax depreciation reconciliation
- ✅ Audit firms verifying depreciation calculations
- ✅ Business owners maintaining fixed asset registers
- ✅ CA students learning depreciation concepts

## 🚀 Key Features

### 📊 **Dual Compliance Calculations**
- **Companies Act, 2013**: SLM with Schedule II useful lives and pro-rata rules
- **Income Tax Act, 1961**: WDV method with block concept and additional depreciation

### ⚡ **Advanced Automation**
- 🔄 **Automatic Pro-rata**: Detects assets used <180 days and applies half-year depreciation
- 📦 **Block-wise Processing**: Groups assets by depreciation rates for Income Tax compliance
- 📅 **Smart Date Calculations**: Automatically calculates days used in financial year
- 🔢 **Dynamic Formulas**: All calculations use Excel formulas (not static values)

### 🎨 **Professional Excel Features**
- 🎯 **Data Validation**: Dropdown menus for asset categories and blocks
- 💰 **Indian Rupee Formatting**: Proper currency display with ₹ symbol
- 📏 **Auto Column Width**: Optimized for readability
- 🖼️ **Professional Styling**: Headers, borders, and color schemes
- 📈 **Summary Dashboard**: Variance analysis between book and tax depreciation

### 🛡️ **Audit & Compliance Ready**
- 📋 **Clear Methodology**: Transparent calculation logic
- 🔍 **Traceable Formulas**: Easy to audit and verify
- 📄 **Documentation**: Comprehensive instructions and methodology
- 💾 **Backup Safe**: Excel formulas ensure data integrity

## 📋 Input Structure

The calculator accepts the following asset details:

| Field | Description | Example |
|-------|-------------|---------|
| **Asset Name** | Description of asset | "Office Computer" |
| **Asset Category** | Type classification | "Plant & Machinery" |
| **Date of Purchase** | Acquisition date | "01/04/2023" |
| **Date Put to Use** | Commissioning date | "01/04/2023" |
| **Cost of Asset** | Purchase price | ₹50,000 |
| **Residual Value** | Scrap value | ₹5,000 |
| **Useful Life** | Schedule II life | 5 years |
| **Depreciation Rate** | Income Tax rate | 15.00% |
| **Block of Asset** | WDV block | "Block A" |
| **Opening WDV** | Previous year WDV | ₹0 |
| **Additions** | New purchases | ₹0 |
| **Deletions** | Disposals | ₹0 |

## 📊 Output Sheets

### 1. **Input Sheet**
- Asset data entry with validation
- Dropdown menus for categories and blocks
- Sample data pre-loaded for reference

### 2. **Companies Act Depreciation**
- SLM calculations with pro-rata logic
- Automatic days-used calculations
- Accumulated depreciation tracking
- Closing WDV computation

### 3. **Income Tax Depreciation**
- Block-wise WDV calculations
- Additional depreciation for <180 days usage
- Normal vs additional depreciation split
- Closing WDV by block

### 4. **Summary**
- Total depreciation comparison
- Variance analysis (Book vs Tax)
- Compliance reporting metrics

## 🧮 Calculation Methodology

### Companies Act, 2013 (SLM)
```
Annual Depreciation = (Cost - Residual Value) ÷ Useful Life
Pro-rata Depreciation = Annual Depreciation × (Days Used ÷ 365)
Condition: If Days Used < 180 → Apply pro-rata
           If Days Used ≥ 180 → Apply full annual depreciation
```

### Income Tax Act, 1961 (WDV)
```
Block WDV = Opening WDV + Additions - Deletions
Normal Depreciation = Block WDV × Depreciation Rate
Additional Depreciation = Normal Depreciation × 50% (if any asset used <180 days)
Total Depreciation = Normal + Additional
Closing WDV = Block WDV - Total Depreciation
```

## 🛠️ Installation & Usage

### Prerequisites
```bash
Python 3.8 or higher
pip package manager
```

### Installation
```bash
# Clone the repository
git clone https://github.com/yourusername/depreciation-calculator.git
cd depreciation-calculator

# Install dependencies
pip install -r requirements.txt
```

### Generate Calculator
```bash
# Run the calculator generator
python depreciation_calculator_pro.py
```

This creates `Depreciation_Calculator.xlsx` in the current directory.

### Usage Instructions
1. **Open** the generated Excel file
2. **Go to "Input" sheet** and enter your asset details
3. **Review calculations** in respective sheets
4. **Use "Summary" sheet** for variance analysis
5. **Save regularly** and backup your work

## 📁 Project Structure

```
depreciation-calculator/
├── depreciation_calculator_pro.py  # Main calculator script
├── requirements.txt                 # Python dependencies
├── README.md                       # This documentation
├── LICENSE                         # MIT License
├── .gitignore                      # Git ignore rules
└── Depreciation_Calculator.xlsx    # Generated Excel file
```

## 📚 Dependencies

```txt
pandas>=2.0.0        # Data manipulation
openpyxl>=3.1.0      # Excel file generation
```

## 🎓 Learning Outcomes

This project demonstrates:
- **Financial Modeling**: Complex depreciation calculations
- **Excel Automation**: Professional spreadsheet generation
- **Python Programming**: Object-oriented design and Excel integration
- **Tax Compliance**: Indian corporate tax law implementation
- **Professional Documentation**: Comprehensive README and code comments

## 🤝 Contributing

Contributions welcome! Areas for improvement:
- Additional depreciation methods
- Multi-year depreciation projections
- PDF report generation
- Web-based interface
- Database integration

## 📄 License

This project is licensed under the MIT License - see the [LICENSE](LICENSE) file for details.

## 👨‍💼 Author

**Chartered Accountant & Python Developer**

*Portfolio Project for GitHub*

---

## 📞 Support

For questions or issues:
- Open a GitHub issue
- Check the documentation
- Review sample calculations

---

**⚠️ Disclaimer**: This tool is for educational and professional use. Always verify calculations with official tax authorities and consult qualified professionals for tax advice.