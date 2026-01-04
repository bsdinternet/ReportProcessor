# Report Processing Suite 📊

[![Python](https://img.shields.io/badge/Python-3.8%2B-blue.svg)](https://www.python.org/downloads/)
[![License](https://img.shields.io/badge/License-MIT-green.svg)](LICENSE)
[![CustomTkinter](https://img.shields.io/badge/GUI-CustomTkinter-orange.svg)](https://github.com/TomSchimansky/CustomTkinter)

A comprehensive desktop application for automating report processing workflows across multiple e-commerce platforms including Meesho, Flipkart, and Amazon.

<p align="center">
  <img src="docs/Images/Screenshot 2026-01-04 233718.png" alt="Report Processing Suite - Main Dashboard" width="800">
  <br>
  <i>Modern desktop application for e-commerce report automation</i>
</p>

## 🌟 Features

### 📦 Pickup Report Processor
- Process pickup reports from multiple courier sources
- Support for Sellerflex, Flipkart KC/LL, and Meesho manifests
- Automatic pivot table generation with visual analytics
- PDF manifest parsing and data extraction
- Real-time progress tracking

### ↩️ Returns Reconciliation
- Reconcile return shipments across platforms
- Generate comprehensive return reports
- Track return status and analyze patterns
- Support for Meesho, Flipkart KC/LL, and SellerFlex
- Automated data consolidation

### ❌ Cancellation Report
- Generate detailed cancellation reports
- Analyze cancellation patterns and trends
- Multi-platform support (Meesho, Flipkart)
- Filter by cancellation type and date
- Export to Excel with formatting

## 🚀 Getting Started

### Prerequisites

```bash
Python 3.8 or higher
pip (Python package installer)
```

### Installation

1. **Clone the repository**
```bash
git clone https://github.com/yourusername/report-processing-suite.git
cd report-processing-suite
```

2. **Install required packages**
```bash
pip install -r requirements.txt
```

3. **Run the application**
```bash
python Homepage.py
```

## 📁 Project Structure

```
report-processing-suite/
├── Homepage.py                 # Main application entry point
├── Cancellationexe.py         # Cancellation report module
├── Returnsreportexe.py        # Returns reconciliation module
├── Pickupreportexe.py         # Pickup report module
├── InputDIR/                  # Input files directory
│   ├── CancellationReport/
│   ├── PickupReportfiles/
│   └── Returnsreportfiles/
├── Template/                  # Excel templates
│   └── ReturnsReconcileReport.xlsx
├── Output/                    # Generated reports (auto-created)
│   └── Pivot_PNGs/           # Pivot table images
├── requirements.txt           # Python dependencies
├── README.md                  # This file
└── LICENSE                    # License file
```

## 📋 Requirements

Create a `requirements.txt` file with the following dependencies:

```
pandas>=1.5.0
openpyxl>=3.1.0
pdfplumber>=0.9.0
customtkinter>=5.2.0
matplotlib>=3.7.0
Pillow>=10.0.0
pypdfium2>=4.0.0
```

## 🎯 Usage

### Pickup Report Processing

1. Place input files in `InputDIR/PickupReportfiles/`:
   - `Sellerflex.csv`
   - `Flipkart KC.csv`
   - `Flipkart LL.csv`
   - `Manifest.pdf`

2. Click "Open Module" for Pickup Report Processor
3. Click "Start Processing"
4. Reports will be saved in `Output/`

### Returns Reconciliation

1. Place input files in `InputDIR/Returnsreportfiles/`:
   - `Returns Meesho.csv`
   - `Returns Flipkart KC.csv`
   - `Returns Flipkart LL.csv`
   - `Returns SellerFlex.csv`

2. Open Returns Reconciliation module
3. Click "Start Returns Processing"
4. View generated report in `Output/`

### Cancellation Report

1. Place required files in appropriate directories:
   - Cancellation CSVs in `InputDIR/CancellationReport/`
   - Pickup files in `InputDIR/PickupReportfiles/`

2. Open Cancellation Report module
3. Click "Process Reports"
4. Save the generated Excel report

## 🔧 Configuration

### Template Files
Place your Excel templates in the `Template/` directory. The application will use these as base templates for report generation.

### Customization
Edit the color scheme and UI settings in each module file:
```python
COLORS = {
    "primary_orange": "#FF6B35",
    "secondary_orange": "#FF8F65",
    # ... customize as needed
}
```

## 🏗️ Building Executables

To create standalone executables using PyInstaller:

```bash
# For the main application
pyinstaller Homepage.spec

# Or use the command line
pyinstaller --onedir --windowed --name "ReportProcessor" Homepage.py
```

## 🤝 Contributing

Contributions are welcome! Please follow these steps:

1. Fork the repository
2. Create a feature branch (`git checkout -b feature/AmazingFeature`)
3. Commit your changes (`git commit -m 'Add some AmazingFeature'`)
4. Push to the branch (`git push origin feature/AmazingFeature`)
5. Open a Pull Request

## 📝 License

This project is licensed under the MIT License - see the [LICENSE](LICENSE) file for details.

## 🐛 Known Issues

- PDF processing requires specific PDF format for Meesho manifests
- Large CSV files (>100MB) may require extended processing time
- Excel templates must maintain specific column structures

## 📧 Contact

Deepak Kumar BS - deepakkumarbscsa2022@gmail.com

Project Link: https://github.com/bsdinternet/ReportProcessor.git

## 🙏 Acknowledgments

- [CustomTkinter](https://github.com/TomSchimansky/CustomTkinter) - Modern UI library
- [pdfplumber](https://github.com/jsvine/pdfplumber) - PDF processing
- [pandas](https://pandas.pydata.org/) - Data manipulation
- [openpyxl](https://openpyxl.readthedocs.io/) - Excel file handling

---

**Note:** This application is designed for internal business use. Ensure compliance with data privacy regulations when processing customer information.
