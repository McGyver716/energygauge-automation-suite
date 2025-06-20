# EnergyGauge USA Automation Suite

**Modern automation system for EnergyGauge USA with intuitive GUI, schematic upload, approval workflow, and Windows 10 compatibility.**

This project provides a comprehensive automation solution for EnergyGauge USA software with:

## ✨ Key Features

### 🖥️ Modern GUI Interface
- **Tabbed Interface**: Organized workflow with Project Setup, Processing & Approval, Results, and Settings tabs
- **Drag & Drop Upload**: Easy schematic and JSON data file upload
- **Real-time Preview**: Schematic image preview and data parameter display
- **Approval Workflow**: Human-in-the-loop approval with Approve/Reject/Skip options
- **Progress Tracking**: Visual progress bars and processing queue status

### 🔧 Advanced Automation
- **COM Interface Discovery**: Automatically discover EnergyGauge COM properties and methods
- **Parallel Processing**: Process multiple lots simultaneously for faster throughput
- **Multi-Layer OCR**: Tesseract and EasyOCR for extracting data from floor plans  
- **Error Recovery**: Comprehensive error handling with retry mechanisms
- **Template Management**: Support for EnergyGauge template files (.egpj)

### 📊 Quality Control & Monitoring
- **Stoplight Indicators**: Green/Yellow/Red status for data quality and system health
- **Processing History**: Complete audit trail of all processing activities
- **Reporting**: Automated report generation with processing statistics
- **Archive Management**: Automatic archiving of inputs, outputs, and processing logs

### 🚀 Windows 10 Ready
- **One-Click Launch**: Easy Windows batch files for instant startup
- **Virtual Environment**: Automatic Python environment setup and package installation
- **File Association**: Direct integration with Windows file explorer
- **Administrator Support**: Enhanced COM interface access when needed

## Project Structure

```
energy_gauge_automation/
├── code/
│   └── energy_gauge_automation.py    # Main automation script
├── templates/
│   └── YourTemplate.egpj            # EnergyGauge template file
├── inputs/
│   ├── Lot101_inputs.json           # Input data files
│   ├── Lot102_inputs.json
│   └── ...
├── outputs/
│   ├── Lot101/                      # Output files for each lot
│   ├── Lot102/
│   └── ...
└── archive/
    └── ...                          # Archived inputs/outputs for regression
```

## Installation

### Prerequisites

1. **Windows 10/11** with Python 3.8+
2. **EnergyGauge USA** installed and licensed
3. **Tesseract OCR** installed and in PATH

### Setup

1. Create a virtual environment (recommended):
```bash
python -m venv energy_gauge_env
energy_gauge_env\Scripts\activate
```

2. Install required packages:
```bash
pip install pywin32 opencv-python pytesseract easyocr watchdog pandas
```

3. Install Tesseract OCR:
   - Download from: https://github.com/UB-Mannheim/tesseract/wiki
   - Add to PATH or update `tesseract_path` in config

## 🚀 Quick Start

### Option 1: One-Click Windows Launch (Recommended)
Double-click `run_modern_gui.bat` for instant setup and launch:
- Automatically creates Python virtual environment
- Installs all required packages
- Creates sample data files
- Launches the modern GUI interface

### Option 2: Manual Launch
```bash
# Modern GUI with full features
python code/energy_gauge_automation.py --mode modern

# Traditional simple GUI  
python code/energy_gauge_automation.py --mode gui

# Batch processing (no GUI)
python code/energy_gauge_automation.py --mode batch

# Discover EnergyGauge COM interface
python code/energy_gauge_automation.py --discover

# Create sample input files
python code/energy_gauge_automation.py --create-sample
```

### Option 3: COM Interface Discovery
Run `discover_energygauge.bat` to:
- Connect to your EnergyGauge USA installation
- Discover available COM properties and methods
- Generate customization guidance for your specific version

## Input Data Format

Each lot requires a JSON file in the `inputs/` directory:

```json
{
  "lot_id": "Lot101",
  "project_info": {
    "name": "Sample Project",
    "address": "123 Main St",
    "city": "Orlando", 
    "state": "FL",
    "zip": "32801"
  },
  "building_data": {
    "conditioned_floor_area": 2402.0,
    "windows": {
      "NE": {"area": 120.5, "u_factor": 0.32, "shgc": 0.25},
      "SW": {"area": 98.3, "u_factor": 0.32, "shgc": 0.25}
    },
    "walls": {
      "WoodFrameExt": {"area": 1650.0, "r_value": 19.0}
    },
    "infiltration": {"ach50": 7.0}
  },
  "hvac": {
    "system1": {"tonnage": 3.0, "seer2": 15.0}
  },
  "duct": {
    "location": "Interior"
  },
  "floor_plan_image": "floorplans/Lot101_plan.png"
}
```

## Quality Control Indicators

### Data Quality
- **Green**: All required data present and valid
- **Yellow**: Some data missing/defaulted or low OCR confidence
- **Red**: Critical data missing or OCR failed

### System Health  
- **Green**: All operations successful
- **Yellow**: Minor recoverable issues (retries, warnings)
- **Red**: Critical errors requiring attention

## OCR Processing

The system uses a two-tier OCR approach:

1. **Primary**: Tesseract OCR with image preprocessing
2. **Fallback**: EasyOCR for text Tesseract cannot read

OCR is used to extract missing data from floor plan images, including:
- Conditioned floor area
- HVAC tonnage
- Window specifications
- Other building parameters

## Duplicate Detection

Prevents redundant processing by:
- Computing MD5 hashes of floor plan images
- Comparing JSON content (excluding unique identifiers)
- Skipping processing if duplicate detected
- Reusing previous results for efficiency

## Error Handling

- All operations wrapped in try/catch blocks
- Detailed logging to `archive/automation.log`
- Processing continues even if individual lots fail
- System health indicators reflect error states
- Archive maintains processing history

## Archive System

Automatically archives:
- Input JSON files with timestamps
- Floor plan images  
- Output project files and reports
- Processing logs with success/failure status
- Support for regression testing

## Configuration

Key configuration options in the script:

```python
config = {
    'ocr_confidence_threshold': 0.6,
    'duplicate_detection_enabled': True,
    'tesseract_path': r'C:\Program Files\Tesseract-OCR\tesseract.exe',
    'energygauge_com_id': 'EnergyGauge.Application',
    'template_file': 'YourTemplate.egpj',
    'max_retries': 3,
    'processing_timeout': 300
}
```

## Troubleshooting

### Common Issues

1. **COM Connection Failed**
   - Ensure EnergyGauge USA is installed and licensed
   - Try running as Administrator
   - Check Windows COM registration

2. **OCR Not Working**
   - Verify Tesseract installation and PATH
   - Check image file paths and formats
   - Review OCR confidence thresholds

3. **Missing Dependencies**
   - Install packages: `pip install pywin32 opencv-python pytesseract easyocr`
   - For Linux testing: dependencies will auto-install (Windows COM features disabled)

### Logs

Check these files for debugging:
- `archive/automation.log` - Detailed processing log
- `archive/processing_log.csv` - Summary of all runs
- Console output - Real-time status updates

## Future Enhancements

The modular design supports future additions:
- Machine learning agents for result optimization
- Web dashboard interface
- Integration with external databases
- Additional OCR engines
- Cloud processing capabilities
- Integration with EnergyGauge Commercial

## Support

For issues or questions:
1. Check the processing log files
2. Review configuration settings
3. Verify all prerequisites are installed
4. Test with sample data first

## License

This automation script is provided as-is for educational and automation purposes. Ensure compliance with EnergyGauge USA licensing terms.