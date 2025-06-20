# Installation Guide - EnergyGauge USA Automation

## System Requirements

- **Operating System**: Windows 10/11 (required for EnergyGauge COM interface)
- **Python**: 3.8 or later
- **RAM**: Minimum 4GB (8GB recommended for OCR processing)
- **Storage**: 2GB free space for archives and outputs

## Required Software

### 1. EnergyGauge USA
- Install EnergyGauge USA from the official source
- Ensure you have a valid license
- Test that the software runs manually before automation

### 2. Python 3.8+
- Download from https://python.org
- During installation, check "Add Python to PATH"
- Verify installation: `python --version`

### 3. Tesseract OCR Engine
- Download from: https://github.com/UB-Mannheim/tesseract/wiki
- Install to default location: `C:\Program Files\Tesseract-OCR\`
- Add to system PATH or update script configuration

## Quick Setup (Windows)

1. **Download the automation project** to your desired location
2. **Run the setup script**:
   ```cmd
   setup.bat
   ```
3. **Place your EnergyGauge template** in `templates\YourTemplate.egpj`
4. **Add input JSON files** to the `inputs\` directory
5. **Run the automation**:
   ```cmd
   run_automation.bat
   ```

## Manual Setup

### Step 1: Create Virtual Environment
```cmd
python -m venv energy_gauge_env
energy_gauge_env\Scripts\activate
```

### Step 2: Install Dependencies
```cmd
pip install -r requirements.txt
```

### Step 3: Configure Tesseract (if needed)
If Tesseract is not in your PATH, edit the script configuration:
```python
'tesseract_path': r'C:\Your\Path\To\tesseract.exe'
```

### Step 4: Set Up Template
1. Create or obtain an EnergyGauge project template
2. Save it as `templates\YourTemplate.egpj`
3. This template should contain your standard defaults

## Verification

### Test Basic Functionality
```cmd
python code\energy_gauge_automation.py --create-sample
```
This creates a sample input file for testing.

### Test OCR (Optional)
If you have floor plan images:
1. Place test images in your project directory
2. Update the `floor_plan_image` path in sample JSON
3. Run a test to verify OCR extraction

### Test COM Interface
The first run will test the EnergyGauge COM connection. If it fails:
1. Ensure EnergyGauge is properly installed
2. Try running the script as Administrator
3. Check Windows Event Viewer for COM errors

## Configuration Options

Edit these settings in `energy_gauge_automation.py`:

```python
config = {
    'ocr_confidence_threshold': 0.6,        # OCR confidence threshold
    'duplicate_detection_enabled': True,     # Enable duplicate detection
    'tesseract_path': r'C:\Program Files\Tesseract-OCR\tesseract.exe',
    'energygauge_com_id': 'EnergyGauge.Application',  # COM interface ID
    'template_file': 'YourTemplate.egpj',   # Template filename
    'max_retries': 3,                       # Maximum retry attempts
    'processing_timeout': 300               # Timeout in seconds
}
```

## Directory Structure After Setup

```
energy_gauge_automation/
├── code/
│   └── energy_gauge_automation.py
├── templates/
│   └── YourTemplate.egpj          # Your template file
├── inputs/
│   ├── Lot101_inputs.json         # Input files you create
│   └── Lot102_inputs.json
├── outputs/                       # Generated automatically
├── archive/                       # Generated automatically
├── energy_gauge_env/              # Virtual environment
├── requirements.txt
├── setup.bat
├── run_automation.bat
└── README.md
```

## Troubleshooting Installation

### Python Issues
- **"python is not recognized"**: Python not in PATH, reinstall with PATH option
- **Permission errors**: Run command prompt as Administrator
- **Virtual environment fails**: Ensure you have write permissions

### Package Installation Issues
- **pywin32 fails**: This is Windows-specific, install manually: `pip install pywin32`
- **easyocr fails**: Large download, ensure stable internet connection
- **opencv fails**: Try: `pip install opencv-python-headless`

### EnergyGauge COM Issues
- **COM connection fails**: 
  - Verify EnergyGauge installation
  - Run script as Administrator
  - Check if EnergyGauge is already running
- **Template not found**: Ensure template file exists and path is correct

### OCR Issues
- **Tesseract not found**: Install Tesseract and add to PATH
- **Poor OCR results**: Check image quality and preprocessing settings
- **OCR crashes**: Verify image file formats (PNG, JPG supported)

## Getting Help

1. **Check the logs**: `archive\automation.log`
2. **Run in batch mode** for detailed console output
3. **Test with sample data** before using real data
4. **Verify all prerequisites** are properly installed

## Security Considerations

- The script requires COM access to EnergyGauge
- Ensure input files are from trusted sources
- Archive directory may contain sensitive project data
- Consider firewall rules if running on networked systems

## Next Steps

After successful installation:
1. Read the main README.md for usage instructions
2. Create your input JSON files following the schema
3. Test with a small batch before processing large datasets
4. Set up regular archiving of processed data