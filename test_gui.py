#!/usr/bin/env python3
"""
Test script for EnergyGauge Automation GUI
This script tests the modern GUI without requiring EnergyGauge to be installed.
"""

import os
import sys
import json
from pathlib import Path

# Add the code directory to Python path
project_root = Path(__file__).parent
code_dir = project_root / "code"
sys.path.insert(0, str(code_dir))

def create_test_data():
    """Create test data files for GUI testing"""
    print("Creating test data...")
    
    # Ensure directories exist
    (project_root / "inputs").mkdir(exist_ok=True)
    (project_root / "templates").mkdir(exist_ok=True)
    (project_root / "outputs").mkdir(exist_ok=True)
    
    # Create sample JSON files
    sample_lots = [
        {
            "lot_id": "Lot301",
            "project_info": {
                "name": "Test Project 1", 
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
            }
        },
        {
            "lot_id": "Lot302", 
            "project_info": {
                "name": "Test Project 2",
                "address": "456 Oak Ave",
                "city": "Tampa",
                "state": "FL", 
                "zip": "33601"
            },
            "building_data": {
                "conditioned_floor_area": 1850.0,
                "windows": {
                    "N": {"area": 85.0, "u_factor": 0.30, "shgc": 0.20},
                    "S": {"area": 110.0, "u_factor": 0.30, "shgc": 0.20}
                },
                "walls": {
                    "WoodFrameExt": {"area": 1200.0, "r_value": 19.0}
                },
                "infiltration": {"ach50": 6.5}
            },
            "hvac": {
                "system1": {"tonnage": 2.5, "seer2": 16.0}
            },
            "duct": {
                "location": "Attic"
            }
        }
    ]
    
    # Save sample files
    for lot_data in sample_lots:
        lot_id = lot_data["lot_id"]
        filename = f"{lot_id}_inputs.json"
        filepath = project_root / "inputs" / filename
        
        with open(filepath, 'w') as f:
            json.dump(lot_data, f, indent=2)
        
        print(f"Created: {filepath}")
    
    # Create a mock template file
    template_file = project_root / "templates" / "TestTemplate.egpj"
    with open(template_file, 'w') as f:
        f.write("# Mock EnergyGauge Template File\n")
        f.write("# This is for testing purposes only\n")
    
    print(f"Created: {template_file}")
    print("Test data creation complete!")

def test_gui():
    """Test the modern GUI"""
    print("Starting EnergyGauge Automation GUI test...")
    
    try:
        # Import the main modules
        from energy_gauge_automation import EnergyGaugeAutomation, ModernEnergyGaugeGUI
        
        print("✅ Successfully imported automation modules")
        
        # Initialize automation (should work even without EnergyGauge installed)
        automation = EnergyGaugeAutomation(project_root)
        print("✅ Successfully initialized automation system")
        
        # Launch the modern GUI
        print("🚀 Launching Modern GUI...")
        print("\nGUI Features to test:")
        print("1. Upload JSON files from the 'inputs' directory")
        print("2. Set template file to 'templates/TestTemplate.egpj'")
        print("3. Test the approval workflow")
        print("4. Try COM interface discovery (will show error without EnergyGauge)")
        print("5. Check results and reporting features")
        
        gui = ModernEnergyGaugeGUI(automation)
        gui.run()
        
    except ImportError as e:
        print(f"❌ Import error: {e}")
        print("Make sure all required packages are installed:")
        print("pip install pywin32 opencv-python pytesseract easyocr watchdog pandas pillow")
    except Exception as e:
        print(f"❌ Error: {e}")
        import traceback
        traceback.print_exc()

if __name__ == "__main__":
    print("=== EnergyGauge Automation GUI Test ===")
    print()
    
    # Create test data first
    create_test_data()
    print()
    
    # Test the GUI
    test_gui()