#!/usr/bin/env python3
"""
EnergyGauge USA Automation Script
Automates EnergyGauge USA software via COM interface with multi-layer OCR,
quality control GUI, and comprehensive error handling.
"""

import os
import sys
import json
import logging
import hashlib
import shutil
import traceback
from datetime import datetime
from pathlib import Path
from typing import Dict, List, Optional, Tuple, Any
import threading
import queue

# Auto-install required packages
def install_package(package):
    """Install a package using pip"""
    import subprocess
    try:
        subprocess.check_call([sys.executable, "-m", "pip", "install", package])
        return True
    except subprocess.CalledProcessError:
        return False

# Import required libraries with auto-install fallback
try:
    import win32com.client
except ImportError:
    print("Installing pywin32...")
    if install_package("pywin32"):
        import win32com.client
    else:
        print("ERROR: Could not install pywin32. Please install manually.")
        sys.exit(1)

try:
    import cv2
except ImportError:
    print("Installing opencv-python...")
    if install_package("opencv-python"):
        import cv2
    else:
        print("WARNING: OpenCV not available. Image preprocessing disabled.")
        cv2 = None

try:
    import pytesseract
except ImportError:
    print("Installing pytesseract...")
    if install_package("pytesseract"):
        import pytesseract
    else:
        print("WARNING: Tesseract OCR not available.")
        pytesseract = None

try:
    import easyocr
except ImportError:
    print("Installing easyocr...")
    if install_package("easyocr"):
        import easyocr
    else:
        print("WARNING: EasyOCR not available.")
        easyocr = None

try:
    from watchdog.observers import Observer
    from watchdog.events import FileSystemEventHandler
except ImportError:
    print("Installing watchdog...")
    if install_package("watchdog"):
        from watchdog.observers import Observer
        from watchdog.events import FileSystemEventHandler
    else:
        print("WARNING: Watchdog not available. File monitoring disabled.")
        Observer = None
        FileSystemEventHandler = None

import tkinter as tk
from tkinter import ttk, scrolledtext
import pandas as pd
from concurrent.futures import ThreadPoolExecutor
import argparse
import time

# Add GUI imports
try:
    from tkinter import filedialog, messagebox
    from PIL import Image, ImageTk
except ImportError:
    print("Installing PIL/Pillow for image handling...")
    if install_package("Pillow"):
        from PIL import Image, ImageTk
    else:
        print("WARNING: PIL not available. Image preview disabled.")
        Image = None
        ImageTk = None


class ConfigManager:
    """Manages configuration settings for the automation"""
    
    def __init__(self):
        self.config = {
            'ocr_confidence_threshold': 0.6,
            'duplicate_detection_enabled': True,
            'tesseract_path': r'C:\Program Files\Tesseract-OCR\tesseract.exe',
            'energygauge_com_id': 'EnergyGauge.Application',
            'template_file': 'YourTemplate.egpj',
            'log_level': 'INFO',
            'max_retries': 3,
            'processing_timeout': 300  # seconds
        }
    
    def get(self, key: str, default=None):
        return self.config.get(key, default)
    
    def set(self, key: str, value):
        self.config[key] = value


class QualityIndicator:
    """Manages quality control indicators"""
    
    def __init__(self):
        self.data_quality = "GREEN"  # GREEN, YELLOW, RED
        self.system_health = "GREEN"
        self.current_lot = ""
        self.status_message = ""
        self.processed_lots = []
        
    def set_data_quality(self, status: str, message: str = ""):
        self.data_quality = status.upper()
        if message:
            self.status_message = message
            
    def set_system_health(self, status: str, message: str = ""):
        self.system_health = status.upper()
        if message:
            self.status_message = message
            
    def set_current_lot(self, lot_id: str):
        self.current_lot = lot_id
        
    def add_processed_lot(self, lot_id: str, status: str):
        self.processed_lots.append({'lot_id': lot_id, 'status': status, 'timestamp': datetime.now()})


class StoplightGUI:
    """Stoplight-style GUI for quality control monitoring"""
    
    def __init__(self, quality_indicator: QualityIndicator):
        self.quality_indicator = quality_indicator
        self.root = tk.Tk()
        self.root.title("EnergyGauge Automation - Quality Control")
        self.root.geometry("600x500")
        self.root.resizable(True, True)
        
        self.setup_gui()
        self.update_indicators()
        
    def setup_gui(self):
        """Setup the GUI layout"""
        # Main frame
        main_frame = ttk.Frame(self.root, padding="10")
        main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        # Configure grid weights
        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(0, weight=1)
        main_frame.columnconfigure(1, weight=1)
        
        # Title
        title_label = ttk.Label(main_frame, text="EnergyGauge USA Automation", 
                               font=("Arial", 16, "bold"))
        title_label.grid(row=0, column=0, columnspan=3, pady=(0, 20))
        
        # Indicators frame
        indicators_frame = ttk.LabelFrame(main_frame, text="Status Indicators", padding="10")
        indicators_frame.grid(row=1, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=(0, 10))
        indicators_frame.columnconfigure(1, weight=1)
        
        # Data Quality Indicator
        ttk.Label(indicators_frame, text="Data Quality:", 
                 font=("Arial", 12, "bold")).grid(row=0, column=0, sticky=tk.W, padx=(0, 10))
        
        self.data_quality_canvas = tk.Canvas(indicators_frame, width=30, height=30)
        self.data_quality_canvas.grid(row=0, column=1, sticky=tk.W, padx=(0, 10))
        self.data_quality_circle = self.data_quality_canvas.create_oval(5, 5, 25, 25, 
                                                                       fill="green", outline="black")
        
        self.data_quality_label = ttk.Label(indicators_frame, text="GREEN", 
                                           font=("Arial", 10, "bold"))
        self.data_quality_label.grid(row=0, column=2, sticky=tk.W)
        
        # System Health Indicator
        ttk.Label(indicators_frame, text="System Health:", 
                 font=("Arial", 12, "bold")).grid(row=1, column=0, sticky=tk.W, padx=(0, 10))
        
        self.system_health_canvas = tk.Canvas(indicators_frame, width=30, height=30)
        self.system_health_canvas.grid(row=1, column=1, sticky=tk.W, padx=(0, 10))
        self.system_health_circle = self.system_health_canvas.create_oval(5, 5, 25, 25, 
                                                                         fill="green", outline="black")
        
        self.system_health_label = ttk.Label(indicators_frame, text="GREEN", 
                                            font=("Arial", 10, "bold"))
        self.system_health_label.grid(row=1, column=2, sticky=tk.W)
        
        # Current Status
        status_frame = ttk.LabelFrame(main_frame, text="Current Status", padding="10")
        status_frame.grid(row=2, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=(0, 10))
        status_frame.columnconfigure(1, weight=1)
        
        ttk.Label(status_frame, text="Current Lot:").grid(row=0, column=0, sticky=tk.W)
        self.current_lot_label = ttk.Label(status_frame, text="None", font=("Arial", 10, "bold"))
        self.current_lot_label.grid(row=0, column=1, sticky=tk.W, padx=(10, 0))
        
        ttk.Label(status_frame, text="Status:").grid(row=1, column=0, sticky=tk.W)
        self.status_message_label = ttk.Label(status_frame, text="Ready")
        self.status_message_label.grid(row=1, column=1, sticky=tk.W, padx=(10, 0))
        
        # Log display
        log_frame = ttk.LabelFrame(main_frame, text="Processing Log", padding="10")
        log_frame.grid(row=3, column=0, columnspan=3, sticky=(tk.W, tk.E, tk.N, tk.S), pady=(0, 10))
        log_frame.columnconfigure(0, weight=1)
        log_frame.rowconfigure(0, weight=1)
        main_frame.rowconfigure(3, weight=1)
        
        self.log_text = scrolledtext.ScrolledText(log_frame, height=10, width=70)
        self.log_text.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        # Control buttons
        button_frame = ttk.Frame(main_frame)
        button_frame.grid(row=4, column=0, columnspan=3, pady=(10, 0))
        
        self.start_button = ttk.Button(button_frame, text="Start Processing", 
                                      command=self.start_processing)
        self.start_button.pack(side=tk.LEFT, padx=(0, 10))
        
        self.stop_button = ttk.Button(button_frame, text="Stop", 
                                     command=self.stop_processing, state=tk.DISABLED)
        self.stop_button.pack(side=tk.LEFT, padx=(0, 10))
        
        ttk.Button(button_frame, text="Clear Log", 
                  command=self.clear_log).pack(side=tk.LEFT)
        
        # Processing control
        self.processing = False
        self.processing_thread = None
        
    def update_indicators(self):
        """Update the visual indicators based on current status"""
        # Update data quality indicator
        color_map = {"GREEN": "green", "YELLOW": "yellow", "RED": "red"}
        
        data_color = color_map.get(self.quality_indicator.data_quality, "gray")
        self.data_quality_canvas.itemconfig(self.data_quality_circle, fill=data_color)
        self.data_quality_label.config(text=self.quality_indicator.data_quality)
        
        health_color = color_map.get(self.quality_indicator.system_health, "gray")
        self.system_health_canvas.itemconfig(self.system_health_circle, fill=health_color)
        self.system_health_label.config(text=self.quality_indicator.system_health)
        
        # Update current status
        self.current_lot_label.config(text=self.quality_indicator.current_lot or "None")
        self.status_message_label.config(text=self.quality_indicator.status_message or "Ready")
        
        # Schedule next update
        self.root.after(1000, self.update_indicators)
        
    def log_message(self, message: str):
        """Add a message to the log display"""
        timestamp = datetime.now().strftime("%H:%M:%S")
        self.log_text.insert(tk.END, f"[{timestamp}] {message}\n")
        self.log_text.see(tk.END)
        
    def start_processing(self):
        """Start the processing in a separate thread"""
        if not self.processing:
            self.processing = True
            self.start_button.config(state=tk.DISABLED)
            self.stop_button.config(state=tk.NORMAL)
            
            # Start processing thread
            self.processing_thread = threading.Thread(target=self.run_automation)
            self.processing_thread.daemon = True
            self.processing_thread.start()
            
    def stop_processing(self):
        """Stop the processing"""
        self.processing = False
        self.start_button.config(state=tk.NORMAL)
        self.stop_button.config(state=tk.DISABLED)
        
    def clear_log(self):
        """Clear the log display"""
        self.log_text.delete(1.0, tk.END)
        
    def run_automation(self):
        """Run the automation process"""
        try:
            # This will be called by the main automation class
            if hasattr(self, 'automation_callback'):
                self.automation_callback()
        except Exception as e:
            self.log_message(f"ERROR: {str(e)}")
            self.quality_indicator.set_system_health("RED", f"Error: {str(e)}")
        finally:
            self.processing = False
            self.root.after(0, lambda: self.start_button.config(state=tk.NORMAL))
            self.root.after(0, lambda: self.stop_button.config(state=tk.DISABLED))
            
    def run(self):
        """Start the GUI main loop"""
        self.root.mainloop()


class OCRProcessor:
    """Multi-layer OCR processing system"""
    
    def __init__(self, config: ConfigManager, logger: logging.Logger):
        self.config = config
        self.logger = logger
        self.tesseract_available = pytesseract is not None
        self.easyocr_available = easyocr is not None
        self.opencv_available = cv2 is not None
        
        # Initialize EasyOCR reader if available
        self.easyocr_reader = None
        if self.easyocr_available:
            try:
                self.easyocr_reader = easyocr.Reader(['en'])
                self.logger.info("EasyOCR initialized successfully")
            except Exception as e:
                self.logger.warning(f"EasyOCR initialization failed: {e}")
                self.easyocr_available = False
        
        # Configure Tesseract path if available
        if self.tesseract_available:
            tesseract_path = self.config.get('tesseract_path')
            if os.path.exists(tesseract_path):
                pytesseract.pytesseract.tesseract_cmd = tesseract_path
            else:
                self.logger.warning(f"Tesseract not found at {tesseract_path}")
                
    def preprocess_image(self, image_path: str) -> Optional[str]:
        """Preprocess image for better OCR results"""
        if not self.opencv_available:
            return image_path
            
        try:
            # Read image
            img = cv2.imread(image_path)
            if img is None:
                self.logger.error(f"Could not read image: {image_path}")
                return image_path
                
            # Convert to grayscale
            gray = cv2.cvtColor(img, cv2.COLOR_BGR2GRAY)
            
            # Apply threshold to get binary image
            _, thresh = cv2.threshold(gray, 0, 255, cv2.THRESH_BINARY + cv2.THRESH_OTSU)
            
            # Save preprocessed image
            processed_path = image_path.replace('.', '_processed.')
            cv2.imwrite(processed_path, thresh)
            
            return processed_path
            
        except Exception as e:
            self.logger.error(f"Image preprocessing failed: {e}")
            return image_path
            
    def extract_with_tesseract(self, image_path: str) -> Tuple[str, float]:
        """Extract text using Tesseract OCR"""
        if not self.tesseract_available:
            return "", 0.0
            
        try:
            # Preprocess image
            processed_path = self.preprocess_image(image_path)
            
            # Extract text with confidence
            data = pytesseract.image_to_data(processed_path, output_type=pytesseract.Output.DICT)
            
            # Calculate average confidence
            confidences = [int(conf) for conf in data['conf'] if int(conf) > 0]
            avg_confidence = sum(confidences) / len(confidences) if confidences else 0
            
            # Extract text
            text = pytesseract.image_to_string(processed_path)
            
            # Clean up processed image if different from original
            if processed_path != image_path and os.path.exists(processed_path):
                os.remove(processed_path)
                
            return text.strip(), avg_confidence / 100.0
            
        except Exception as e:
            self.logger.error(f"Tesseract OCR failed: {e}")
            return "", 0.0
            
    def extract_with_easyocr(self, image_path: str) -> Tuple[str, float]:
        """Extract text using EasyOCR"""
        if not self.easyocr_available or not self.easyocr_reader:
            return "", 0.0
            
        try:
            results = self.easyocr_reader.readtext(image_path)
            
            if not results:
                return "", 0.0
                
            # Combine all text and calculate average confidence
            all_text = []
            confidences = []
            
            for (bbox, text, confidence) in results:
                all_text.append(text)
                confidences.append(confidence)
                
            combined_text = ' '.join(all_text)
            avg_confidence = sum(confidences) / len(confidences) if confidences else 0
            
            return combined_text.strip(), avg_confidence
            
        except Exception as e:
            self.logger.error(f"EasyOCR failed: {e}")
            return "", 0.0
            
    def extract_text(self, image_path: str) -> Dict[str, Any]:
        """Extract text using multi-layer OCR approach"""
        if not os.path.exists(image_path):
            self.logger.error(f"Image file not found: {image_path}")
            return {"text": "", "confidence": 0.0, "method": "none", "success": False}
            
        confidence_threshold = self.config.get('ocr_confidence_threshold', 0.6)
        
        # Try Tesseract first
        tesseract_text, tesseract_conf = self.extract_with_tesseract(image_path)
        
        if tesseract_conf >= confidence_threshold and tesseract_text:
            self.logger.info(f"Tesseract OCR successful (confidence: {tesseract_conf:.2f})")
            return {
                "text": tesseract_text,
                "confidence": tesseract_conf,
                "method": "tesseract",
                "success": True
            }
            
        # Fallback to EasyOCR
        self.logger.info("Tesseract confidence low, trying EasyOCR...")
        easyocr_text, easyocr_conf = self.extract_with_easyocr(image_path)
        
        if easyocr_conf >= confidence_threshold and easyocr_text:
            self.logger.info(f"EasyOCR successful (confidence: {easyocr_conf:.2f})")
            return {
                "text": easyocr_text,
                "confidence": easyocr_conf,
                "method": "easyocr",
                "success": True
            }
            
        # Return best available result even if below threshold
        if tesseract_conf >= easyocr_conf:
            method = "tesseract"
            text = tesseract_text
            confidence = tesseract_conf
        else:
            method = "easyocr"
            text = easyocr_text
            confidence = easyocr_conf
            
        self.logger.warning(f"OCR confidence below threshold ({confidence:.2f} < {confidence_threshold})")
        
        return {
            "text": text,
            "confidence": confidence,
            "method": method,
            "success": confidence >= confidence_threshold
        }
        
    def extract_field_from_text(self, text: str, field_name: str) -> Optional[str]:
        """Extract specific field value from OCR text"""
        import re
        
        # Define patterns for common fields
        patterns = {
            'conditioned_floor_area': [
                r'floor\s*area[:\s]*([0-9,]+\.?[0-9]*)',
                r'area[:\s]*([0-9,]+\.?[0-9]*)\s*sq\s*ft',
                r'([0-9,]+\.?[0-9]*)\s*sq\s*ft'
            ],
            'tonnage': [
                r'([0-9]+\.?[0-9]*)\s*ton',
                r'tonnage[:\s]*([0-9]+\.?[0-9]*)'
            ],
            'seer': [
                r'seer[:\s]*([0-9]+\.?[0-9]*)',
                r'([0-9]+\.?[0-9]*)\s*seer'
            ],
            'u_factor': [
                r'u[:\s]*([0-9]+\.?[0-9]*)',
                r'u-factor[:\s]*([0-9]+\.?[0-9]*)'
            ],
            'shgc': [
                r'shgc[:\s]*([0-9]+\.?[0-9]*)',
                r'solar\s*heat\s*gain[:\s]*([0-9]+\.?[0-9]*)'
            ]
        }
        
        field_patterns = patterns.get(field_name.lower(), [])
        
        for pattern in field_patterns:
            match = re.search(pattern, text.lower())
            if match:
                value = match.group(1).replace(',', '')
                try:
                    return str(float(value))
                except ValueError:
                    continue
                    
        return None


class DuplicateDetector:
    """Handles duplicate floor plan detection"""
    
    def __init__(self, config: ConfigManager, logger: logging.Logger):
        self.config = config
        self.logger = logger
        self.processed_hashes = set()
        self.hash_file = Path("archive/processed_hashes.json")
        self.load_processed_hashes()
        
    def load_processed_hashes(self):
        """Load previously processed file hashes"""
        if self.hash_file.exists():
            try:
                with open(self.hash_file, 'r') as f:
                    data = json.load(f)
                    self.processed_hashes = set(data.get('hashes', []))
                    self.logger.info(f"Loaded {len(self.processed_hashes)} processed hashes")
            except Exception as e:
                self.logger.error(f"Failed to load hash file: {e}")
                
    def save_processed_hashes(self):
        """Save processed file hashes"""
        try:
            self.hash_file.parent.mkdir(parents=True, exist_ok=True)
            with open(self.hash_file, 'w') as f:
                json.dump({'hashes': list(self.processed_hashes)}, f, indent=2)
        except Exception as e:
            self.logger.error(f"Failed to save hash file: {e}")
            
    def calculate_hash(self, file_path: str) -> str:
        """Calculate MD5 hash of a file"""
        hash_md5 = hashlib.md5()
        try:
            with open(file_path, "rb") as f:
                for chunk in iter(lambda: f.read(4096), b""):
                    hash_md5.update(chunk)
            return hash_md5.hexdigest()
        except Exception as e:
            self.logger.error(f"Failed to calculate hash for {file_path}: {e}")
            return ""
            
    def is_duplicate(self, input_data: Dict) -> bool:
        """Check if input data represents a duplicate"""
        if not self.config.get('duplicate_detection_enabled', True):
            return False
            
        # Check if floor plan image exists and calculate hash
        floor_plan_path = input_data.get('floor_plan_image')
        if floor_plan_path and os.path.exists(floor_plan_path):
            file_hash = self.calculate_hash(floor_plan_path)
            if file_hash in self.processed_hashes:
                self.logger.info(f"Duplicate detected based on floor plan hash: {file_hash}")
                return True
                
        # Alternative: hash the JSON content (excluding unique identifiers)
        content_for_hash = input_data.copy()
        content_for_hash.pop('lot_id', None)  # Remove unique identifier
        content_for_hash.pop('floor_plan_image', None)  # Remove file path
        
        content_json = json.dumps(content_for_hash, sort_keys=True)
        content_hash = hashlib.md5(content_json.encode()).hexdigest()
        
        if content_hash in self.processed_hashes:
            self.logger.info(f"Duplicate detected based on content hash: {content_hash}")
            return True
            
        return False
        
    def mark_as_processed(self, input_data: Dict):
        """Mark input data as processed"""
        # Mark floor plan hash if available
        floor_plan_path = input_data.get('floor_plan_image')
        if floor_plan_path and os.path.exists(floor_plan_path):
            file_hash = self.calculate_hash(floor_plan_path)
            if file_hash:
                self.processed_hashes.add(file_hash)
                
        # Mark content hash
        content_for_hash = input_data.copy()
        content_for_hash.pop('lot_id', None)
        content_for_hash.pop('floor_plan_image', None)
        
        content_json = json.dumps(content_for_hash, sort_keys=True)
        content_hash = hashlib.md5(content_json.encode()).hexdigest()
        self.processed_hashes.add(content_hash)
        
        self.save_processed_hashes()


class ArchiveManager:
    """Manages archiving of inputs and outputs for regression testing"""
    
    def __init__(self, logger: logging.Logger):
        self.logger = logger
        self.archive_dir = Path("archive")
        self.archive_dir.mkdir(exist_ok=True)
        
        # Initialize processing log
        self.log_file = self.archive_dir / "processing_log.csv"
        if not self.log_file.exists():
            self.create_log_file()
            
    def create_log_file(self):
        """Create the processing log CSV file"""
        try:
            with open(self.log_file, 'w') as f:
                f.write("timestamp,lot_id,status,errors,human_fixes_applied,notes\n")
        except Exception as e:
            self.logger.error(f"Failed to create log file: {e}")
            
    def archive_input(self, lot_id: str, input_data: Dict, floor_plan_path: str = None):
        """Archive input data for a lot"""
        try:
            timestamp = datetime.now().strftime("%Y-%m-%d_%H-%M-%S")
            archive_path = self.archive_dir / f"{lot_id}_{timestamp}_input.json"
            
            with open(archive_path, 'w') as f:
                json.dump(input_data, f, indent=2)
                
            # Archive floor plan if available
            if floor_plan_path and os.path.exists(floor_plan_path):
                floor_plan_ext = Path(floor_plan_path).suffix
                archive_image_path = self.archive_dir / f"{lot_id}_{timestamp}_floorplan{floor_plan_ext}"
                shutil.copy2(floor_plan_path, archive_image_path)
                
            self.logger.info(f"Archived input for {lot_id}")
            
        except Exception as e:
            self.logger.error(f"Failed to archive input for {lot_id}: {e}")
            
    def archive_output(self, lot_id: str, output_path: str):
        """Archive output files for a lot"""
        try:
            timestamp = datetime.now().strftime("%Y-%m-%d_%H-%M-%S")
            
            if os.path.isfile(output_path):
                # Single file
                ext = Path(output_path).suffix
                archive_path = self.archive_dir / f"{lot_id}_{timestamp}_output{ext}"
                shutil.copy2(output_path, archive_path)
            elif os.path.isdir(output_path):
                # Directory - create zip archive
                import zipfile
                archive_path = self.archive_dir / f"{lot_id}_{timestamp}_output.zip"
                with zipfile.ZipFile(archive_path, 'w') as zipf:
                    for root, dirs, files in os.walk(output_path):
                        for file in files:
                            file_path = os.path.join(root, file)
                            arcname = os.path.relpath(file_path, output_path)
                            zipf.write(file_path, arcname)
                            
            self.logger.info(f"Archived output for {lot_id}")
            
        except Exception as e:
            self.logger.error(f"Failed to archive output for {lot_id}: {e}")
            
    def log_processing_result(self, lot_id: str, status: str, errors: str = "", 
                            human_fixes: bool = False, notes: str = ""):
        """Log processing result to CSV"""
        try:
            timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            with open(self.log_file, 'a') as f:
                f.write(f'"{timestamp}","{lot_id}","{status}","{errors}","{human_fixes}","{notes}"\n')
        except Exception as e:
            self.logger.error(f"Failed to log processing result: {e}")


class EnergyGaugeCOMInterface:
    """COM interface for EnergyGauge USA automation"""
    
    def __init__(self, config: ConfigManager, logger: logging.Logger):
        self.config = config
        self.logger = logger
        self.app = None
        self.project = None
        self.connected = False
        
    def connect(self, visible: bool = False) -> bool:
        """Connect to EnergyGauge application via COM"""
        try:
            com_id = self.config.get('energygauge_com_id', 'EnergyGauge.Application')
            self.app = win32com.client.Dispatch(com_id)
            
            if self.app:
                # Make EnergyGauge visible if requested (useful for debugging)
                if visible and hasattr(self.app, 'Visible'):
                    self.app.Visible = True
                
                self.connected = True
                self.logger.info("Successfully connected to EnergyGauge via COM")
                return True
            else:
                self.logger.error("Failed to connect to EnergyGauge - app object is None")
                return False
                
        except Exception as e:
            self.logger.error(f"Failed to connect to EnergyGauge COM interface: {e}")
            return False
    
    def discover_com_interface(self) -> Dict[str, Any]:
        """Discover available COM interface properties and methods"""
        if not self.connected:
            self.logger.error("Not connected to EnergyGauge for discovery")
            return {}
            
        discovery_data = {
            'application_properties': [],
            'application_methods': [],
            'project_properties': [],
            'project_methods': [],
            'com_interface_version': 'unknown'
        }
        
        try:
            # Discover application-level properties and methods
            if self.app:
                try:
                    app_props = []
                    app_methods = []
                    
                    # Try to get application properties
                    for attr_name in dir(self.app):
                        if not attr_name.startswith('_'):
                            try:
                                attr_value = getattr(self.app, attr_name)
                                if callable(attr_value):
                                    app_methods.append(attr_name)
                                else:
                                    app_props.append(f"{attr_name}: {type(attr_value).__name__}")
                            except:
                                app_props.append(f"{attr_name}: <property>")
                    
                    discovery_data['application_properties'] = app_props[:50]  # Limit output
                    discovery_data['application_methods'] = app_methods[:50]
                    
                except Exception as e:
                    self.logger.warning(f"Could not discover application interface: {e}")
            
            # Try to get version information
            try:
                if hasattr(self.app, 'Version'):
                    discovery_data['com_interface_version'] = str(self.app.Version)
                elif hasattr(self.app, 'ApplicationVersion'):
                    discovery_data['com_interface_version'] = str(self.app.ApplicationVersion)
            except:
                pass
            
            self.logger.info("COM interface discovery completed")
            return discovery_data
            
        except Exception as e:
            self.logger.error(f"Error during COM interface discovery: {e}")
            return discovery_data
            
    def disconnect(self):
        """Disconnect from EnergyGauge application"""
        try:
            if self.project:
                # TODO: Implement project closing based on actual COM interface
                # self.project.Close()
                pass
                
            if self.app:
                # TODO: Implement application quit based on actual COM interface
                # self.app.Quit()
                pass
                
            self.app = None
            self.project = None
            self.connected = False
            self.logger.info("Disconnected from EnergyGauge")
            
        except Exception as e:
            self.logger.error(f"Error during disconnect: {e}")
            
    def open_template(self, template_path: str) -> bool:
        """Open EnergyGauge template project"""
        if not self.connected:
            self.logger.error("Not connected to EnergyGauge")
            return False
            
        try:
            if not os.path.exists(template_path):
                self.logger.error(f"Template file not found: {template_path}")
                return False
                
            # TODO: Implement template opening based on actual COM interface
            # Example: self.project = self.app.Projects.Open(template_path)
            self.logger.info(f"TODO: Open template {template_path}")
            
            # For now, simulate success
            self.project = "simulated_project"
            self.logger.info(f"Opened template: {template_path}")
            return True
            
        except Exception as e:
            self.logger.error(f"Failed to open template: {e}")
            return False
            
    def set_project_info(self, project_info: Dict) -> bool:
        """Set project information fields"""
        if not self.project:
            self.logger.error("No project loaded")
            return False
            
        try:
            # TODO: Implement based on actual COM interface
            # Example mappings:
            # self.project.Name = project_info.get('name', '')
            # self.project.Address = project_info.get('address', '')
            # self.project.City = project_info.get('city', '')
            # self.project.State = project_info.get('state', '')
            # self.project.ZipCode = project_info.get('zip', '')
            
            self.logger.info(f"TODO: Set project info: {project_info}")
            return True
            
        except Exception as e:
            self.logger.error(f"Failed to set project info: {e}")
            return False
            
    def set_building_data(self, building_data: Dict) -> bool:
        """Set building data fields"""
        if not self.project:
            self.logger.error("No project loaded")
            return False
            
        try:
            # TODO: Implement based on actual COM interface
            # Example:
            # if 'conditioned_floor_area' in building_data:
            #     self.project.Building.ConditionedFloorArea = building_data['conditioned_floor_area']
            
            self.logger.info(f"TODO: Set building data: {building_data}")
            return True
            
        except Exception as e:
            self.logger.error(f"Failed to set building data: {e}")
            return False
            
    def set_windows(self, windows_data: Dict) -> bool:
        """Set window data"""
        if not self.project:
            self.logger.error("No project loaded")
            return False
            
        try:
            # TODO: Implement based on actual COM interface
            for orientation, data in windows_data.items():
                # Example:
                # window = self.project.Building.Windows.Add(orientation)
                # window.Area = data.get('area', 0)
                # window.UFactor = data.get('u_factor', 0)
                # window.SHGC = data.get('shgc', 0)
                pass
                
            self.logger.info(f"TODO: Set windows data: {windows_data}")
            return True
            
        except Exception as e:
            self.logger.error(f"Failed to set windows data: {e}")
            return False
            
    def set_hvac_system(self, hvac_data: Dict) -> bool:
        """Set HVAC system data"""
        if not self.project:
            self.logger.error("No project loaded")
            return False
            
        try:
            # TODO: Implement based on actual COM interface
            for system_id, data in hvac_data.items():
                # Example:
                # hvac = self.project.Building.HVAC.Add(system_id)
                # hvac.Tonnage = data.get('tonnage', 0)
                # hvac.SEER2 = data.get('seer2', 0)
                pass
                
            self.logger.info(f"TODO: Set HVAC data: {hvac_data}")
            return True
            
        except Exception as e:
            self.logger.error(f"Failed to set HVAC data: {e}")
            return False
            
    def calculate(self) -> bool:
        """Run calculation or compliance check"""
        if not self.project:
            self.logger.error("No project loaded")
            return False
            
        try:
            # TODO: Implement based on actual COM interface
            # Example: self.project.Calculate()
            self.logger.info("TODO: Run calculation")
            return True
            
        except Exception as e:
            self.logger.error(f"Failed to run calculation: {e}")
            return False
            
    def save_project(self, output_path: str) -> bool:
        """Save project to specified path"""
        if not self.project:
            self.logger.error("No project loaded")
            return False
            
        try:
            # TODO: Implement based on actual COM interface
            # Example: self.project.SaveAs(output_path)
            
            # For now, create a dummy file
            os.makedirs(os.path.dirname(output_path), exist_ok=True)
            with open(output_path, 'w') as f:
                f.write(f"# EnergyGauge project saved at {datetime.now()}\n")
                
            self.logger.info(f"TODO: Save project to {output_path}")
            return True
            
        except Exception as e:
            self.logger.error(f"Failed to save project: {e}")
            return False
            
    def export_report(self, output_path: str) -> bool:
        """Export results report"""
        if not self.project:
            self.logger.error("No project loaded")
            return False
            
        try:
            # TODO: Implement based on actual COM interface
            # Example: self.project.ExportReport(output_path)
            
            # For now, create a dummy report
            os.makedirs(os.path.dirname(output_path), exist_ok=True)
            with open(output_path, 'w') as f:
                f.write(f"EnergyGauge Results Report\nGenerated: {datetime.now()}\n")
                
            self.logger.info(f"TODO: Export report to {output_path}")
            return True
            
        except Exception as e:
            self.logger.error(f"Failed to export report: {e}")
            return False


class InputProcessor:
    """Processes and validates JSON input data"""
    
    def __init__(self, config: ConfigManager, logger: logging.Logger):
        self.config = config
        self.logger = logger
        
    def load_json(self, file_path: str) -> Optional[Dict]:
        """Load and validate JSON input file"""
        try:
            with open(file_path, 'r') as f:
                data = json.load(f)
                
            # Basic validation
            if not isinstance(data, dict):
                self.logger.error(f"Invalid JSON format in {file_path}")
                return None
                
            if 'lot_id' not in data:
                self.logger.error(f"Missing lot_id in {file_path}")
                return None
                
            self.logger.info(f"Loaded input data for lot: {data['lot_id']}")
            return data
            
        except json.JSONDecodeError as e:
            self.logger.error(f"JSON decode error in {file_path}: {e}")
            return None
        except Exception as e:
            self.logger.error(f"Failed to load {file_path}: {e}")
            return None
            
    def validate_and_fill_defaults(self, data: Dict) -> Tuple[Dict, List[str]]:
        """Validate data and fill in defaults for missing fields"""
        warnings = []
        
        # Default values for missing fields
        defaults = {
            'project_info': {},
            'building_data': {
                'conditioned_floor_area': 2000.0,
                'windows': {},
                'walls': {},
                'infiltration': {'ach50': 7.0}
            },
            'hvac': {
                'system1': {'tonnage': 2.5, 'seer2': 14.0}
            },
            'duct': {
                'location': 'Interior'
            }
        }
        
        # Fill missing top-level sections
        for key, default_value in defaults.items():
            if key not in data:
                data[key] = default_value
                warnings.append(f"Missing {key} section, using defaults")
                
        # Validate required fields in project_info
        project_info = data.get('project_info', {})
        required_project_fields = ['name', 'address', 'city', 'state', 'zip']
        for field in required_project_fields:
            if field not in project_info:
                project_info[field] = f"Default_{field}"
                warnings.append(f"Missing project_info.{field}, using default")
                
        # Validate building data
        building_data = data.get('building_data', {})
        if 'conditioned_floor_area' not in building_data:
            building_data['conditioned_floor_area'] = defaults['building_data']['conditioned_floor_area']
            warnings.append("Missing conditioned_floor_area, using default 2000 sqft")
            
        return data, warnings
        
    def get_input_files(self, inputs_dir: str) -> List[str]:
        """Get list of JSON input files to process"""
        try:
            files = []
            for file in os.listdir(inputs_dir):
                if file.lower().endswith('.json'):
                    files.append(os.path.join(inputs_dir, file))
            return sorted(files)
        except Exception as e:
            self.logger.error(f"Failed to get input files from {inputs_dir}: {e}")
            return []


class EnergyGaugeAutomation:
    """Main automation class that coordinates all components"""
    
    def __init__(self, project_root: str = None):
        # Set up project paths
        if project_root is None:
            project_root = os.path.dirname(os.path.abspath(__file__))
            project_root = os.path.dirname(project_root)  # Go up one level from code/
            
        self.project_root = Path(project_root)
        self.inputs_dir = self.project_root / "inputs"
        self.outputs_dir = self.project_root / "outputs"
        self.templates_dir = self.project_root / "templates"
        self.archive_dir = self.project_root / "archive"
        
        # Create directories if they don't exist
        for dir_path in [self.inputs_dir, self.outputs_dir, self.templates_dir, self.archive_dir]:
            dir_path.mkdir(parents=True, exist_ok=True)
            
        # Set up logging
        self.setup_logging()
        
        # Initialize components
        self.config = ConfigManager()
        self.quality_indicator = QualityIndicator()
        self.input_processor = InputProcessor(self.config, self.logger)
        self.ocr_processor = OCRProcessor(self.config, self.logger)
        self.duplicate_detector = DuplicateDetector(self.config, self.logger)
        self.archive_manager = ArchiveManager(self.logger)
        self.energygauge_com = EnergyGaugeCOMInterface(self.config, self.logger)
        
        # GUI will be initialized separately
        self.gui = None
        
    def setup_logging(self):
        """Set up logging configuration"""
        log_file = self.archive_dir / "automation.log"
        
        logging.basicConfig(
            level=logging.INFO,
            format='%(asctime)s - %(levelname)s - %(message)s',
            handlers=[
                logging.FileHandler(log_file),
                logging.StreamHandler()
            ]
        )
        
        self.logger = logging.getLogger(__name__)
        self.logger.info("EnergyGauge Automation initialized")
        
    def process_single_lot(self, input_file: str) -> bool:
        """Process a single lot input file"""
        try:
            # Load and validate input data
            input_data = self.input_processor.load_json(input_file)
            if not input_data:
                return False
                
            lot_id = input_data['lot_id']
            self.quality_indicator.set_current_lot(lot_id)
            self.quality_indicator.set_data_quality("GREEN", f"Processing {lot_id}")
            
            if self.gui:
                self.gui.log_message(f"Starting processing for {lot_id}")
                
            # Check for duplicates
            if self.duplicate_detector.is_duplicate(input_data):
                self.quality_indicator.set_data_quality("YELLOW", f"{lot_id}: Duplicate detected, skipping")
                if self.gui:
                    self.gui.log_message(f"{lot_id}: Duplicate detected, skipping")
                return True
                
            # Validate and fill defaults
            input_data, warnings = self.input_processor.validate_and_fill_defaults(input_data)
            
            if warnings:
                self.quality_indicator.set_data_quality("YELLOW", f"{lot_id}: Using defaults for missing data")
                for warning in warnings:
                    self.logger.warning(f"{lot_id}: {warning}")
                    if self.gui:
                        self.gui.log_message(f"{lot_id}: {warning}")
                        
            # Process OCR if floor plan image is provided
            floor_plan_path = input_data.get('floor_plan_image')
            if floor_plan_path and os.path.exists(floor_plan_path):
                if self.gui:
                    self.gui.log_message(f"{lot_id}: Processing floor plan image")
                    
                ocr_result = self.ocr_processor.extract_text(floor_plan_path)
                if ocr_result['success']:
                    # Try to extract missing fields from OCR
                    self.enhance_data_with_ocr(input_data, ocr_result['text'])
                else:
                    self.quality_indicator.set_data_quality("YELLOW", f"{lot_id}: OCR failed or low confidence")
                    
            # Connect to EnergyGauge
            if not self.energygauge_com.connected:
                if not self.energygauge_com.connect():
                    self.quality_indicator.set_system_health("RED", f"{lot_id}: Failed to connect to EnergyGauge")
                    return False
                    
            # Open template
            template_file = self.templates_dir / self.config.get('template_file', 'YourTemplate.egpj')
            if not self.energygauge_com.open_template(str(template_file)):
                self.quality_indicator.set_system_health("RED", f"{lot_id}: Failed to open template")
                return False
                
            # Set project data
            if not self.set_energygauge_data(input_data):
                self.quality_indicator.set_system_health("YELLOW", f"{lot_id}: Some data could not be set")
                
            # Run calculation
            if not self.energygauge_com.calculate():
                self.quality_indicator.set_system_health("YELLOW", f"{lot_id}: Calculation may have failed")
                
            # Save results
            output_dir = self.outputs_dir / lot_id
            output_dir.mkdir(exist_ok=True)
            
            project_file = output_dir / f"{lot_id}.egpj"
            report_file = output_dir / f"{lot_id}_report.txt"
            
            self.energygauge_com.save_project(str(project_file))
            self.energygauge_com.export_report(str(report_file))
            
            # Archive input and output
            self.archive_manager.archive_input(lot_id, input_data, floor_plan_path)
            self.archive_manager.archive_output(lot_id, str(output_dir))
            
            # Mark as processed
            self.duplicate_detector.mark_as_processed(input_data)
            
            # Log success
            self.archive_manager.log_processing_result(lot_id, "SUCCESS")
            self.quality_indicator.add_processed_lot(lot_id, "SUCCESS")
            
            if self.gui:
                self.gui.log_message(f"{lot_id}: Processing completed successfully")
                
            self.logger.info(f"Successfully processed {lot_id}")
            return True
            
        except Exception as e:
            error_msg = f"Error processing {lot_id}: {str(e)}"
            self.logger.error(error_msg)
            self.logger.error(traceback.format_exc())
            
            self.quality_indicator.set_system_health("RED", error_msg)
            if self.gui:
                self.gui.log_message(f"ERROR: {error_msg}")
                
            # Log failure
            if 'lot_id' in locals():
                self.archive_manager.log_processing_result(lot_id, "FAILED", str(e))
                self.quality_indicator.add_processed_lot(lot_id, "FAILED")
                
            return False
            
    def enhance_data_with_ocr(self, input_data: Dict, ocr_text: str):
        """Enhance input data with OCR-extracted information"""
        building_data = input_data.get('building_data', {})
        
        # Try to extract conditioned floor area if missing
        if 'conditioned_floor_area' not in building_data or not building_data['conditioned_floor_area']:
            area = self.ocr_processor.extract_field_from_text(ocr_text, 'conditioned_floor_area')
            if area:
                building_data['conditioned_floor_area'] = float(area)
                self.logger.info(f"OCR extracted floor area: {area}")
                
        # Try to extract HVAC tonnage if missing
        hvac_data = input_data.get('hvac', {})
        for system_id, system_data in hvac_data.items():
            if 'tonnage' not in system_data or not system_data['tonnage']:
                tonnage = self.ocr_processor.extract_field_from_text(ocr_text, 'tonnage')
                if tonnage:
                    system_data['tonnage'] = float(tonnage)
                    self.logger.info(f"OCR extracted tonnage: {tonnage}")
                    
    def set_energygauge_data(self, input_data: Dict) -> bool:
        """Set all data in EnergyGauge"""
        success = True
        
        # Set project information
        if 'project_info' in input_data:
            if not self.energygauge_com.set_project_info(input_data['project_info']):
                success = False
                
        # Set building data
        if 'building_data' in input_data:
            if not self.energygauge_com.set_building_data(input_data['building_data']):
                success = False
                
            # Set windows
            windows_data = input_data['building_data'].get('windows', {})
            if windows_data and not self.energygauge_com.set_windows(windows_data):
                success = False
                
        # Set HVAC data
        if 'hvac' in input_data:
            if not self.energygauge_com.set_hvac_system(input_data['hvac']):
                success = False
                
        return success
        
    def process_all_inputs(self):
        """Process all input files in the inputs directory"""
        input_files = self.input_processor.get_input_files(str(self.inputs_dir))
        
        if not input_files:
            self.logger.warning("No input files found")
            if self.gui:
                self.gui.log_message("No input files found in inputs directory")
            return
            
        self.logger.info(f"Found {len(input_files)} input files to process")
        if self.gui:
            self.gui.log_message(f"Found {len(input_files)} input files to process")
            
        processed_count = 0
        failed_count = 0
        
        for input_file in input_files:
            if self.gui and not self.gui.processing:
                self.logger.info("Processing stopped by user")
                break
                
            if self.process_single_lot(input_file):
                processed_count += 1
            else:
                failed_count += 1
                
        # Final status
        total_files = len(input_files)
        self.logger.info(f"Processing complete: {processed_count}/{total_files} successful, {failed_count} failed")
        
        if self.gui:
            self.gui.log_message(f"Processing complete: {processed_count}/{total_files} successful, {failed_count} failed")
            
        # Set final status indicators
        if failed_count == 0:
            self.quality_indicator.set_system_health("GREEN", "All lots processed successfully")
        elif failed_count < total_files:
            self.quality_indicator.set_system_health("YELLOW", f"{failed_count} lots failed processing")
        else:
            self.quality_indicator.set_system_health("RED", "All lots failed processing")
            
    def run_gui_mode(self):
        """Run the automation with GUI"""
        self.gui = StoplightGUI(self.quality_indicator)
        self.gui.automation_callback = self.process_all_inputs
        
        self.logger.info("Starting GUI mode")
        self.gui.run()
        
    def run_batch_mode(self):
        """Run the automation in batch mode (no GUI)"""
        self.logger.info("Starting batch mode")
        self.process_all_inputs()
        
    def process_multiple_lots_parallel(self, max_workers: int = 3) -> List[Dict]:
        """Process multiple lots in parallel for faster throughput"""
        self.logger.info(f"Starting parallel processing with {max_workers} workers")
        
        # Get input files
        input_files = list(self.inputs_dir.glob("*.json"))
        
        if not input_files:
            self.logger.warning("No input files found for parallel processing")
            return []
        
        results = []
        
        # Process files in parallel using ThreadPoolExecutor
        with ThreadPoolExecutor(max_workers=max_workers) as executor:
            # Submit all jobs
            futures = []
            for input_file in input_files:
                lot_id = input_file.stem
                if lot_id.endswith("_inputs") or lot_id.endswith("_input"):
                    lot_id = lot_id.rsplit("_", 1)[0]
                
                # Load input data
                try:
                    with open(input_file, 'r') as f:
                        input_data = json.load(f)
                    
                    # Submit the processing job
                    future = executor.submit(self.process_single_lot, lot_id, input_data)
                    futures.append((lot_id, future, input_file))
                    
                except Exception as e:
                    self.logger.error(f"Error loading input file {input_file}: {e}")
                    results.append({
                        "lot_id": lot_id,
                        "input_file": str(input_file),
                        "status": "failed",
                        "error": str(e),
                        "start_time": datetime.now().isoformat(),
                        "end_time": datetime.now().isoformat()
                    })
            
            # Collect results as they complete
            for lot_id, future, input_file in futures:
                start_time = datetime.now()
                try:
                    success = future.result(timeout=600)  # 10 minute timeout per lot
                    end_time = datetime.now()
                    
                    results.append({
                        "lot_id": lot_id,
                        "input_file": str(input_file),
                        "status": "success" if success else "failed",
                        "start_time": start_time.isoformat(),
                        "end_time": end_time.isoformat(),
                        "duration_seconds": (end_time - start_time).total_seconds()
                    })
                    
                    self.logger.info(f"Parallel processing completed for {lot_id}: {'SUCCESS' if success else 'FAILED'}")
                    
                except Exception as e:
                    end_time = datetime.now()
                    error_msg = str(e)
                    self.logger.error(f"Parallel processing error for {lot_id}: {error_msg}")
                    
                    results.append({
                        "lot_id": lot_id,
                        "input_file": str(input_file),
                        "status": "failed",
                        "error": error_msg,
                        "start_time": start_time.isoformat(),
                        "end_time": end_time.isoformat(),
                        "duration_seconds": (end_time - start_time).total_seconds()
                    })
        
        # Summary statistics
        successful = len([r for r in results if r["status"] == "success"])
        failed = len([r for r in results if r["status"] == "failed"])
        total_time = sum(r.get("duration_seconds", 0) for r in results)
        
        self.logger.info(f"Parallel processing complete: {successful}/{len(results)} successful, {failed} failed")
        self.logger.info(f"Total processing time: {total_time:.2f} seconds")
        
        # Save parallel processing report
        report_file = self.archive_dir / f"parallel_processing_report_{datetime.now().strftime('%Y%m%d_%H%M%S')}.json"
        with open(report_file, 'w') as f:
            json.dump({
                "summary": {
                    "total_lots": len(results),
                    "successful_lots": successful,
                    "failed_lots": failed,
                    "success_rate": successful / len(results) if results else 0,
                    "total_duration_seconds": total_time,
                    "max_workers": max_workers,
                    "timestamp": datetime.now().isoformat()
                },
                "detailed_results": results
            }, f, indent=2)
        
        self.logger.info(f"Parallel processing report saved: {report_file}")
        return results
        
    def cleanup(self):
        """Clean up resources"""
        if self.energygauge_com:
            self.energygauge_com.disconnect()
            
        self.logger.info("Cleanup completed")


def create_sample_input():
    """Create a sample input file for testing"""
    sample_data = {
        "lot_id": "Lot101_Sample",
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
    
    inputs_dir = Path("inputs")
    inputs_dir.mkdir(exist_ok=True)
    
    with open(inputs_dir / "Lot101_Sample_inputs.json", 'w') as f:
        json.dump(sample_data, f, indent=2)
        
    print("Sample input file created: inputs/Lot101_Sample_inputs.json")


class ModernEnergyGaugeGUI:
    """Modern GUI for EnergyGauge automation with schematic upload and approval workflow"""
    
    def __init__(self, automation: 'EnergyGaugeAutomation'):
        self.automation = automation
        self.root = tk.Tk()
        self.root.title("EnergyGauge USA Automation Suite")
        self.root.geometry("1200x800")
        self.root.configure(bg='#f0f0f0')
        
        # Data storage
        self.uploaded_files = []
        self.processing_queue = []
        self.current_project = None
        self.processing_active = False
        
        # Create GUI components
        self.setup_gui()
        
    def setup_gui(self):
        """Setup the modern GUI interface"""
        # Main container with padding
        main_frame = ttk.Frame(self.root)
        main_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        
        # Header section
        self.create_header(main_frame)
        
        # Create notebook for tabs
        self.notebook = ttk.Notebook(main_frame)
        self.notebook.pack(fill=tk.BOTH, expand=True, pady=(10, 0))
        
        # Tab 1: Project Setup & Upload
        self.setup_upload_tab()
        
        # Tab 2: Processing & Approval
        self.setup_processing_tab()
        
        # Tab 3: Results & History
        self.setup_results_tab()
        
        # Tab 4: Settings & Discovery
        self.setup_settings_tab()
        
        # Status bar
        self.create_status_bar(main_frame)
        
    def create_header(self, parent):
        """Create application header"""
        header_frame = ttk.Frame(parent)
        header_frame.pack(fill=tk.X, pady=(0, 10))
        
        # Title
        title_label = ttk.Label(header_frame, text="EnergyGauge USA Automation Suite", 
                               font=('Arial', 16, 'bold'))
        title_label.pack(side=tk.LEFT)
        
        # System status indicators
        status_frame = ttk.Frame(header_frame)
        status_frame.pack(side=tk.RIGHT)
        
        # EnergyGauge connection status
        self.eg_status_var = tk.StringVar(value="Disconnected")
        ttk.Label(status_frame, text="EnergyGauge: ").pack(side=tk.LEFT)
        
        self.eg_status_indicator = ttk.Label(status_frame, textvariable=self.eg_status_var,
                                           foreground="red")
        self.eg_status_indicator.pack(side=tk.LEFT, padx=(0, 20))
        
    def setup_upload_tab(self):
        """Setup the upload and project configuration tab"""
        upload_frame = ttk.Frame(self.notebook)
        self.notebook.add(upload_frame, text="📁 Project Setup")
        
        # Left panel - File upload
        left_panel = ttk.LabelFrame(upload_frame, text="Schematic & Data Upload", padding="10")
        left_panel.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=(0, 5))
        
        # Upload area
        upload_area = ttk.Frame(left_panel)
        upload_area.pack(fill=tk.BOTH, expand=True)
        
        # File list with details
        self.upload_listbox = tk.Listbox(upload_area, height=12)
        self.upload_listbox.pack(fill=tk.BOTH, expand=True, pady=(0, 10))
        
        # Upload buttons
        button_frame = ttk.Frame(upload_area)
        button_frame.pack(fill=tk.X)
        
        ttk.Button(button_frame, text="📎 Upload Schematics", 
                  command=self.upload_schematics).pack(side=tk.TOP, fill=tk.X, pady=(0, 5))
        ttk.Button(button_frame, text="📄 Upload JSON Data", 
                  command=self.upload_json_data).pack(side=tk.TOP, fill=tk.X, pady=(0, 5))
        ttk.Button(button_frame, text="🗑️ Clear All", 
                  command=self.clear_uploads).pack(side=tk.TOP, fill=tk.X)
        
        # Right panel - Project configuration
        right_panel = ttk.LabelFrame(upload_frame, text="Project Configuration", padding="10")
        right_panel.pack(side=tk.RIGHT, fill=tk.BOTH, expand=True, padx=(5, 0))
        
        # Project info form
        self.create_project_form(right_panel)
        
    def setup_processing_tab(self):
        """Setup the processing and approval tab"""
        processing_frame = ttk.Frame(self.notebook)
        self.notebook.add(processing_frame, text="⚙️ Processing & Approval")
        
        # Top section - Processing queue
        queue_frame = ttk.LabelFrame(processing_frame, text="Processing Queue", padding="10")
        queue_frame.pack(fill=tk.X, pady=(0, 10))
        
        # Queue listbox with details
        self.queue_tree = ttk.Treeview(queue_frame, columns=('Status', 'Progress'), 
                                      show='tree headings', height=6)
        self.queue_tree.heading('#0', text='Project/Lot')
        self.queue_tree.heading('Status', text='Status')
        self.queue_tree.heading('Progress', text='Progress')
        self.queue_tree.pack(fill=tk.X)
        
        # Processing controls
        control_frame = ttk.Frame(queue_frame)
        control_frame.pack(fill=tk.X, pady=(10, 0))
        
        self.start_btn = ttk.Button(control_frame, text="▶️ Start Processing", 
                                   command=self.start_processing, state=tk.DISABLED)
        self.start_btn.pack(side=tk.LEFT, padx=(0, 5))
        
        self.stop_btn = ttk.Button(control_frame, text="⏹️ Stop", 
                                  command=self.stop_processing, state=tk.DISABLED)
        self.stop_btn.pack(side=tk.LEFT)
        
        # Bottom section - Approval workflow
        approval_frame = ttk.LabelFrame(processing_frame, text="Approval Workflow", padding="10")
        approval_frame.pack(fill=tk.BOTH, expand=True, pady=(10, 0))
        
        # Preview area
        preview_paned = ttk.PanedWindow(approval_frame, orient=tk.HORIZONTAL)
        preview_paned.pack(fill=tk.BOTH, expand=True)
        
        # Left - Image preview
        image_frame = ttk.LabelFrame(preview_paned, text="Schematic Preview")
        preview_paned.add(image_frame, weight=1)
        
        self.image_canvas = tk.Canvas(image_frame, bg='white', width=400, height=300)
        self.image_canvas.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
        
        # Right - Data preview and approval
        data_frame = ttk.LabelFrame(preview_paned, text="Extracted Data & Parameters")
        preview_paned.add(data_frame, weight=1)
        
        # Data display
        self.data_text = scrolledtext.ScrolledText(data_frame, height=15, width=50)
        self.data_text.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
        
        # Approval buttons
        approval_btn_frame = ttk.Frame(approval_frame)
        approval_btn_frame.pack(fill=tk.X, pady=(10, 0))
        
        self.approve_btn = ttk.Button(approval_btn_frame, text="✅ Approve & Process", 
                                     command=self.approve_current, state=tk.DISABLED)
        self.approve_btn.pack(side=tk.LEFT, padx=(0, 10))
        
        self.reject_btn = ttk.Button(approval_btn_frame, text="❌ Reject & Edit", 
                                    command=self.reject_current, state=tk.DISABLED)
        self.reject_btn.pack(side=tk.LEFT, padx=(0, 10))
        
        self.skip_btn = ttk.Button(approval_btn_frame, text="⏭️ Skip for Now", 
                                  command=self.skip_current, state=tk.DISABLED)
        self.skip_btn.pack(side=tk.LEFT)
        
    def setup_results_tab(self):
        """Setup the results and history tab"""
        results_frame = ttk.Frame(self.notebook)
        self.notebook.add(results_frame, text="📊 Results & History")
        
        # Results tree
        self.results_tree = ttk.Treeview(results_frame, 
                                       columns=('Date', 'Status', 'Lots', 'Duration'), 
                                       show='tree headings')
        self.results_tree.heading('#0', text='Project')
        self.results_tree.heading('Date', text='Date')
        self.results_tree.heading('Status', text='Status')
        self.results_tree.heading('Lots', text='Lots Processed')
        self.results_tree.heading('Duration', text='Duration')
        self.results_tree.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        
        # Results controls
        results_control_frame = ttk.Frame(results_frame)
        results_control_frame.pack(fill=tk.X, padx=10, pady=(0, 10))
        
        ttk.Button(results_control_frame, text="📂 Open Output Folder", 
                  command=self.open_output_folder).pack(side=tk.LEFT, padx=(0, 5))
        ttk.Button(results_control_frame, text="📈 Generate Report", 
                  command=self.generate_report).pack(side=tk.LEFT, padx=(0, 5))
        ttk.Button(results_control_frame, text="🔄 Refresh", 
                  command=self.refresh_results).pack(side=tk.LEFT)
        
    def setup_settings_tab(self):
        """Setup the settings and COM discovery tab"""
        settings_frame = ttk.Frame(self.notebook)
        self.notebook.add(settings_frame, text="⚙️ Settings & Discovery")
        
        # EnergyGauge connection section
        eg_section = ttk.LabelFrame(settings_frame, text="EnergyGauge Connection", padding="10")
        eg_section.pack(fill=tk.X, padx=10, pady=10)
        
        conn_frame = ttk.Frame(eg_section)
        conn_frame.pack(fill=tk.X)
        
        ttk.Button(conn_frame, text="🔗 Connect to EnergyGauge", 
                  command=self.connect_energygauge).pack(side=tk.LEFT, padx=(0, 5))
        ttk.Button(conn_frame, text="🔍 Discover COM Interface", 
                  command=self.discover_com_interface).pack(side=tk.LEFT, padx=(0, 5))
        ttk.Button(conn_frame, text="👁️ Make EnergyGauge Visible", 
                  command=self.make_energygauge_visible).pack(side=tk.LEFT)
        
        # Discovery results
        discovery_section = ttk.LabelFrame(settings_frame, text="COM Interface Discovery", padding="10")
        discovery_section.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        
        self.discovery_text = scrolledtext.ScrolledText(discovery_section, height=20)
        self.discovery_text.pack(fill=tk.BOTH, expand=True)
        
    def create_project_form(self, parent):
        """Create project information form"""
        # Project name
        ttk.Label(parent, text="Project Name:").pack(anchor=tk.W)
        self.project_name_var = tk.StringVar(value="Sample Project")
        ttk.Entry(parent, textvariable=self.project_name_var, width=40).pack(fill=tk.X, pady=(0, 10))
        
        # Address
        ttk.Label(parent, text="Address:").pack(anchor=tk.W)
        self.address_var = tk.StringVar(value="123 Main Street")
        ttk.Entry(parent, textvariable=self.address_var, width=40).pack(fill=tk.X, pady=(0, 10))
        
        # City, State, ZIP
        location_frame = ttk.Frame(parent)
        location_frame.pack(fill=tk.X, pady=(0, 10))
        
        city_frame = ttk.Frame(location_frame)
        city_frame.pack(fill=tk.X, pady=(0, 5))
        ttk.Label(city_frame, text="City:").pack(anchor=tk.W)
        self.city_var = tk.StringVar(value="Orlando")
        ttk.Entry(city_frame, textvariable=self.city_var, width=30).pack(fill=tk.X)
        
        state_zip_frame = ttk.Frame(location_frame)
        state_zip_frame.pack(fill=tk.X)
        
        state_subframe = ttk.Frame(state_zip_frame)
        state_subframe.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 5))
        ttk.Label(state_subframe, text="State:").pack(anchor=tk.W)
        self.state_var = tk.StringVar(value="FL")
        ttk.Entry(state_subframe, textvariable=self.state_var, width=5).pack(fill=tk.X)
        
        zip_subframe = ttk.Frame(state_zip_frame)
        zip_subframe.pack(side=tk.RIGHT, fill=tk.X, expand=True, padx=(5, 0))
        ttk.Label(zip_subframe, text="ZIP:").pack(anchor=tk.W)
        self.zip_var = tk.StringVar(value="32801")
        ttk.Entry(zip_subframe, textvariable=self.zip_var, width=10).pack(fill=tk.X)
        
        # Template file selection
        template_frame = ttk.Frame(parent)
        template_frame.pack(fill=tk.X, pady=(20, 0))
        
        ttk.Label(template_frame, text="EnergyGauge Template:").pack(anchor=tk.W)
        template_select_frame = ttk.Frame(template_frame)
        template_select_frame.pack(fill=tk.X, pady=(5, 0))
        
        self.template_var = tk.StringVar(value="templates/YourTemplate.egpj")
        ttk.Entry(template_select_frame, textvariable=self.template_var).pack(side=tk.LEFT, fill=tk.X, expand=True)
        ttk.Button(template_select_frame, text="Browse", 
                  command=self.browse_template).pack(side=tk.RIGHT, padx=(5, 0))
        
    def create_status_bar(self, parent):
        """Create status bar at bottom"""
        status_frame = ttk.Frame(parent)
        status_frame.pack(fill=tk.X, pady=(10, 0))
        
        # Progress bar
        self.progress_var = tk.DoubleVar()
        self.progress_bar = ttk.Progressbar(status_frame, variable=self.progress_var, 
                                          maximum=100, length=300)
        self.progress_bar.pack(side=tk.LEFT, padx=(0, 10))
        
        # Status text
        self.status_var = tk.StringVar(value="Ready - Upload files to begin")
        status_label = ttk.Label(status_frame, textvariable=self.status_var)
        status_label.pack(side=tk.LEFT)
        
    # Event handlers
    def upload_schematics(self):
        """Handle schematic file upload"""
        files = filedialog.askopenfilenames(
            title="Select Schematic Files",
            filetypes=[
                ("Image files", "*.png *.jpg *.jpeg *.gif *.bmp *.tiff"),
                ("PDF files", "*.pdf"),
                ("All files", "*.*")
            ]
        )
        
        for file_path in files:
            if file_path not in [f[0] for f in self.uploaded_files]:
                self.uploaded_files.append((file_path, "schematic"))
                filename = os.path.basename(file_path)
                self.upload_listbox.insert(tk.END, f"📐 {filename}")
        
        self.update_processing_queue()
        
    def upload_json_data(self):
        """Handle JSON data file upload"""
        files = filedialog.askopenfilenames(
            title="Select JSON Data Files",
            filetypes=[
                ("JSON files", "*.json"),
                ("All files", "*.*")
            ]
        )
        
        for file_path in files:
            if file_path not in [f[0] for f in self.uploaded_files]:
                self.uploaded_files.append((file_path, "json"))
                filename = os.path.basename(file_path)
                self.upload_listbox.insert(tk.END, f"📄 {filename}")
        
        self.update_processing_queue()
        
    def clear_uploads(self):
        """Clear all uploaded files"""
        self.uploaded_files.clear()
        self.upload_listbox.delete(0, tk.END)
        self.update_processing_queue()
        
    def browse_template(self):
        """Browse for EnergyGauge template file"""
        file_path = filedialog.askopenfilename(
            title="Select EnergyGauge Template",
            filetypes=[
                ("EnergyGauge files", "*.egpj *.egp"),
                ("All files", "*.*")
            ]
        )
        
        if file_path:
            self.template_var.set(file_path)
            
    def update_processing_queue(self):
        """Update the processing queue display"""
        # Clear existing items
        for item in self.queue_tree.get_children():
            self.queue_tree.delete(item)
            
        # Add uploaded files to queue
        json_files = [f for f in self.uploaded_files if f[1] == "json"]
        
        for json_file, _ in json_files:
            lot_name = os.path.splitext(os.path.basename(json_file))[0]
            if lot_name.endswith("_inputs") or lot_name.endswith("_input"):
                lot_name = lot_name.rsplit("_", 1)[0]
            self.queue_tree.insert('', 'end', text=lot_name, 
                                 values=('Ready', '0%'))
        
        # Enable/disable start button based on files and template
        has_data = len(json_files) > 0
        has_template = bool(self.template_var.get().strip())
        self.start_btn.config(state=tk.NORMAL if (has_data and has_template) else tk.DISABLED)
        
        # Update status
        if has_data and has_template:
            self.status_var.set(f"Ready to process {len(json_files)} lot(s)")
        elif has_data:
            self.status_var.set("Select EnergyGauge template to continue")
        else:
            self.status_var.set("Upload JSON data files to begin")
        
    def start_processing(self):
        """Start the processing workflow"""
        if not self.uploaded_files:
            messagebox.showwarning("No Files", "Please upload files first.")
            return
            
        if not self.template_var.get().strip():
            messagebox.showwarning("No Template", "Please select an EnergyGauge template file.")
            return
            
        self.processing_active = True
        self.status_var.set("Processing started...")
        self.start_btn.config(state=tk.DISABLED)
        self.stop_btn.config(state=tk.NORMAL)
        
        # Switch to processing tab
        self.notebook.select(1)
        
        # Start processing in background thread
        threading.Thread(target=self.process_files_background, daemon=True).start()
        
    def process_files_background(self):
        """Background processing of files"""
        try:
            json_files = [f for f in self.uploaded_files if f[1] == "json"]
            
            for i, (json_file, _) in enumerate(json_files):
                if not self.processing_active:
                    break
                    
                # Update progress
                progress = (i / len(json_files)) * 100
                self.root.after(0, lambda p=progress: self.progress_var.set(p))
                
                # Load and process the file
                try:
                    with open(json_file, 'r') as f:
                        data = json.load(f)
                    
                    # Show in approval workflow
                    self.root.after(0, lambda f=json_file, d=data: self.show_for_approval(f, d))
                    
                    # Simulate processing time
                    time.sleep(3)
                    
                    # Update queue status
                    lot_name = os.path.splitext(os.path.basename(json_file))[0]
                    if lot_name.endswith("_inputs") or lot_name.endswith("_input"):
                        lot_name = lot_name.rsplit("_", 1)[0]
                    
                    self.root.after(0, lambda ln=lot_name: self.update_queue_status(ln, "Completed", "100%"))
                    
                except Exception as e:
                    self.root.after(0, lambda e=e: messagebox.showerror("Processing Error", f"Error processing {json_file}: {e}"))
                
            self.root.after(0, self.processing_complete)
            
        except Exception as e:
            self.root.after(0, lambda e=e: messagebox.showerror("Processing Error", str(e)))
            
    def show_for_approval(self, json_file, data):
        """Show current item for approval"""
        # Update data preview
        self.data_text.delete(1.0, tk.END)
        self.data_text.insert(tk.END, "=== PROJECT PARAMETERS ===\n\n")
        self.data_text.insert(tk.END, json.dumps(data, indent=2))
        
        # Enable approval buttons
        self.approve_btn.config(state=tk.NORMAL)
        self.reject_btn.config(state=tk.NORMAL)
        self.skip_btn.config(state=tk.NORMAL)
        
        # Update status
        filename = os.path.basename(json_file)
        self.status_var.set(f"Review: {filename}")
        
        # Load schematic preview if available
        if 'floor_plan_image' in data:
            self.load_schematic_preview(data['floor_plan_image'])
        
    def load_schematic_preview(self, image_path):
        """Load and display schematic preview"""
        try:
            if Image and os.path.exists(image_path):
                # Load and resize image
                img = Image.open(image_path)
                img.thumbnail((350, 250), Image.Resampling.LANCZOS)
                
                # Convert to PhotoImage
                photo = ImageTk.PhotoImage(img)
                
                # Display in canvas
                self.image_canvas.delete("all")
                canvas_width = self.image_canvas.winfo_width()
                canvas_height = self.image_canvas.winfo_height()
                x = (canvas_width - img.width) // 2
                y = (canvas_height - img.height) // 2
                self.image_canvas.create_image(x, y, anchor=tk.NW, image=photo)
                
                # Keep reference to prevent garbage collection
                self.image_canvas.image = photo
            else:
                self.image_canvas.delete("all")
                self.image_canvas.create_text(200, 150, text="No schematic preview available", 
                                            fill="gray")
        except Exception as e:
            self.image_canvas.delete("all")
            self.image_canvas.create_text(200, 150, text=f"Preview error: {e}", 
                                        fill="red")
        
    def update_queue_status(self, lot_name, status, progress):
        """Update queue item status"""
        for item in self.queue_tree.get_children():
            if self.queue_tree.item(item, 'text') == lot_name:
                self.queue_tree.item(item, values=(status, progress))
                break
        
    def approve_current(self):
        """Approve current item and continue processing"""
        self.status_var.set("Processing approved item...")
        self.approve_btn.config(state=tk.DISABLED)
        self.reject_btn.config(state=tk.DISABLED)
        self.skip_btn.config(state=tk.DISABLED)
        
    def reject_current(self):
        """Reject current item and request editing"""
        messagebox.showinfo("Item Rejected", "Item has been rejected for editing.")
        self.approve_btn.config(state=tk.DISABLED)
        self.reject_btn.config(state=tk.DISABLED)
        self.skip_btn.config(state=tk.DISABLED)
        
    def skip_current(self):
        """Skip current item and continue"""
        self.approve_btn.config(state=tk.DISABLED)
        self.reject_btn.config(state=tk.DISABLED)
        self.skip_btn.config(state=tk.DISABLED)
        
    def stop_processing(self):
        """Stop processing"""
        self.processing_active = False
        self.status_var.set("Processing stopped")
        self.start_btn.config(state=tk.NORMAL)
        self.stop_btn.config(state=tk.DISABLED)
        
    def processing_complete(self):
        """Handle processing completion"""
        self.processing_active = False
        self.progress_var.set(100)
        self.status_var.set("Processing complete!")
        self.start_btn.config(state=tk.NORMAL)
        self.stop_btn.config(state=tk.DISABLED)
        messagebox.showinfo("Complete", "Processing completed successfully!")
        
        # Switch to results tab
        self.notebook.select(2)
        self.refresh_results()
        
    def connect_energygauge(self):
        """Connect to EnergyGauge via COM"""
        try:
            success = self.automation.com_interface.connect(visible=False)
            if success:
                self.eg_status_var.set("Connected")
                self.eg_status_indicator.config(foreground="green")
                self.status_var.set("Connected to EnergyGauge")
                messagebox.showinfo("Success", "Successfully connected to EnergyGauge USA")
            else:
                self.eg_status_var.set("Failed")
                self.eg_status_indicator.config(foreground="red")
                self.status_var.set("Failed to connect to EnergyGauge")
                messagebox.showerror("Connection Failed", "Could not connect to EnergyGauge. Ensure it's installed and try running as Administrator.")
        except Exception as e:
            messagebox.showerror("Connection Error", f"Failed to connect: {e}")
            
    def make_energygauge_visible(self):
        """Make EnergyGauge application visible for debugging"""
        try:
            success = self.automation.com_interface.connect(visible=True)
            if success:
                self.eg_status_var.set("Connected (Visible)")
                self.eg_status_indicator.config(foreground="green")
                self.status_var.set("EnergyGauge is now visible")
                messagebox.showinfo("Success", "EnergyGauge application is now visible for debugging")
            else:
                messagebox.showerror("Error", "Failed to make EnergyGauge visible")
        except Exception as e:
            messagebox.showerror("Error", f"Failed to make EnergyGauge visible: {e}")
            
    def discover_com_interface(self):
        """Discover and display COM interface details"""
        try:
            if not self.automation.com_interface.connected:
                success = self.automation.com_interface.connect()
                if not success:
                    messagebox.showerror("Error", "Failed to connect to EnergyGauge first")
                    return
                    
            discovery_data = self.automation.com_interface.discover_com_interface()
            
            # Display discovery results
            self.discovery_text.delete(1.0, tk.END)
            self.discovery_text.insert(tk.END, "=== EnergyGauge COM Interface Discovery ===\n\n")
            self.discovery_text.insert(tk.END, f"Version: {discovery_data.get('com_interface_version', 'Unknown')}\n\n")
            
            self.discovery_text.insert(tk.END, "Application Properties:\n")
            for prop in discovery_data.get('application_properties', []):
                self.discovery_text.insert(tk.END, f"  • {prop}\n")
                
            self.discovery_text.insert(tk.END, "\nApplication Methods:\n")
            for method in discovery_data.get('application_methods', []):
                self.discovery_text.insert(tk.END, f"  • {method}()\n")
                
            self.discovery_text.insert(tk.END, "\n\n=== USAGE INSTRUCTIONS ===\n")
            self.discovery_text.insert(tk.END, "Use these discovered properties and methods to customize\n")
            self.discovery_text.insert(tk.END, "the COM interface integration in the source code.\n")
            self.discovery_text.insert(tk.END, "Replace TODO comments with actual property/method calls.\n")
                
            self.status_var.set("COM interface discovery completed")
            messagebox.showinfo("Discovery Complete", "COM interface discovery completed. Check the Settings tab for detailed results.")
            
        except Exception as e:
            messagebox.showerror("Discovery Error", f"Failed to discover COM interface: {e}")
            
    def open_output_folder(self):
        """Open the output folder"""
        try:
            output_path = self.automation.project_root / "outputs"
            if output_path.exists():
                # Windows-specific command
                os.startfile(str(output_path))
            else:
                messagebox.showinfo("No Output", "No output folder found. Process some files first.")
        except Exception as e:
            messagebox.showerror("Error", f"Could not open output folder: {e}")
            
    def generate_report(self):
        """Generate processing report"""
        try:
            # Create a simple report
            report_data = {
                "generated": datetime.now().isoformat(),
                "uploaded_files": len(self.uploaded_files),
                "json_files": len([f for f in self.uploaded_files if f[1] == "json"]),
                "schematic_files": len([f for f in self.uploaded_files if f[1] == "schematic"]),
                "project_info": {
                    "name": self.project_name_var.get(),
                    "address": self.address_var.get(),
                    "city": self.city_var.get(),
                    "state": self.state_var.get(),
                    "zip": self.zip_var.get()
                }
            }
            
            # Save report
            report_file = self.automation.project_root / "reports" / f"processing_report_{datetime.now().strftime('%Y%m%d_%H%M%S')}.json"
            report_file.parent.mkdir(exist_ok=True)
            
            with open(report_file, 'w') as f:
                json.dump(report_data, f, indent=2)
                
            messagebox.showinfo("Report Generated", f"Processing report saved to:\n{report_file}")
            
        except Exception as e:
            messagebox.showerror("Report Error", f"Failed to generate report: {e}")
        
    def refresh_results(self):
        """Refresh the results display"""
        # Clear existing results
        for item in self.results_tree.get_children():
            self.results_tree.delete(item)
            
        # Load results from archive/outputs
        try:
            outputs_path = self.automation.project_root / "outputs"
            if outputs_path.exists():
                for project_dir in outputs_path.iterdir():
                    if project_dir.is_dir():
                        # Get project info
                        project_name = project_dir.name
                        mod_time = datetime.fromtimestamp(project_dir.stat().st_mtime)
                        date_str = mod_time.strftime("%Y-%m-%d %H:%M")
                        
                        # Count lots (subdirectories or files)
                        lot_count = len(list(project_dir.iterdir()))
                        
                        # Estimate duration (placeholder)
                        duration = "00:15:30"
                        
                        self.results_tree.insert('', 'end', text=project_name, 
                                               values=(date_str, 'Completed', lot_count, duration))
        except Exception as e:
            print(f"Error refreshing results: {e}")
            
    def run(self):
        """Run the GUI application"""
        # Initialize with some data
        self.refresh_results()
        self.update_processing_queue()
        
        # Start the GUI
        self.root.mainloop()


def main():
    """Main entry point"""
    parser = argparse.ArgumentParser(description='EnergyGauge USA Automation Suite')
    parser.add_argument('--mode', default='modern', choices=['gui', 'batch', 'modern'],
                       help='Run mode: modern (default), gui, or batch')
    parser.add_argument('--discover', action='store_true',
                       help='Discover COM interface properties and methods')
    parser.add_argument('--create-sample', action='store_true',
                       help='Create a sample input file')
    parser.add_argument('--project-root', type=str,
                       help='Project root directory (default: parent of script directory)')
    
    args = parser.parse_args()
    
    if args.create_sample:
        create_sample_input()
        return
        
    # Initialize automation
    automation = EnergyGaugeAutomation(args.project_root)
    
    if args.discover:
        print("Discovering EnergyGauge COM interface...")
        try:
            if automation.com_interface.connect(visible=True):
                discovery_data = automation.com_interface.discover_com_interface()
                print("\n=== EnergyGauge COM Interface Discovery ===")
                print(f"Version: {discovery_data.get('com_interface_version', 'Unknown')}")
                
                print("\nApplication Properties:")
                for prop in discovery_data.get('application_properties', []):
                    print(f"  • {prop}")
                    
                print("\nApplication Methods:")
                for method in discovery_data.get('application_methods', []):
                    print(f"  • {method}()")
                    
                print("\n=== Instructions ===")
                print("Use these discovered properties and methods to customize")
                print("the COM interface integration by replacing TODO comments")
                print("in the source code with actual property/method calls.")
                    
                automation.com_interface.disconnect()
            else:
                print("Failed to connect to EnergyGauge for discovery")
                print("Make sure EnergyGauge USA is installed and try running as Administrator")
        except Exception as e:
            print(f"Discovery failed: {e}")
        return
    
    try:
        if args.mode == 'modern':
            # Launch modern GUI
            print("Starting EnergyGauge Automation Suite...")
            gui = ModernEnergyGaugeGUI(automation)
            gui.run()
        elif args.mode == 'gui':
            automation.run_gui_mode()
        else:
            automation.run_batch_mode()
    except KeyboardInterrupt:
        print("\nProcessing interrupted by user")
    except Exception as e:
        print(f"Error: {e}")
        traceback.print_exc()
    finally:
        automation.cleanup()


if __name__ == "__main__":
    main()