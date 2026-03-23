#!/usr/bin/env python3
"""
Run Script for Professional Depreciation Calculator
==================================================

This script provides an easy way to generate the depreciation calculator.
Run this instead of calling the main script directly.

Usage:
    python run_calculator.py

Or make it executable and run:
    chmod +x run_calculator.py
    ./run_calculator.py
"""

import sys
import os
from pathlib import Path

def main():
    """Main execution function"""
    print("🇮🇳 Professional Depreciation Calculator Runner")
    print("=" * 55)

    # Check if main script exists
    main_script = Path("depreciation_calculator_pro.py")
    if not main_script.exists():
        print("❌ Error: Main script 'depreciation_calculator_pro.py' not found!")
        sys.exit(1)

    # Check Python version
    if sys.version_info < (3, 8):
        print("❌ Error: Python 3.8 or higher required!")
        print(f"   Current version: {sys.version}")
        sys.exit(1)

    # Check if required packages are installed
    try:
        import pandas
        import openpyxl
        print("✅ Dependencies verified")
    except ImportError as e:
        print("❌ Missing dependency:")
        print(f"   {e}")
        print("   Run: pip install -r requirements.txt")
        sys.exit(1)

    # Run the main script
    print("🚀 Generating depreciation calculator...")
    print()

    # Import and run the main function
    try:
        from depreciation_calculator_pro import main as calculator_main
        result = calculator_main()

        print()
        print("🎉 Calculator generated successfully!")
        print(f"📁 Output: {result}")

        # Check if file exists
        if os.path.exists(result):
            file_size = os.path.getsize(result)
            print(f"📊 File size: {file_size:,} bytes")
        else:
            print("⚠️  Warning: Output file not found!")

    except Exception as e:
        print(f"❌ Error running calculator: {e}")
        sys.exit(1)

if __name__ == "__main__":
    main()