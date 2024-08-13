import os
import pandas as pd
from openpyxl import Workbook
from openpyxl.drawing.image import Image
from paraview.simple import *
from paraview.simple import SaveScreenshot

def load_data(file_path):
    """Load the data into ParaView."""
    reader = OpenDataFile(file_path)
    if reader:
    	print("Success")
    else:
    	print("Failed")

    return reader

def main():
    input_file = "C:/Users/SW-048/Desktop/CFD_RawData/Clean_Rev1_M1_A20-fluid.vtu"
    
    # Load Data
    data = load_data(input_file)

if __name__ == "__main__":
    main()
