import os
import pandas as pd
from paraview.simple import *
from paraview.simple import SaveScreenshot

def load_data(file_path):
    """Load the data into ParaView."""
    reader = OpenDataFile(file_path)
    return reader

def main():
    input_file = "C:/Users/SW-048/Documents/2_空力/Training/ParaView/CFD_RawData/Clean_Rev1_M1_A20-fluid.vtu"
    
    # Load Data
    data = load_data(input_file)

if __name__ == "__main__":
    main()
