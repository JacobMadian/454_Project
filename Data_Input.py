"""
This file is reserved for testing of the data input functions and code.
"""
"""
Make sure to download an install openpyxl
It's the module the professor reccomended for excel interactions
"""
import openpyxl
from openpyxl import Workbook

"Hello this is new"

"Initializing new Workbook"
wb = Workbook()

"This uses the unique path on my PC to the data file, it will change for other users"
wb1 = openpyxl.load_workbook('C:/Users/JakeLaptop/Documents/GitHub/454_Project/Data/Bus_Loads.xlsx')
cells = wb1.active['A1':'K2']

"Need to find a more efficient way to do this."
for c1, c2, c3, c4, c5, c6, c7, c8, c9, c10, c11 in cells:
    print("{0:2} {1:2} {2:2} {3:2} {4:2} {5:2} {6:2} {7:2} {8:2} {9:2} {10:2}".format(c1.value, c2.value, c3.value, c4.value, c5.value, c6.value, c7.value, c8.value, c9.value, c10.value, c11.value))