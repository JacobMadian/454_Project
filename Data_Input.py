"""
This file is reserved for testing of the data input functions and code.
"""
"""
Make sure to download an install openpyxl
It's the module the professor reccomended for excel interactions
"""
import openpyxl as xl
from openpyxl import Workbook
import numpy as np

"Initializing new Workbook"
wb = Workbook()

"This uses the unique path on my PC to the data file, it will change for other users"
Load_Book = xl.load_workbook('C:/Users/JakeLaptop/Documents/GitHub/454_Project/Data/Bus_Loads.xlsx').active
Generator_Book = xl.load_workbook('C:/Users/JakeLaptop/Documents/GitHub/454_Project/Data/Active_Power_Production.xlsx').active
PV_Book = xl.load_workbook('C:/Users/JakeLaptop/Documents/GitHub/454_Project/Data/PV_Bus_Reference_Voltages.xlsx').active
Line_Book = xl.load_workbook('C:/Users/JakeLaptop/Documents/GitHub/454_Project/Data/Line_Data.xlsx').active

P_List = []
Q_List = []
Generator_List = []
Generator_Power = []
Reference_Voltage = []
Admittance = np.zeros((14,14), dtype = object)

for r in Load_Book['1']:
    P_List.append(r.value)
for r in Load_Book['2']:
    Q_List.append(r.value)
for r in Generator_Book['1']:
    Generator_List.append(r.value)
for r in Generator_Book['2']:
    Generator_Power.append(r.value)
for r in PV_Book['2']:
    Reference_Voltage.append(r.value)

for r in Line_Book['A']:
    for count in range (0,20):
        if r.value == count:
            value = Line_Book['B'+str(r.row)].value
            admit = Line_Book['D'+str(r.row)].value
            Admittance[count - 1,value - 1] = (1/float(admit))
            Admittance[count - 1, count - 1] = (Admittance[count - 1, count - 1] - (1/float(admit)))
print(Admittance)

"""
print("P_List: " + str(P_List))
print("Q_List: " + str(Q_List))
print("PV Busses are buss numbers: " + str(Generator_List))
print("Those generators produce this much power(MW): " + str(Generator_Power))
print("PV Bus reference voltages: " + str(Reference_Voltage))

"""