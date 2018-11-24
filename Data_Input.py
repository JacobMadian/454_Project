"""
This file is reserved for testing of the data input functions and code.
"""
"""
Make sure to download an install openpyxl
It's the module the professor reccomended for excel interactions
"""
import openpyxl as xl
from openpyxl import Workbook

"Initializing new Workbook"
wb = Workbook()

"This uses the unique path on my PC to the data file, it will change for other users"
Load_Book = xl.load_workbook('C:/Users/JakeLaptop/Documents/GitHub/454_Project/Data/Bus_Loads.xlsx').active
Generator_Book = xl.load_workbook('C:/Users/JakeLaptop/Documents/GitHub/454_Project/Data/Active_Power_Production.xlsx').active
PV_Book = xl.load_workbook('C:/Users/JakeLaptop/Documents/GitHub/454_Project/Data/PV_Bus_Reference_Voltages.xlsx').active

P_List = []
Q_List = []
Generator_List = []
Generator_Power = []
Reference_Voltage = []

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

print("P_List: " + str(P_List))
print("Q_List: " + str(Q_List))
print("PV Busses are buss numbers: " + str(Generator_List))
print("Those generators produce this much power(MW): " + str(Generator_Power))
print("PV Bus reference voltages: " + str(Reference_Voltage))

