"""
This file is reserved for testing of the data input functions and code.
"""
import pandas
import numpy
Bus_Loads = pandas.read_excel('C:/Users/JakeLaptop/Documents/GitHub/454_Project/Bus_Loads.xlsx')
Active_Power_Prod = pandas.read_excel('C:/Users/JakeLaptop/Documents/GitHub/454_Project/Active_Power_Production.xlsx')
PV_Bus_Ref = pandas.read_excel('C:/Users/JakeLaptop/Documents/GitHub/454_Project/PV_Bus_Reference_Voltages.xlsx')

print('Bus Loads: ' + str(Bus_Loads))
print('Active Power Produced: ' + str(Active_Power_Prod))
print('PV Bus Reference Voltages: ' + str(PV_Bus_Ref))