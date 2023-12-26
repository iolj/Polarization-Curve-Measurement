import pyvisa
import time
from datetime import date
import numpy as np
import pandas as pd

rm = pyvisa.ResourceManager()
print(rm.list_resources())
PWR_01 = rm.open_resource('USB0::0x0B3E::0x1049::CY001179::0::INSTR')
print("Make sure to open the working folder and close Excel before starting!")
input("Press enter to continue...")
PWR_01.write('output on')
PWR_01.write('CURR 1.0')
activationTime = 60
voltageLimit = 2.0
print("Activating...")
time.sleep(activationTime)
current = [0.1, 0.2, 0.3, 0.4, 0.5, 1.0, 1.5, 2.0, 2.5, 3.0, 3.5, 4.0, 4.5, 5.0, 5.5, 6.0, 6.5, 7.0, 7.5, 8.0, 8.5, 9.0, 9.5, 10.0, 
           10.5, 11.0, 11.5, 12.0, 12.5, 13.0, 13.5, 14.0, 14.5, 15.0, 17.5, 20.0, 22.5, 25.0, 27.5, 30.0, 32.5, 35.0, 37.5, 40.0]
voltage = np.array(0)
intervalTime = 60 # unit: s
PWR_01.write('CURR 0.01')
input("When voltage is stable, press enter to continue...")
voltage = np.append(voltage, float(PWR_01.query('MEASure:VOLTage?')))
print(date.today(), time.strftime("%H:%M:%S", time.localtime()))
print(f'0A {str(voltage[voltage.size - 1])}V')
for i in range (len(current)):
    measured_voltage = float(PWR_01.query('MEASure:VOLTage?'))
    if (measured_voltage < voltageLimit):
        PWR_01.write(f'CURR {str(current[i])}')
        time.sleep(intervalTime)
        measured_voltage = float(PWR_01.query('MEASure:VOLTage?'))
        voltage = np.append(voltage, measured_voltage)
        print(date.today(), time.strftime("%H:%M:%S", time.localtime()))
        print(f'{str(current[i])}A {str(voltage[voltage.size - 1])}V')
    else:
        PWR_01.write('CURR 0')
        PWR_01.write('output off')
        print('Exceed voltage limit, ending the program...')
        break
voltage = np.trim_zeros(voltage)
print(voltage)
df = pd.DataFrame(voltage)
excel_file = 'output.xlsx'
input("Make sure to close Excel!!! Press enter to continue...")
df.to_excel(excel_file, index=False, header=False)