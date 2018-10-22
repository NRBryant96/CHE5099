# -*- coding: utf-8 -*-
"""
Created on Fri Oct 19 11:45:18 2018
@author: Nathan
@author: pingl
"""

import xlwings as xw
import math
import pandas as pd

file = r'C:\Users\Nathan\Desktop\Senior Fall 2018\Computational Methods\Project 2\specsheet.xls'
wb_costing = xw.Book()
sht_cost = wb_costing.sheets['Sheet1']
wb = xw.Book(file)
plant_type = input("What plant type are you designing? (solid, solid/fluid, fluid)  ")

# Dictionaries for cost report
dictHX = {
        "Heat Exchanger": [],
        "Area" : [],
        "Tube Length" : [],
        "A" : [],
        "B" : [],
        "C" : [],
        "Cost" : [],
        }

dictHEAT = {
        "Fired Heater": [],
        "Heat Absorbed" : [],
        "Q" : [],
        "Cost" : [],
        }

dictREA = {
        "Reactor": [],
        "Orientation" : [],
        "Diameter" : [],
        "Length" : [],
        "Design Pressure" : [],
        "Wall Thickness" : [],
        "Weight" : [],
        "Platform Cost" : [],
        "Reactor Cost" : [],
        }

dictPUMP = {
        "Heat Exchanger": [],
        "Area" : [],
        "Tube Length" : [],
        "A" : [],
        "B" : [],
        "C" : [],
        "Cost" : [],
        }
dictFLAS = {
        "Heat Exchanger": [],
        "Area" : [],
        "Tube Length" : [],
        "A" : [],
        "B" : [],
        "C" : [],
        "Cost" : [],
        }
dictCOMP = {
        "Compressor": [],
        "Actual Power (hp) " : [],
        "Cost" : [],
        }
dictTOWR = {
        "Heat Exchanger": [],
        "Area" : [],
        "Tube Length" : [],
        "A" : [],
        "B" : [],
        "C" : [],
        "Cost" : [],
        }

count = 0
#Standard Equations and Factors
def Cb_Heat_Ex_Pump(S,A,B,C):
    Cb = math.e**(A - B * math.log(S) + C * math.log(S)**2)
    return(Cb)
def Fm_Heat_Exchangers(Area,a,b):
    Fm = a + (Area/100)**b
    return(Fm)
def Fp_Heat_Exchangers(Pressure,Alpha,Bravo,Charlie):
    Fp = Alpha + Bravo*(Pressure/100) + Charlie*(Pressure/100)**2
    return(Fp)
def Get_Lang_Factors(plant_type):
    if plant_type == "solid":
        lang = 3.97
    elif plant_type == "solid/fluid":
        lang = 4.28
    else:
        lang = 5.04
    return(lang)



for x in range(len(wb.sheets)):
    sht = wb.sheets[x]
    count += 1
    
    if "HTXR" in sht.name:
        check_list = sht.range('AA1:AA100').value
        values_list = sht.range('AD1:AD100').value
        check_list_htxr = sht.range('AI1:AI100').value
        stream1_list = sht.range('AK1:AK100').value
        stream2_list = sht.range('AL1:AL100').value
        stream3_list = sht.range('AM1:AM100').value
        stream4_list = sht.range('AN1:AN100').value
        combined_list_htxr = list(zip(check_list,values_list,check_list_htxr,stream1_list,stream2_list,stream3_list,stream4_list))
        tube_length = input("What is the tube length for " +str(sht.name) + "(ex. 20)  ")
        if tube_length == "4":
            Fl = 1.47
        elif tube_length == "6":
            Fl = 1.34
        elif tube_length == "8":
            Fl = 1.25
        elif tube_length == "10":
            Fl = 1.18
        elif tube_length == "12":
            Fl = 1.12
        elif tube_length == "16":
            Fl = 1.05
        elif tube_length == "20":
            Fl = 1.00
        else:
            Fl = 1.2
        for x in range(1,len(combined_list_htxr)):
            if combined_list_htxr[x][0] == r'Area/shell':
                Area = combined_list_htxr[x][1]
                if Area is None:
                    Area = 500
            elif combined_list_htxr[x][0] == r'Exchanger type':
                ex_type = combined_list_htxr[x][1]
            elif combined_list_htxr[x][0] == r'Cost model':
                cost_model = combined_list_htxr[x][1]
            elif combined_list_htxr[x][0] == r'Shell and tube':
                material = combined_list_htxr[x][1]
            elif combined_list_htxr[x][2] == r'pressure':
                first_stream_in_pressure = combined_list_htxr[x][3]
                first_stream_out_pressure = combined_list_htxr[x][5]
                second_stream_in_pressure = combined_list_htxr[x][4]
                second_stream_out_pressure = combined_list_htxr[x][6]
        if first_stream_in_pressure is None:
            first_stream_in_pressure = 15
        if first_stream_out_pressure is None:
            first_stream_out_pressure = 15
        if second_stream_in_pressure is None:
            second_stream_in_pressure = 15
        if second_stream_out_pressure is None:
            second_stream_out_pressure = 15
        if first_stream_in_pressure >= first_stream_out_pressure:
            first_stream_pressure = first_stream_in_pressure - 14.7
        else:
            first_stream_pressure = first_stream_out_pressure - 14.7
        if second_stream_in_pressure >= second_stream_out_pressure:
            second_stream_pressure = second_stream_in_pressure - 14.7
        else:
            second_stream_pressure = second_stream_out_pressure - 14.7
        if first_stream_pressure >= second_stream_pressure:
            Pressure = first_stream_pressure
        else:
            Pressure = second_stream_pressure
        if ex_type is None:
            A = 11.4185
            B = 0.9228
            C = 0.0986
            Alpha = 0.9803
            Bravo = 0.018
            Charlie = 0.0017
        elif ex_type == 1:
            A = 12.331
            B = 0.8709
            C = 0.09005
            Alpha = 0.851
            Bravo = 0.1292
            Charlie = 0.0198
        elif ex_type == 2:
            A = 11.551
            B = 0.9186
            C = 0.0979
            Alpha = 0.9803
            Bravo = 0.018
            Charlie = 0.0017
        else:
            A = 11.4185
            B = 0.9228
            C = 0.0986
            Alpha = 0.9803
            Bravo = 0.018
            Charlie = 0.0017
        
#        elif "Floating Head"
#            A = 12.031
#            B = 0.8709
#            C = 0.09005
        if material is None:
            a = 0
            b = 0
        elif material == 1 or 2 or 3:
            a = 2.7
            b = 0.07
#        elif material == 4:
            #Unknown
#            1
        elif material == 5:
            a = 3.3
            b = 0.08
#        elif material == 6 or 7:
            #unknown
#            1
        elif material == 8:
            a = 9.6
            b = 0.06
        else:
            a = 0
            b = 0
        Fm = Fm_Heat_Exchangers(Area,a,b)
        if cost_model == 4:
            A = 7.2718
            B = -0.16
            C = 0
            if material == 1 or 2 or 3:
                Fm = 3
            else:
                Fm = 2
        Cb_htxr = Cb_Heat_Ex_Pump(Area,A,B,C)
        Fp = Fp_Heat_Exchangers(Pressure,Alpha,Bravo,Charlie)
        
        if cost_model == 4:
            cost_htxr = Cb_htxr*Fp*Fm
        else:
            cost_htxr = Cb_htxr*Fl*Fp*Fm
        sht_cost.range('A'+str(count)).value = cost_htxr
        dictHX['Area'].append(Area)
        dictHX['A'].append(A)
        dictHX['B'].append(B)
        dictHX['C'].append(C)
        dictHX['Tube Length'].append(Fl)
        dictHX["Heat Exchanger"].append(sht.name)
        dictHX['Cost'].append(cost_htxr)
   
    elif "FIRE" in sht.name:
        A = -0.15241
        B = -0.785
        C = 0
        check_list = sht.range('AA1:AA100').value
        values_list = sht.range('AD1:AD100').value
        combined_list = list(zip(check_list,values_list))
        for x in range(1,len(combined_list)):
            if combined_list[x][0] == r'Heat Absorbed':
                Q = combined_list[x][1]
        
    elif "REA" in sht.name:
        dimmension = input("Is the vessel vertical or horizontal? (vertical, horizontal))  ")
        check_list = sht.range('AA1:AA100').value
        values_list = sht.range('AD1:AD100').value
        check_list_reac = sht.range('AI1:AI100').value
        combined_list_reac = list(zip(check_list,values_list))
        for x in range(1,len(combined_list_reac)):
            if combined_list_reac[x][0] == r'Reactor volume':
                volume = combined_list_reac[x][1]
            elif combined_list_reac[x][0] == r'Pressure In':
                pressure = combined_list_reac[x][1]
                if pressure is None:
                    pressure = 15
            elif combined_list_reac[x][0] == r'Tout':
                temperature = combined_list_reac[x][1]
        diameter = int((4 * volume/(10 * math.pi))**(1/3))
        length = volume / ((math.pi / 4) * diameter**2)
        design_pressure = math.e**(0.60608 + 0.91615 * math.log(pressure) + 0.0015655 * math.log(pressure)**2)
        if temperature is None or (temperature >= -20 and temperature < 750):
            S = 15000
        elif temperature >= 750 and temperature < 800:
            S = 14750
        elif temperature >= 800 and temperature < 850:
            S = 14200
        elif temperature >= 850:
            S = 13100
        wall_thickness = int(((design_pressure * diameter * 12 / (2 * S * 0.85 - 1.2 * design_pressure))+(1/8))*16)/16
        weight = math.pi*(diameter + wall_thickness/12) * (length + 0.8 * diameter)* wall_thickness/12 * 490
        cost_vessel = math.e**(10.5449 - 0.4672 * math.log(weight) + 0.05482 * math.log(weight)**2)
        if dimmension == "vertical":
            cost_platform = 410 * diameter**0.73960 * length**0.70684
        elif dimmension == "horizontal":
            cost_platform = 2275 * diameter**0.2094
        sht_cost.range('A'+str(count)).value = cost_vessel
        sht_cost.range('B'+str(count)).value = cost_platform
    
    elif "PUMP" in sht.name:
        check_list = sht.range('AA1:AA100').value
        values_list = sht.range('AD1:AD100').value
        pump_check_list = sht.range('AI1:AI100').value
        pump_values_list1 = sht.range('AK1:AK100').value
        pump_values_list2 = sht.range('AL1:AL100').value
        combined_list_pump = list(zip(check_list,values_list,pump_check_list,pump_values_list1,pump_values_list2))
        for x in range(1,len(combined_list)):
            if combined_list_pump[x][0] == r'Pump type':
                pump = combined_list_pump[x][1]
            elif combined_list_pump[x][0] == r'Centrifugal pump':
                centrifugal_pump = combined_list_pump[x][1]
            elif combined_list_pump[x][0] == r'Material':
                material = combined_list_pump[x][1]
            elif combined_list_pump[x][0] == r'Motor type':
                motor_type = combined_list_pump[x][1]
            elif combined_list_pump[x][2] == r'pressure':
                pressure_change = combined_list_pump[x][4] - combined_list_pump[x][3]
            elif combined_list_pump[x][0] == r'Vol. flow rate':
                Q = combined_list_pump[x][1]
            elif combined_list_pump[x][0] == r'Head':
                head = combined_list_pump[x][1]
            elif combined_list_pump[x][0] == r'Calculated power':
                Pt = combined_list_pump[x][1]
            elif combined_list_pump[x][0] == r'Motor RPM':
                motor_RPM = combined_list_pump[x][1]
            
        Np = -0.316 + 0.24015*math.log(Q) - 0.01199 * math.log(Q)**2
        Pb = Pt/Np
        Nm = 0.8 + 0.0319 * math.log(Pb) - 0.00182 * math.log(Pb)**2
        Pc = int(4*Pb/(Nm+1))/4
        if centrifugal_pump is None:
            Ft = 1
        elif centrifugal_pump == 1:
            Ft = 1.5
        elif centrifugal_pump == 2:
            Ft = 1.7
        elif centrifugal_pump == 3:
            Ft = 2
        elif centrifugal_pump == 4:
            Ft = 2.7
        elif centrifugal_pump == 5:
            Ft = 8.9
        if pump is None:
            S = Q * head**0.5
            A = 12.1656
            B = 1.1448
            C = 0.0862
            if material is None:
                Fm = 1
            elif material == 1:
                Fm = 1.35
            elif material == 12:
                Fm = 1.9
            elif material == 3:
                Fm = 2
            elif material == 10:
                Fm = 2.95
            elif material == 6:
                Fm = 3.3
            elif material == 5:
                Fm = 3.5
            elif material == 9:
                Fm = 9.7
        elif pump == 1:
            S = Pb
            A = 7.9361
            B = -0.26986
            C = 0.06718
            if material is None:
                1
            elif material == 11:
                Fm = 1
            elif material == 5 or 12:
                Fm = 1.15
            elif material == 1:
                Fm = 1.5
            elif material == 3:
                Fm = 2.2
        elif pump == 2:
            S = Q
            A = 8.2816
            B = 0.2918
            C = 0.0743
            if material is None:
                Fm = 1
            elif material is 1:
                Fm = 1.35
            elif material == 12:
                Fm = 1.9
            elif material == 3:
                Fm = 2
            elif material == 10:
                Fm = 2.95
            elif material == 6:
                Fm = 3.3
            elif material == 5:
                Fm = 3.5
            elif material == 9:
                Fm = 9.7
        ### for motors
        if motor_RPM is None:
            if motor_type is None:
                Ft = 1
            elif motor_type == 1:
                Ft = 1.4
            else:
                Ft = 1.8
        else:
            if motor_type is None:
                Ft = 0.9
            elif motor_type == 1:
                Ft = 1.3
            else:
                Ft = 1.7
        Cb_pump = Cb_Heat_Ex_Pump(S,A,B,C)
        Cb_motor = math.e**(5.9332 + 0.16829 * math.log(Pc) - 0.110056 
                            * math.log(Pc)**2 + 0.071413 
                            * math.log(Pc)**3 - 0.0063788 
                            * math.log(Pc)**4)
        
        if pump is None or pump == 1:
            cost_pump = Ft * Fm * Cb_pump
        elif pump == 2:
            cost_pump = Ft * Cb_pump
        cost_motor = Cb_motor * Ft
        
        sht_cost.range('A'+str(count)).value = cost_pump
        sht_cost.range('B'+str(count)).value = cost_motor
        
    elif "FLAS" in sht.name:
        1
        
    elif "COMP" in sht.name:
        check_list = sht.range('AA1:AA100').value
        values_list = sht.range('AD1:AD100').value
        comp_check_list = sht.range('AI1:AI100').value
        comp_values_list1 = sht.range('AK1:AK100').value
        comp_values_list2 = sht.range('AL1:AL100').value
        combined_list_comp = list(zip(check_list,values_list,
                                      comp_check_list,comp_values_list1,
                                      comp_values_list2))
        
        for x in range(1,len(combined_list)):
            if combined_list_comp[x][0] == r'Type of compressor':
                comp_type = combined_list_comp[x][1]
            
            elif combined_list_comp[x][0] == r'Actual power':
                comp_power_con = combined_list_comp[x][1]
            
            #elif combined_list_pump[x][0] == r'Material':
                #material = combined_list_pump[x][1]
            
            #elif combined_list_pump[x][0] == r'Motor type':
                #motor_type = combined_list_pump[x][1]
        # Cb cost calculation based on compressor type and power consumption        
        if comp_type ==  1:
            cb = math.e**(9.1553 + 0.63 * math.log(comp_power_con))
        elif comp_type == 2:
            cb = math.e**(4.6762 + 1.23 * math.log(comp_power_con))
        elif comp_type == 3:
            cb = math.e**(8.2496 + 0.78243 * math.log(comp_power_con))
        # Motor Drive type
        Fd = 1 # electric motor
        # Fd = 1.15 for steam
        # Fd = 1.25 for gas
        Fm = 1 # cast iron, carbon steel
        # Fm = 2.5 for stainless steel
        # Fm = 5 for nickel alloy
        cost_compressor = cb * Fd * Fm   
            
    elif "TOWR" in sht.name:
        1

df = pd.DataFrame(dictHX)
writer = pd.ExcelWriter('PythonExport.xlsx')
df.to_excel(writer,'HXcost')
writer.save()