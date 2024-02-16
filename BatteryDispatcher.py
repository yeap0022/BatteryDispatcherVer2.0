# -*- coding: utf-8 -*-
"""
Created by Yeap Yew Ming
"""

import pandas as pd
import datetime
import os
import timeit
from ortools.linear_solver import pywraplp

# Read Excel file
file_name = os.fsdecode("input/input_file.xlsx")
workbook = pd.ExcelFile(file_name)

output_folder = "output"
if not os.path.isdir(output_folder):
    os.makedirs(output_folder)

# Load timeseries data
marketDF = workbook.parse("Timeseries data")
marketDF = marketDF.iloc[:,:5]
marketDF.columns = ["time", "market_price_1", "load", "wind", "solar"]
marketDF = marketDF[~pd.isnull(marketDF["time"])].fillna(0)

market1DF = marketDF.copy()
market1DF.sort_values(by=["time"], inplace=True)
market1DF["time_string"] = market1DF.apply(lambda x:(x["time"]+ datetime.timedelta(seconds=0.002)).strftime("%d/%m/%Y %H:%M"), axis=1)
market1DF.set_index("time_string", inplace=True)
marketDF = market1DF

# Load static data
gridDF = workbook.parse("Grid")
gridDF = gridDF.iloc[:,:4]
gridDF.columns = ["max_buy_power", "max_sell_power", "max_import_power", "max_export_power"]

# Load battery data
battDF = workbook.parse("Battery")
battDF = battDF.iloc[:,:8]
battDF.columns = ["max_charge_rate", "max_discharge_rate", "capacity", "charge_eff", "discharge_eff", "min_soc", "max_soc", "initial_soc"]

# Convert dataframe to dictionary
marketDict = marketDF.to_dict()
gridDict = gridDF.to_dict()
battDict = battDF.to_dict()

timeInterval = marketDF.iloc[1]["time"] - marketDF.iloc[0]["time"]

# Assign the data to right places
input = type("input", (dict,), {})()
input.update({
    "simData": {
        "startTime": datetime.datetime.strptime(marketDF.index[0], "%d/%m/%Y %H:%M"),
        "dt": int(round(timeInterval.total_seconds())) / (60 * 60), #in hour
        "tIndex": marketDF.shape[0]
        },
    "market": {
        key: {
            sub_key: sub_item for sub_key, sub_item in marketDict[key].items()
            } for key in marketDict.keys() if key != "time"
        },
    "grid": {
        key: item[0] for key, item in gridDict.items()
        },
    "batt": {
        key: item[0] for key, item in battDict.items()
        }
    })


# Create the mip solver with the CBC backend.
solver = pywraplp.Solver.CreateSolver("CBC")

inf = solver.infinity()

tIndex = input["simData"]["tIndex"] # number of timeslots
dt = input["simData"]["dt"] # time interval in hour

# Create datetime array
startTime = input["simData"]["startTime"].strftime("%d/%m/%Y %H:%M")
tIndex = input["simData"]["tIndex"]
timestamp = pd.date_range(startTime, periods=tIndex, freq=str(dt * 60) + "min")
time = [timestamp[i].strftime("%d/%m/%Y %H:%M") for i in range(len(timestamp))]

time_s = timeit.default_timer()
# Add timeseries variables
vGrid = [solver.NumVar(lb=-inf, ub=inf, name="") for _ in range(tIndex)]

vBattPower = [solver.NumVar(lb=-inf, ub=inf, name="") for _ in range(tIndex)]
vCharge = [solver.NumVar(lb=-inf, ub=0, name="") for _ in range(tIndex)]
vDischarge = [solver.NumVar(lb=0, ub=inf, name="") for _ in range(tIndex)]
vChargeStatus = [solver.BoolVar(name="") for _ in range(tIndex)]
vSOC = [solver.NumVar(lb=0, ub=1, name="") for _ in range(tIndex)]

# Add constraints
for i in range(tIndex):
    
    t = time[i]
    
    # Grid constraints
    solver.Add(vGrid[i] == input["market"]["load"][t] - input["market"]["solar"][t] - input["market"]["wind"][t] - vBattPower[i]) # Eqn. 1
    solver.Add(vGrid[i] <= input["grid"]["max_buy_power"]) # Eqn. 2
    solver.Add(vGrid[i] >= -input["grid"]["max_sell_power"]) # Eqn. 2
    solver.Add(input["market"]["load"][t] - input["market"]["solar"][t] - input["market"]["wind"][t] - (vDischarge[i] + vCharge[i]) <= input["grid"]["max_import_power"]) # Eqn. 3
    solver.Add(input["market"]["load"][t] - input["market"]["solar"][t] - input["market"]["wind"][t] - (vDischarge[i] + vCharge[i]) >= -input["grid"]["max_export_power"]) # Eqn. 3
    
    # Battery constraints
    solver.Add(vBattPower[i] == vCharge[i] + vDischarge[i]) # Eqn. 4
    solver.Add(vCharge[i] >= -input["batt"]["max_charge_rate"] * vChargeStatus[i]) # Eqn. 5(a)
    solver.Add(vDischarge[i] <= input["batt"]["max_discharge_rate"] * (1-vChargeStatus[i])) # Eqn. 5(b)
    
    if i == 0:
        solver.Add(vSOC[i] == input["batt"]["initial_soc"] - dt / input["batt"]["capacity"] * (vCharge[i] * (1-input["batt"]["charge_eff"]) + vDischarge[i] / (1-input["batt"]["discharge_eff"]))) # Eqn. 6
    else:
        solver.Add(vSOC[i] == vSOC[i-1] - dt / input["batt"]["capacity"] * (vCharge[i] * (1-input["batt"]["charge_eff"]) + vDischarge[i] / (1-input["batt"]["discharge_eff"]))) # Eqn. 6
        
    solver.Add(vSOC[i] >= input["batt"]["min_soc"]) # Eqn. 7
    solver.Add(vSOC[i] <= input["batt"]["max_soc"]) # Eqn. 7
    
# Add objective
obj = 0
obj += sum([vGrid[i] * input["market"]["market_price_1"][time[i]] * dt for i in range(tIndex)])
solver.Minimize(obj)

status = solver.Solve()

time_e = timeit.default_timer()
runTime = round(time_e - time_s, 4)

if status == solver.OPTIMAL or status == solver.FEASIBLE:
    print("Solution is found.")
    print("Number of variables =", solver.NumVariables())
    print("Number of constraints =", solver.NumConstraints())
    print("Computation time = ", runTime)
    
    # Extract solution values
    excelWriter = pd.ExcelWriter('output/Result.xlsx', engine='xlsxwriter')
    
    objValue = round(solver.Objective().Value() / 100, 2)
    
    objValueDF = pd.DataFrame.from_dict({"obj_value": objValue}, orient="index", columns=["Total Cost of Importing Power ($)"])
    
    result = list(zip([round(vGrid[i].solution_value(), 4) for i in range(tIndex)], 
                      [round(vBattPower[i].solution_value(), 4) for i in range(tIndex)],
                      [round(vCharge[i].solution_value(), 4) for i in range(tIndex)],
                      [round(vDischarge[i].solution_value(), 4) for i in range(tIndex)],
                      [round(vSOC[i].solution_value(), 4) for i in range(tIndex)],
                      [int(vChargeStatus[i].solution_value()) for i in range(tIndex)]
                      ))
    resultDF = pd.DataFrame(result, index=timestamp, columns=["Grid Power Flow (kW)", "Battery Output (kW)", "Charging Power (kW)", "Discharging Power (kW)", "State-of-charge (SOC)", "Charge Status"])
    
    objValueDF.to_excel(excelWriter, sheet_name='Cost')
    resultDF.to_excel(excelWriter, sheet_name='Operation')
    excelWriter.close()
else:
    print("Solution cannot be found.")
