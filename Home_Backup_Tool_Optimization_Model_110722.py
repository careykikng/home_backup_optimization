## Pyomo optimization model

# import packages
from __future__ import division
from pyomo.environ import *
from pyomo.opt import SolverFactory
import pandas as pd
import numpy as np
import json
import xlwings as xw


def button(): # function called by Excel VBA macro

    # get user inputs from Excel workbook
    comp_tool_xl_wb = xw.Book('Computational_Tool_Home_Backup_Optimization_110722.xlsm')

    # name variables for sheets
    user_interface = comp_tool_xl_wb.sheets['User Interface']
    calc_values_sheet = comp_tool_xl_wb.sheets['Calculation Values']
    optim_results_sheet = comp_tool_xl_wb.sheets['Optimization Results']
    actual_load_profiles_sheet = comp_tool_xl_wb.sheets['Actual Home Load Profiles']
    PV_profiles_sheet = comp_tool_xl_wb.sheets['PV Profiles']
    HVAC_app_sheet = comp_tool_xl_wb.sheets['HVAC Appliances']
    preferred_app_timings_sheet = comp_tool_xl_wb.sheets['Preferred Appliance Timings']
    preferred_backup_timings_sheet = comp_tool_xl_wb.sheets['Preferred Backup Source Timings']

    # save backup source input data as dataframe
    backup_source_config_entered = pd.DataFrame(user_interface.range('C14:F42').value)
    backup_source_config_entered[0] = backup_source_config_entered[0].str.replace(" ", "_")
    backup_source_config_entered = backup_source_config_entered.fillna(0)

    # save appliances input data as dataframe
    appliances_running_inputs = pd.DataFrame(user_interface.range('C50:H115').value)
    appliances_running_inputs[0] = appliances_running_inputs[0].str.replace(" ", "_")
    appliances_running_inputs = appliances_running_inputs.fillna(0)

    # define the time steps
    number_of_days = user_interface.range('C127').value
    number_of_hours = number_of_days*24
    time_step_interval = 0.25 # in terms of hours, so 0.25 is 15 minute intervals
    number_of_time_steps = int(number_of_hours/time_step_interval)
    number_of_time_steps_one_day = int(24/time_step_interval)

    start_time = user_interface.range('E127').value

    # find the backup sources entered by the user
    count_backup_sources = 0
    count_backup_sources_batt = 0
    count_backup_sources_gen = 0
    count_backup_PV = 0
    count_backup_BEV = 0
    count_backup_PHEV = 0
    names_backup_sources_all = []
    priority_backup_all = []
    priority_backup_batt_charge = []
    backup_kW_all = []
    backup_surge_kW_all = []
    backup_kWh_all = []
    backup_min_SoC_all = []
    backup_max_SoC_all = []
    backup_initial_SoC_all = []
    backup_c_d_eff_all = []
    total_batt_power_cap = 0
    preferred_backup_timings_day = pd.DataFrame()

    if backup_source_config_entered.loc[0,2] > 0: # battery electric vehicle
        count_backup_sources = count_backup_sources + 1
        count_backup_sources_batt = count_backup_sources_batt + 1
        count_backup_BEV = count_backup_BEV + 1

        BEV_name = 'Battery_Electric_Vehicle'
        names_backup_sources_all.append(BEV_name)

        BEV_kW = backup_source_config_entered.loc[0,2]
        backup_kW_all.append(BEV_kW)

        BEV_surge_kW = BEV_kW
        backup_surge_kW_all.append(BEV_surge_kW)

        BEV_kWh = backup_source_config_entered.loc[1,2]
        backup_kWh_all.append(BEV_kWh)

        BEV_min_SoC = backup_source_config_entered.loc[2,2]
        backup_min_SoC_all.append(BEV_min_SoC)

        BEV_max_SoC = backup_source_config_entered.loc[3,2]
        backup_max_SoC_all.append(BEV_max_SoC)

        BEV_initial_SoC = backup_source_config_entered.loc[4,2]
        backup_initial_SoC_all.append(BEV_initial_SoC)

        BEV_c_d_eff = backup_source_config_entered.loc[5,2]
        backup_c_d_eff_all.append(BEV_c_d_eff)

        BEV_priority_c = 0.1* user_interface.range('I15').value
        priority_backup_batt_charge.append(BEV_priority_c)

        BEV_priority_d = -0.1 * user_interface.range('I14').value
        priority_backup_all.append(BEV_priority_d)

        total_batt_power_cap = total_batt_power_cap + BEV_kW

        preferred_BEV_timings_day = pd.DataFrame(preferred_backup_timings_sheet.range('D4:D' + str(3+number_of_time_steps_one_day)).value)
        preferred_backup_timings_day = pd.concat([preferred_backup_timings_day, preferred_BEV_timings_day], axis=1, ignore_index=True)
        

    if backup_source_config_entered.loc[6,2] > 0: # plug-in hybrid electric vehicle battery
        count_backup_sources = count_backup_sources + 1
        count_backup_sources_batt = count_backup_sources_batt + 1
        count_backup_PHEV = count_backup_PHEV + 1

        PHEV_name = 'Plug-In_Hybrid_Electric_Vehicle'
        names_backup_sources_all.append(PHEV_name)

        PHEV_kW = backup_source_config_entered.loc[6,2]
        backup_kW_all.append(PHEV_kW)

        PHEV_surge_kW = PHEV_kW
        backup_surge_kW_all.append(PHEV_surge_kW)

        PHEV_battery_kWh = backup_source_config_entered.loc[7,2]
        backup_kWh_all.append(PHEV_battery_kWh)

        PHEV_min_SoC = backup_source_config_entered.loc[8,2]
        backup_min_SoC_all.append(PHEV_min_SoC)

        PHEV_max_SoC = backup_source_config_entered.loc[9,2]
        backup_max_SoC_all.append(PHEV_max_SoC)

        PHEV_initial_SoC = backup_source_config_entered.loc[10,2]
        backup_initial_SoC_all.append(PHEV_initial_SoC)

        PHEV_c_d_eff = backup_source_config_entered.loc[11,2]
        backup_c_d_eff_all.append(PHEV_c_d_eff)

        PHEV_priority_c = 0.1 * user_interface.range('I17').value
        priority_backup_batt_charge.append(PHEV_priority_c)

        PHEV_priority_d = -0.1 * user_interface.range('I16').value
        priority_backup_all.append(PHEV_priority_d)

        total_batt_power_cap = total_batt_power_cap + PHEV_kW

        preferred_PHEV_timings_day = pd.DataFrame(preferred_backup_timings_sheet.range('E4:E' + str(3+number_of_time_steps_one_day)).value)
        preferred_backup_timings_day = pd.concat([preferred_backup_timings_day, preferred_PHEV_timings_day], axis=1, ignore_index=True)
  

    if backup_source_config_entered.loc[15,2] > 0: # behind the meter storage
        count_backup_sources = count_backup_sources + 1
        count_backup_sources_batt = count_backup_sources_batt + 1

        BTM_name = 'Behind-the-Meter_Storage'
        names_backup_sources_all.append(BTM_name)

        BTM_kW = backup_source_config_entered.loc[15,2]
        backup_kW_all.append(BTM_kW)

        BTM_surge_kW = BTM_kW
        backup_surge_kW_all.append(BTM_surge_kW)

        BTM_kWh = backup_source_config_entered.loc[16,2]
        backup_kWh_all.append(BTM_kWh)

        BTM_min_SoC = backup_source_config_entered.loc[17,2]
        backup_min_SoC_all.append(BTM_min_SoC)

        BTM_max_SoC = backup_source_config_entered.loc[18,2]
        backup_max_SoC_all.append(BTM_max_SoC)

        BTM_initial_SoC = backup_source_config_entered.loc[19,2]
        backup_initial_SoC_all.append(BTM_initial_SoC)

        BTM_c_d_eff = backup_source_config_entered.loc[20,2]
        backup_c_d_eff_all.append(BTM_c_d_eff)

        BTM_priority_c = 0.1 * user_interface.range('I21').value
        priority_backup_batt_charge.append(BTM_priority_c)

        BTM_priority_d = -0.1 * user_interface.range('I20').value
        priority_backup_all.append(BTM_priority_d)

        total_batt_power_cap = total_batt_power_cap + BTM_kW

        preferred_BTM_timings_day = pd.DataFrame(preferred_backup_timings_sheet.range('F4:F' + str(3+number_of_time_steps_one_day)).value)
        preferred_backup_timings_day = pd.concat([preferred_backup_timings_day, preferred_BTM_timings_day], axis=1, ignore_index=True)
        

    if backup_source_config_entered.loc[6,2] > 0: # plug-in hybrid electric vehicle ICE
        count_backup_sources = count_backup_sources + 1
        count_backup_sources_gen = count_backup_sources_gen + 1

        PHEV_name = 'Plug-In_Hybrid_Electric_Vehicle'
        names_backup_sources_all.append(PHEV_name)

        PHEV_kW = backup_source_config_entered.loc[6,2]
        backup_kW_all.append(PHEV_kW)

        PHEV_surge_kW = PHEV_kW
        backup_surge_kW_all.append(PHEV_surge_kW)

        PHEV_fuel_kWh = backup_source_config_entered.loc[13,3]
        backup_kWh_all.append(PHEV_fuel_kWh)

        PHEV_priority = -0.1 * user_interface.range('I18').value
        priority_backup_all.append(PHEV_priority)

        preferred_PHEV_timings_day = pd.DataFrame(preferred_backup_timings_sheet.range('E4:E' + str(3+number_of_time_steps_one_day)).value)
        preferred_backup_timings_day = pd.concat([preferred_backup_timings_day, preferred_PHEV_timings_day], axis=1, ignore_index=True)
        

    if backup_source_config_entered.loc[21,2] > 0: # portable generator
        count_backup_sources = count_backup_sources + 1
        count_backup_sources_gen = count_backup_sources_gen + 1

        port_gen_name = 'Portable_Generator'
        names_backup_sources_all.append(port_gen_name)

        port_gen_kW = backup_source_config_entered.loc[21,2]
        backup_kW_all.append(port_gen_kW)

        port_gen_surge_kW = backup_source_config_entered.loc[22,2]
        backup_surge_kW_all.append(port_gen_surge_kW)

        port_gen_kWh = backup_source_config_entered.loc[24,3]
        backup_kWh_all.append(port_gen_kWh)

        port_gen_priority = -0.1 * user_interface.range('I22').value
        priority_backup_all.append(port_gen_priority)

        preferred_port_gen_timings_day = pd.DataFrame(preferred_backup_timings_sheet.range('G4:G' + str(3+number_of_time_steps_one_day)).value)
        preferred_backup_timings_day = pd.concat([preferred_backup_timings_day, preferred_port_gen_timings_day], axis=1, ignore_index=True)
        

    if backup_source_config_entered.loc[25,2] > 0: # standby generator
        count_backup_sources = count_backup_sources + 1
        count_backup_sources_gen = count_backup_sources_gen + 1

        standby_name = 'Standby_Generator'
        names_backup_sources_all.append(standby_name)

        standby_kW = backup_source_config_entered.loc[25,2]
        backup_kW_all.append(standby_kW)

        standby_surge_kW = backup_source_config_entered.loc[26,2]
        backup_surge_kW_all.append(standby_surge_kW)

        standby_kWh = backup_source_config_entered.loc[28,3]
        if backup_source_config_entered.loc[27,2] == 'Natural Gas':
            standby_kWh = 1000000000
        backup_kWh_all.append(standby_kWh)

        standby_priority = -0.1 * user_interface.range('I23').value
        priority_backup_all.append(standby_priority)

        preferred_standby_timings_day = pd.DataFrame(preferred_backup_timings_sheet.range('H4:H' + str(3+number_of_time_steps_one_day)).value)
        preferred_backup_timings_day = pd.concat([preferred_backup_timings_day, preferred_standby_timings_day], axis=1, ignore_index=True)
        

    if backup_source_config_entered.loc[14,2] > 0: # photovoltaic generation
        count_backup_PV = 1 
        PV_kW = backup_source_config_entered.loc[14,2]

        PV_profile_day = pd.DataFrame(PV_profiles_sheet.range('B9:B104').value, index=range(number_of_time_steps_one_day), columns=range(1))
        PV_profile_day[PV_profile_day > total_batt_power_cap*1000] = total_batt_power_cap*1000
        PV_profile_week = pd.DataFrame()
        for r in range(8):
            PV_profile_week = pd.concat([PV_profile_week, PV_profile_day], ignore_index=True)
        
        PV_profile = pd.DataFrame(PV_profile_week.iloc[int(start_time*number_of_time_steps_one_day):int(start_time*number_of_time_steps_one_day)+number_of_time_steps])
        PV_profile = PV_profile.reset_index(drop=True)
        PV_profile.to_csv('PV_profile_data.csv')
        
        PV_priority = 1 - (user_interface.range('I19').value/100)
    else:
        PV_profile = pd.DataFrame(0, index=range(number_of_time_steps), columns=range(1))
        PV_profile.to_csv('PV_profile_data.csv')



    # find the appliances running selected by the user
    names_appliances_running_all = []
    priority_appliances_all = []
    hours_running_per_day_all = []
    running_watts_all = []
    add_starting_watts_all = []
    number_temp_dep_appliances_selected = 0
    names_temp_dep_appliances_running_all = []
    priority_temp_dep_appliances_all = []
    temp_dep_running_watts_all = pd.DataFrame()
    temp_dep_add_starting_watts_all = []

    for i, row in appliances_running_inputs.iterrows():
        if (row[1]) > 0:
            if row[0] == 'Central_air_conditioner_(temperature_dependent)' or row[0] == 'Gas_furnace_(temperature_dependent)' or row[0] == 'Air_source_heat_pump_(temperature_dependent)':
                if row[0] == 'Central_air_conditioner_(temperature_dependent)':
                    number_temp_dep_appliances_selected = number_temp_dep_appliances_selected + 1
                    names_temp_dep_appliances_running_all.append(row[0])
                    priority_temp_dep_appliances_all.append(row[2])
                    temp_dep_add_starting_watts_all.append(row[5]*row[1])
                    temp_dep_AC_running_watts_all = pd.DataFrame(HVAC_app_sheet.range('B6:B101').value, index=range(number_of_time_steps_one_day), columns=range(1))
                    temp_dep_running_watts_all = pd.concat([temp_dep_running_watts_all, temp_dep_AC_running_watts_all], axis=1, ignore_index=True)
                
                if row[0] == 'Gas_furnace_(temperature_dependent)':
                    number_temp_dep_appliances_selected = number_temp_dep_appliances_selected + 1
                    names_temp_dep_appliances_running_all.append(row[0])
                    priority_temp_dep_appliances_all.append(row[2])
                    temp_dep_add_starting_watts_all.append(row[5]*row[1])
                    temp_dep_GF_running_watts_all = pd.DataFrame(HVAC_app_sheet.range('C6:C101').value, index=range(number_of_time_steps_one_day), columns=range(1))
                    temp_dep_running_watts_all = pd.concat([temp_dep_running_watts_all, temp_dep_GF_running_watts_all], axis=1, ignore_index=True) 
                
                if row[0] == 'Air_source_heat_pump_(temperature_dependent)':
                    number_temp_dep_appliances_selected = number_temp_dep_appliances_selected + 1
                    names_temp_dep_appliances_running_all.append(row[0])
                    priority_temp_dep_appliances_all.append(row[2])
                    temp_dep_add_starting_watts_all.append(row[5]*row[1])
                    temp_dep_HP_running_watts_all = pd.DataFrame(HVAC_app_sheet.range('D6:D101').value, index=range(number_of_time_steps_one_day), columns=range(1))
                    temp_dep_running_watts_all = pd.concat([temp_dep_running_watts_all, temp_dep_HP_running_watts_all], axis=1, ignore_index=True)
                
                temp_dep_running_watts_all.columns = names_temp_dep_appliances_running_all 
            
    for i, row in appliances_running_inputs.iterrows():
        if (row[1]) > 0:  
            if row[0] != 'Central_air_conditioner_(temperature_dependent)' and row[0] != 'Gas_furnace_(temperature_dependent)' and row[0] != 'Air_source_heat_pump_(temperature_dependent)': 
                names_appliances_running = row[0]
                names_appliances_running_all.append(names_appliances_running)
                quantity_running = row[1]
                priority_appliances = row[2]
                priority_appliances_all.append(priority_appliances)
                hours_running_per_day = row[3]
                hours_running_per_day_all.append(hours_running_per_day)
                running_watts = row[4]*quantity_running
                running_watts_all.append(running_watts)
                add_starting_watts = row[5]*quantity_running
                add_starting_watts_all.append(add_starting_watts)

    number_appliances_selected = len(names_appliances_running_all)


    # if actual home load profile is selected
    count_actual_load_profile = 0
    if str(actual_load_profiles_sheet.range('B6').value) == 'True':
        # remove any inputs for single appliances
        number_appliances_selected = 0
        number_temp_dep_appliances_selected = 0
        count_actual_load_profile = 1
        # save the load profile as a dataframe
        actual_home_load_profile_day = pd.DataFrame(actual_load_profiles_sheet.range('B12:B107').value, index=range(number_of_time_steps_one_day), columns=range(1))
        actual_home_load_profile_week = pd.DataFrame()
        for r in range(8):
            actual_home_load_profile_week = pd.concat([actual_home_load_profile_week, actual_home_load_profile_day], ignore_index=True)
        
        actual_home_load_profile = pd.DataFrame(actual_home_load_profile_week.iloc[int(start_time*number_of_time_steps_one_day):int(start_time*number_of_time_steps_one_day)+number_of_time_steps])
        actual_home_load_profile = actual_home_load_profile.reset_index(drop=True)
        # save the starting watts
        actual_home_add_starting_watts = actual_load_profiles_sheet.range('B113').value


    # save preferred timings of appliances if appliances are selected
    if number_appliances_selected > 0:
        preferred_app_timings_all = pd.DataFrame(preferred_app_timings_sheet.range('D3:DQ' + str(3+number_of_time_steps_one_day)).value)
        preferred_app_timings_all.columns = preferred_app_timings_all.iloc[0]
        preferred_app_timings_all = preferred_app_timings_all.drop([0])
        preferred_app_timings_all.columns = preferred_app_timings_all.columns.str.replace(" ", "_")
        
        preferred_app_timings_day = pd.DataFrame()
        preferred_app_timings_week = pd.DataFrame()
        # find the right timings for each appliance
        for m in range(number_appliances_selected):
            name_of_appliance = names_appliances_running_all[m]
            preferred_app_timings_appliances = preferred_app_timings_all[name_of_appliance]
            preferred_app_timings_day = pd.concat([preferred_app_timings_day, preferred_app_timings_appliances], axis=1, ignore_index=True)
        
        for r in range(8):
            preferred_app_timings_week = pd.concat([preferred_app_timings_week, preferred_app_timings_day], ignore_index=True)

        preferred_app_timings = pd.DataFrame(preferred_app_timings_week.iloc[int(start_time*number_of_time_steps_one_day):int(start_time*number_of_time_steps_one_day)+number_of_time_steps])
        preferred_app_timings = preferred_app_timings.reset_index(drop=True)
        preferred_app_timings.columns = names_appliances_running_all


    # save preferred timings of backup sources
    preferred_backup_timings_week = pd.DataFrame()
    for r in range(8):
        preferred_backup_timings_week = pd.concat([preferred_backup_timings_week, preferred_backup_timings_day], ignore_index=True)

    preferred_backup_timings = pd.DataFrame(preferred_backup_timings_week.iloc[int(start_time*number_of_time_steps_one_day):int(start_time*number_of_time_steps_one_day)+number_of_time_steps])
    preferred_backup_timings = preferred_backup_timings.reset_index(drop=True)
    preferred_backup_timings.columns = names_backup_sources_all


    # making the matrix and vector parameters, A b and c
    len_i = number_of_time_steps + count_backup_sources_gen + number_appliances_selected # the number time steps plus an additional row per backup source and appliance selected
        # i = number of constraints
    len_j = number_of_time_steps*(count_backup_sources + count_backup_PV + count_actual_load_profile + number_temp_dep_appliances_selected + number_appliances_selected) # the number of time steps multiplied by the number of backup sources and appliances selected
        # j = number of variables

    # make the A matrix, size i by j
    df_A = pd.DataFrame()
    # get the diagonal of backup source capacity
    for n in range(count_backup_sources):
        df_A_n = np.zeros((number_of_time_steps, number_of_time_steps))
        np.fill_diagonal(df_A_n, backup_kW_all[n]*1000)
        df_A_n = pd.DataFrame(df_A_n)
        df_A = pd.concat([df_A, df_A_n], axis=1, ignore_index=True)  
    # get the diagonal of the PV profile if PV is selected
    if count_backup_PV > 0:
        df_A_pv = np.zeros((number_of_time_steps, number_of_time_steps))
        np.fill_diagonal(df_A_pv, PV_profile)
        df_A_pv = pd.DataFrame(df_A_pv)
        df_A = pd.concat([df_A, df_A_pv], axis=1, ignore_index=True)
    # get the diagonal of the negative of the actual home load profile if it is selected
    if count_actual_load_profile > 0:
        df_A_actual_load = np.zeros((number_of_time_steps, number_of_time_steps))
        np.fill_diagonal(df_A_actual_load, actual_home_load_profile * -1)
        df_A_actual_load = pd.DataFrame(df_A_actual_load)
        df_A = pd.concat([df_A, df_A_actual_load], axis=1, ignore_index=True)
    # get the diagonal of the negative of the running watts of the temperature dependent appliances
    for d in range(number_temp_dep_appliances_selected):
        df_A_d = np.zeros((number_of_time_steps, number_of_time_steps))
        np.fill_diagonal(df_A_d, temp_dep_running_watts_all.iloc[:,d] * -1)
        df_A_d = pd.DataFrame(df_A_d)
        df_A = pd.concat([df_A, df_A_d], axis=1, ignore_index=True)  
    # get the diagonal of the negative of the running watts of appliances
    for m in range(number_appliances_selected):
        df_A_m = np.zeros((number_of_time_steps, number_of_time_steps))
        np.fill_diagonal(df_A_m, running_watts_all[m]*-1)
        df_A_m = pd.DataFrame(df_A_m)
        df_A = pd.concat([df_A, df_A_m], axis=1, ignore_index=True)
    # add rows for power capacity of backup sources in units of [W-15min]
    for n in range(count_backup_sources_gen):
        row_A_n = pd.DataFrame(0, index=range(1), columns=range(len(df_A.columns)))
        row_A_n.loc[0,(number_of_time_steps*(n + count_backup_sources_batt)):((number_of_time_steps*(n + count_backup_sources_batt + 1))-1)] = backup_kW_all[n + count_backup_sources_batt]*1000*time_step_interval
        df_A = pd.concat([df_A, row_A_n], ignore_index=True)
    # add rows for running watts of appliances in units of [W-15min]
    for m in range(number_appliances_selected):
        row_A_m = pd.DataFrame(0, index=range(1), columns=range(len(df_A.columns)))
        row_A_m.loc[0,(count_backup_sources + count_backup_PV + number_temp_dep_appliances_selected + m)*number_of_time_steps:((count_backup_sources + count_backup_PV + number_temp_dep_appliances_selected + m + 1)*number_of_time_steps)-1] = running_watts_all[m]*time_step_interval
        df_A = pd.concat([df_A, row_A_m], ignore_index=True)
    # make a csv file for A
    df_A.to_csv('df_A_data.csv')

    # make the vector b, size i
    df_b = pd.DataFrame(0, index=range(number_of_time_steps), columns=range(1))
    # add maximum kWh generation for each backup source
    for n in range(count_backup_sources_gen):
        row_b_n = pd.DataFrame(0, index=range(1), columns=range(1))
        row_b_n.loc[0,0] = backup_kWh_all[n + count_backup_sources_batt]*1000
        df_b = pd.concat([df_b, row_b_n], ignore_index=True)
    # add daily target energy in Wh for each appliance
    for m in range(number_appliances_selected):
        row_b_m = pd.DataFrame(0, index=range(1), columns=range(1))
        row_b_m.loc[0,0] = hours_running_per_day_all[m]*running_watts_all[m]*number_of_days
        df_b = pd.concat([df_b, row_b_m], ignore_index=True)
    # rename column to 'b'
    df_b = df_b.rename(columns={0:"b"})
    # make a csv file for b
    df_b.to_csv('df_b_data.csv', index_label='i')

    # make the vector c, size j
    df_c = pd.DataFrame()
    # get the rows for backup sources to be equal to the source priority and preferred timings
    for n in range(count_backup_sources):
        df_c_n = pd.DataFrame(preferred_backup_timings.iloc[:,n]*priority_backup_all[n])
        df_c_n.columns= range(df_c_n.shape[1])
        df_c = pd.concat([df_c, df_c_n], ignore_index=True)
    # get rows for PV priority if PV is selected
    if count_backup_PV > 0:
        df_c_pv = pd.DataFrame(PV_priority, index=range(number_of_time_steps), columns=range(1))
        df_c = pd.concat([df_c, df_c_pv], ignore_index=True)
    # make rows for actual home load priority if it is selected
    if count_actual_load_profile > 0:
        df_c_actual_load = pd.DataFrame(10, index=range(number_of_time_steps), columns=range(1))
        df_c = pd.concat([df_c, df_c_actual_load], ignore_index=True)    
    # get the rows for temperature dependent appliances to be equal to the appliance priority
    df_c_d = pd.DataFrame(1, index=range(number_of_time_steps*number_temp_dep_appliances_selected), columns=range(1))
    for d in range(number_temp_dep_appliances_selected):
        df_c_d = pd.DataFrame(priority_temp_dep_appliances_all[d], index=range(number_of_time_steps), columns=range(1))
        df_c_d.columns= range(df_c_d.shape[1])
        df_c = pd.concat([df_c, df_c_d], ignore_index=True)
    # get the rows for appliances to be equal to the appliance priority and preferred timings
    df_c_m = pd.DataFrame(1, index=range(number_of_time_steps*number_appliances_selected), columns=range(1))
    for m in range(number_appliances_selected):
        df_c_m = pd.DataFrame(preferred_app_timings.iloc[:,m]*priority_appliances_all[m])
        df_c_m.columns= range(df_c_m.shape[1])
        df_c = pd.concat([df_c, df_c_m], ignore_index=True)
    # rename column to 'c'
    df_c = df_c.rename(columns={0:"c"})
    # make a csv file for c
    df_c.to_csv('df_c_data.csv', index_label='j')

    # make an equivalent 'c' vector for only battery charging (priority*preferred timings)
    df_c_batt = pd.DataFrame()
    for n in range(count_backup_sources_batt):
        df_c_batt_c = pd.DataFrame(preferred_backup_timings.iloc[:,n]*priority_backup_batt_charge[n])
        df_c_batt_c.columns= range(df_c_batt_c.shape[1])
        df_c_batt = pd.concat([df_c_batt, df_c_batt_c], ignore_index=True)
    if count_backup_sources_batt < 1:
        df_c_batt = pd.DataFrame(0, index=range(number_of_time_steps), columns=range(1))
    df_c_batt.to_csv('df_c_batt_data.csv')

    # delete previous data from A b and c sheets & print A, b, and c onto respective Excel sheets 
    comp_tool_xl_wb.sheets['A'].clear()
    comp_tool_xl_wb.sheets['A']["A1"].options(pd.DataFrame, header=1, index=True, expand='table').value = df_A
    comp_tool_xl_wb.sheets['b'].clear()
    comp_tool_xl_wb.sheets['b']["A1"].options(pd.DataFrame, header=1, index=True, expand='table').value = df_b
    comp_tool_xl_wb.sheets['c'].clear()
    comp_tool_xl_wb.sheets['c']["A1"].options(pd.DataFrame, header=1, index=True, expand='table').value = df_c







    # define the optimization model
    model = AbstractModel(name = 'optimization abstract model with objective of maximixing appliances running')

    # define sets
    model.i = Set(initialize = [i for i in range(len_i)], ordered=True) # i = number of constraints
    model.j = Set(initialize = [j for j in range(len_j)], ordered=True) # j = number of variables
    model.t = Set(initialize = [t for t in range(number_of_time_steps)], ordered=True) # t = number of time steps
    model.t_batt = Set(initialize = [b for b in range(number_of_time_steps*count_backup_sources_batt)], ordered=True) # t_batt = number of time steps times number of battery backup sources
    model.num_batt = Set(initialize = [n for n in range(count_backup_sources_batt)], ordered=True) # num_batt = number of battery backup sources selected
    model.num_temp_dep_appliances = Set(initialize = [d for d in range(number_temp_dep_appliances_selected)], ordered=True) # num_temp_dep_appliances = number of temperature dependent appliances selected
    model.num_appliances = Set(initialize = [m for m in range(number_appliances_selected)], ordered=True) # num_appliances = number of appliances selected
    model.num_total_appliances = Set(initialize = [m for m in range(number_temp_dep_appliances_selected+number_appliances_selected)], ordered=True) # num_total_appliances = total number of appliances selected
    model.num_actual_home_load_profile = Set(initialize = [p for p in range(count_actual_load_profile)], ordered=True) # num_actual_home_load_profile = number of actual home load profile (there can be 1 or none)
    model.num_PHEV = Set(initialize = [p for p in range(count_backup_PHEV)], ordered=True) # num_PHEV = number of plug-in hybrid electric vehicles (there can be 1 or none)
    model.num_PV = Set(initialize = [v for v in range(count_backup_PV)], ordered=True) # num_PV = number of photovoltaic generation systems (there can be 1 or none)

    # define the parameters
    model.A = Param(model.i, model.j)
    model.b = Param(model.i)
    model.c = Param(model.j)
    model.c_batt = Param(model.t_batt)
    model.PV_profile = Param(model.t)

    # load data into parameters
    data = DataPortal()
    data.load(filename = 'df_A_data.csv', param = model.A, format='array')
    data.load(filename = 'df_b_data.csv', select = ('i', 'b'), param = model.b, index = model.i)
    data.load(filename = 'df_c_data.csv', select = ('j', 'c'), param = model.c, index = model.j)
    if count_backup_sources_batt > 0:
        data.load(filename = 'df_c_batt_data.csv', param = model.c_batt)
    data.load(filename = 'PV_profile_data.csv', param = model.PV_profile)    

    # define the variables
    model.x = Var(model.j, bounds=(0, 1), domain=lambda m, j: NonNegativeReals if j <= (number_of_time_steps*(count_backup_sources + count_backup_PV + count_actual_load_profile)-1) else Binary)
    model.batt_charge_CF = Var(model.t, model.num_batt, bounds=(-1,0), domain = Reals)
    model.starting_watts_td_constr = Var(model.t, model.num_temp_dep_appliances)
    model.starting_watts_constr = Var(model.t, model.num_appliances)
    model.starting_watts_actual_home_constr = Var(model.t, model.num_actual_home_load_profile, bounds=(0,1), domain = Binary)
    model.battery_SoC = Var(model.t, model.num_batt, bounds=(0,1), domain = NonNegativeReals)
    model.batt_status = Var(model.t, model.num_batt, bounds=(0,1), domain = Binary)
    model.PHEV_gen_status = Var(model.t, model.num_PHEV, bounds=(0,1), domain = Binary)
    model.PV_to_batt = Var(model.t, bounds=(0,1), domain = NonNegativeReals)

    # define the objective
    def obj_expression(model):
        # return summation(model.c, model.x)
        return sum(model.c[j] * model.x[j] for j in model.j) + sum((1/model.c_batt[t_c + (number_of_time_steps*num_batt)]) * model.battery_SoC[t_c, num_batt] for t_c in model.t for num_batt in model.num_batt)
    model.OBJ = Objective(rule=obj_expression, sense=maximize)

    # define the constraints
    def Axb_constraint_rule(model, i):
        # supply = demand power constraint
        if i < number_of_time_steps: 
            return sum(model.A[i,j] * model.x[j] for j in model.j) == model.b[i]
        # target energy >= energy solved for each appliance constraint AND
        # electricity output for generator <= maximum kWh constraint
        if i >= number_of_time_steps:
            return sum(model.A[i,j] * model.x[j] for j in model.j) <= model.b[i]
    model.Axb_Constraint = Constraint(model.i, rule = Axb_constraint_rule)

    # excess PV constraint
    def excess_PV_constraint_rule(model, t):
        j_PV = number_of_time_steps*count_backup_sources
        return model.PV_to_batt[t] + model.x[t + j_PV]*count_backup_PV <= 1
    model.excess_PV_Constraint = Constraint(model.t, rule = excess_PV_constraint_rule)

    # charging batteries from excess PV constraint
    if count_backup_sources_batt > 0:
        def charging_batteries_constraint_rule(model, t):
            return sum(model.batt_charge_CF[t, n]*backup_kW_all[n]*1000 for n in model.num_batt) + model.PV_to_batt[t]*model.PV_profile[t] == 0
        model.charging_batteries_Constraint = Constraint(model.t, rule = charging_batteries_constraint_rule)
        
    # charge/discharge constraint, can't charge and discharge battery at same time
        def charging_constraint_rule(model, t, num_batt):
            return model.batt_status[t,num_batt] >= (-1 * model.batt_charge_CF[t, num_batt])
        model.charging_Constraint = Constraint(model.t, model.num_batt, rule = charging_constraint_rule)
        
        def discharging_constraint_rule(model, t, num_batt):
            return model.batt_status[t,num_batt] <= (-1 * model.x[(number_of_time_steps*num_batt) + t]) +1
        model.discharging_Constraint = Constraint(model.t, model.num_batt, rule = discharging_constraint_rule)

    # starting watts constraints
    if number_temp_dep_appliances_selected > 0:
        def starting_watts_td_constraint_rule(model, t, num_temp_dep_appliances):
            srt_watts_constr_start_td = number_of_time_steps*(count_backup_sources + count_backup_PV + num_temp_dep_appliances)
            return model.starting_watts_td_constr[t, num_temp_dep_appliances] == model.x[t + srt_watts_constr_start_td]
        model.starting_watts_td_Constraint = Constraint(model.t, model.num_temp_dep_appliances, rule = starting_watts_td_constraint_rule)
        
        def max_starting_watts_td_constraint_rule(model, t, num_temp_dep_appliances):     
            j_app = range(number_of_time_steps*(count_backup_sources + count_backup_PV), number_of_time_steps*(count_backup_sources + count_backup_PV + number_temp_dep_appliances_selected + number_appliances_selected))
            return (model.starting_watts_td_constr[t,num_temp_dep_appliances] * temp_dep_add_starting_watts_all[num_temp_dep_appliances]) - sum(model.A[t,j] * model.x[j] for j in j_app) <= sum(backup_surge_kW_all)*1000
        model.max_starting_watts_td_Constraint = Constraint(model.t, model.num_temp_dep_appliances, rule = max_starting_watts_td_constraint_rule)
    
    if number_appliances_selected > 0:
        def starting_watts_constraint_rule(model, t, num_appliances):
            srt_watts_constr_start = number_of_time_steps*(count_backup_sources + count_backup_PV + number_temp_dep_appliances_selected + num_appliances)
            if t == 0:
                return model.starting_watts_constr[t, num_appliances] == (model.x[srt_watts_constr_start] - model.x[srt_watts_constr_start + number_of_time_steps - 1])
            else:
                return model.starting_watts_constr[t, num_appliances] == (model.x[t + srt_watts_constr_start] - model.x[t + srt_watts_constr_start - 1])
        model.starting_watts_Constraint = Constraint(model.t, model.num_appliances, rule = starting_watts_constraint_rule)

        def max_starting_watts_constraint_rule(model, t, num_appliances):     
            j_app = range(number_of_time_steps*(count_backup_sources + count_backup_PV), number_of_time_steps*(count_backup_sources + count_backup_PV + number_temp_dep_appliances_selected + number_appliances_selected))
            return (model.starting_watts_constr[t,num_appliances] * add_starting_watts_all[num_appliances]) - sum(model.A[t,j] * model.x[j] for j in j_app) <= sum(backup_surge_kW_all)*1000
        model.max_starting_watts_Constraint = Constraint(model.t, model.num_appliances, rule = max_starting_watts_constraint_rule)
    
    if count_actual_load_profile > 0:
        def starting_watts_actual_home_constraint_rule(model, t, num_actual_home_load_profile):
            srt_watts_constr_start_actual_home = number_of_time_steps*(count_backup_sources + count_backup_PV + num_actual_home_load_profile)
            return model.starting_watts_actual_home_constr[t, num_actual_home_load_profile] >= model.x[t + srt_watts_constr_start_actual_home]
        model.starting_watts_actual_home_Constraint = Constraint(model.t, model.num_actual_home_load_profile, rule = starting_watts_actual_home_constraint_rule)
        
        def max_starting_watts_actual_home_constraint_rule(model, t, num_actual_home_load_profile):     
            j_app = range(number_of_time_steps*(count_backup_sources + count_backup_PV), number_of_time_steps*(count_backup_sources + count_backup_PV + count_actual_load_profile))
            return (model.starting_watts_actual_home_constr[t,num_actual_home_load_profile] * actual_home_add_starting_watts) - sum(model.A[t,j] * model.x[j] for j in j_app) <= sum(backup_surge_kW_all)*1000
        model.max_starting_watts_actual_home_Constraint = Constraint(model.t, model.num_actual_home_load_profile, rule = max_starting_watts_actual_home_constraint_rule)
    

    # PHEV constraints
    def PHEV_gen_constraint_rule(model, t, num_PHEV): # track when engine is on/off
        j_PHEV_gen_start = number_of_time_steps*count_backup_sources_batt
        return model.PHEV_gen_status[t, num_PHEV] >= model.x[t + j_PHEV_gen_start]
    model.PHEV_gen_Constraint = Constraint(model.t, model.num_PHEV, rule= PHEV_gen_constraint_rule)

    def PHEV_constraint_rule(model, t, num_PHEV): # can't use battery and engine at same time
        return model.PHEV_gen_status[t, num_PHEV] + (-1* model.batt_charge_CF[t, num_batt]) + model.x[((count_backup_BEV+num_PHEV) * number_of_time_steps) + t] <= 1
    model.PHEV_Constraint = Constraint(model.t, model.num_PHEV, rule= PHEV_constraint_rule)

    # SoC constraints
    if count_backup_sources_batt > 0:
        def SoC_constraint_rule(model, t, num_batt): # state of charge constraint
            if t == 0:
                return model.battery_SoC[t,num_batt] == backup_initial_SoC_all[num_batt] - (model.batt_charge_CF[t, num_batt] * backup_kW_all[num_batt] * time_step_interval * backup_c_d_eff_all[num_batt])/ backup_kWh_all[num_batt] - (model.x[(num_batt* number_of_time_steps) + t] * backup_kW_all[num_batt] * time_step_interval / backup_c_d_eff_all[num_batt])/ backup_kWh_all[num_batt]
            else:
                return model.battery_SoC[t,num_batt] == model.battery_SoC[t-1,num_batt] - (model.batt_charge_CF[t, num_batt] * backup_kW_all[num_batt] * time_step_interval * backup_c_d_eff_all[num_batt])/ backup_kWh_all[num_batt] - (model.x[(num_batt * number_of_time_steps) + t] * backup_kW_all[num_batt] * time_step_interval / backup_c_d_eff_all[num_batt])/ backup_kWh_all[num_batt]
        model.SoC_Constraint = Constraint(model.t, model.num_batt, rule = SoC_constraint_rule)

        def max_SoC_constraint_rule(model, t, num_batt): # maximum SoC constraint
            return model.battery_SoC[t,num_batt] <= backup_max_SoC_all[num_batt]
        model.max_SoC_Constraint = Constraint(model.t, model.num_batt, rule = max_SoC_constraint_rule)

        def min_SoC_constraint_rule(model, t, num_batt): # minimum SoC constraint
            return model.battery_SoC[t,num_batt] >= backup_min_SoC_all[num_batt]
        model.min_SoC_Constraint = Constraint(model.t, model.num_batt, rule = min_SoC_constraint_rule)
        

    # create instance of the model (abstract only)
    model = model.create_instance(data)


    # solve the model
    opt = SolverFactory('gurobi')
    opt.options['TimeLimit'] = 2
    opt.options['OptimalityTol'] = 0.01

    # to use glpk as the solver instead of gurobi, comment out lines 668-670, and uncomment lines 673-675
#    opt = SolverFactory('glpk')
#    opt.options['tmlim'] = 2
#    opt.options['mipgap'] = 0.01

    status = opt.solve(model)





    # write model outputs to a JSON file
    model.solutions.store_to(status)
    status.write(filename='optimized_timings_obj_max_results.json', format='json')

    # read results from model
    input_file = 'optimized_timings_obj_max_results.json'
    with open(input_file) as json_data:
        results_data = json.load(json_data)

    sol = results_data['Solution']
    sol_df = pd.DataFrame(sol)
    obj = sol_df['Objective']
    var = sol_df['Variable'][1]
    var2 = pd.DataFrame.from_dict(var)

    solver_res = results_data['Solver']
    solver_res_df = pd.DataFrame(solver_res)
    opt_model_time = solver_res_df['Time'] # time it took to build and solve the model
    user_interface.range("Q11").clear_contents()
    user_interface.range('Q11').options(index=False, header=False).value = opt_model_time

    # get variable 'x' results into dataframe
    x_results = []
    for r in range(len_j):
        if 'x[' + str(r) +']' in var2:
            loc_col_x = var2.columns.get_loc("x[" + str(r) +"]")
            x_results_loop = var2.iloc[0,loc_col_x]
        elif 'x[' + str(r) +']' not in var2:
            x_results_loop = 0
        x_results.append(x_results_loop)
    x_results = pd.DataFrame(x_results)

    # get variable 'batt_charge_CF' into dataframe
    batt_charge_CF_results = pd.DataFrame()
    for n in range(count_backup_sources_batt):
        for t in range(number_of_time_steps):
            if 'batt_charge_CF[' + str(t) + ',' + str(n) + ']' in var2:
                loc_col_batt_charge_CF = var2.columns.get_loc("batt_charge_CF[" + str(t) + "," + str(n) + "]")
                batt_charge_CF_results_loop = pd.DataFrame(var2.iloc[0,loc_col_batt_charge_CF], index=range(1), columns=range(1))
            elif 'batt_charge_CF[' + str(t) + ',' + str(n) + ']' not in var2:
                batt_charge_CF_results_loop = pd.DataFrame(0, index=range(1), columns=range(1))
            batt_charge_CF_results = pd.concat([batt_charge_CF_results, batt_charge_CF_results_loop], ignore_index=True)

    # get variable 'PV_to_batt' into datafram
    PV_to_batt_results = pd.DataFrame()
    for t in range(number_of_time_steps):
        if 'PV_to_batt[' + str(t) + ']' in var2:
            loc_col_PV_to_batt = var2.columns.get_loc("PV_to_batt[" + str(t) + "]")
            PV_to_batt_results_loop = pd.DataFrame(var2.iloc[0,loc_col_PV_to_batt], index=range(1), columns=range(1))
        elif 'PV_to_batt[' + str(t) + ']' not in var2:
            PV_to_batt_results_loop = pd.DataFrame(0, index=range(1), columns=range(1))
        PV_to_batt_results = pd.concat([PV_to_batt_results, PV_to_batt_results_loop], ignore_index=True)

    # get variable 'battery_SoC' into dataframe
    battery_SoC_results = pd.DataFrame()
    for n in range(count_backup_sources_batt):
        for t in range(number_of_time_steps):
            if 'battery_SoC[' + str(t) + ',' + str(n) + ']' in var2:
                loc_col_battery_SoC = var2.columns.get_loc("battery_SoC[" + str(t) + "," + str(n) + "]")
                battery_SoC_results_loop = pd.DataFrame(var2.iloc[0,loc_col_battery_SoC], index=range(1), columns=range(1))
            elif 'battery_SoC[' + str(t) + ',' + str(n) + ']' not in var2:
                battery_SoC_results_loop = pd.DataFrame(0, index=range(1), columns=range(1))
            battery_SoC_results = pd.concat([battery_SoC_results, battery_SoC_results_loop], ignore_index=True)

    # get variable 'starting_watts_constr' results into dataframe
    starting_watts_td_constr_results = pd.DataFrame()
    for m in range(number_temp_dep_appliances_selected):
        for t in range(number_of_time_steps):
            if 'starting_watts_td_constr[' + str(t) + ',' + str(m) + ']' in var2:
                loc_col_starting_watts_td_constr = var2.columns.get_loc("starting_watts_td_constr[" + str(t) + "," + str(m) + "]")
                starting_watts_td_constr_results_loop = pd.DataFrame(var2.iloc[0,loc_col_starting_watts_td_constr], index=range(1), columns=range(1))
            elif 'starting_watts_td_constr[' + str(t) + ',' + str(m) + ']' not in var2:
                starting_watts_td_constr_results_loop = pd.DataFrame(0, index=range(1), columns=range(1))
            starting_watts_td_constr_results = pd.concat([starting_watts_td_constr_results, starting_watts_td_constr_results_loop], ignore_index=True)
    
    starting_watts_constr_results = pd.DataFrame()
    for m in range(number_appliances_selected):
        for t in range(number_of_time_steps):
            if 'starting_watts_constr[' + str(t) + ',' + str(m) + ']' in var2:
                loc_col_starting_watts_constr = var2.columns.get_loc("starting_watts_constr[" + str(t) + "," + str(m) + "]")
                starting_watts_constr_results_loop = pd.DataFrame(var2.iloc[0,loc_col_starting_watts_constr], index=range(1), columns=range(1))
            elif 'starting_watts_constr[' + str(t) + ',' + str(m) + ']' not in var2:
                starting_watts_constr_results_loop = pd.DataFrame(0, index=range(1), columns=range(1))
            starting_watts_constr_results = pd.concat([starting_watts_constr_results, starting_watts_constr_results_loop], ignore_index=True)

    starting_watts_actual_home_constr_results = pd.DataFrame()
    for m in range(count_actual_load_profile):
        for t in range(number_of_time_steps):
            if 'starting_watts_actual_home_constr[' + str(t) + ',' + str(m) + ']' in var2:
                loc_col_starting_watts_actual_home_constr = var2.columns.get_loc("starting_watts_actual_home_constr[" + str(t) + "," + str(m) + "]")
                starting_watts_actual_home_constr_results_loop = pd.DataFrame(var2.iloc[0,loc_col_starting_watts_actual_home_constr], index=range(1), columns=range(1))
            elif 'starting_watts_actual_home_constr[' + str(t) + ',' + str(m) + ']' not in var2:
                starting_watts_actual_home_constr_results_loop = pd.DataFrame(0, index=range(1), columns=range(1))
            starting_watts_actual_home_constr_results = pd.concat([starting_watts_actual_home_constr_results, starting_watts_actual_home_constr_results_loop], ignore_index=True)
    


    # clear previous results from optimization results and calculation values sheets on Excel
    optim_results_sheet.range("D7:BY1000").clear_contents()
    optim_results_sheet.range("S6:BW6").clear_contents()
    calc_values_sheet.range("J7:J1000").clear_contents()
    calc_values_sheet.range("S6:BY1000").clear_contents()
    calc_values_sheet.range("CA6:EE1000").clear_contents()
    calc_values_sheet.range("S2:BW2").clear_contents()

    # print the capacity factor and SoC results to Excel
    backup_res_count = 0
    energy_start_all = []
    energy_end_all = []

    if backup_source_config_entered.loc[0,2] > 0: # battery electric vehicle
        backup_res_count = backup_res_count + 1
        charge_results = batt_charge_CF_results.iloc[(backup_res_count-1)*number_of_time_steps:backup_res_count*number_of_time_steps,0]      
        discharge_results = x_results.iloc[(backup_res_count-1)*number_of_time_steps:(backup_res_count*number_of_time_steps),0]
        charge_results = charge_results.reset_index(drop=True)
        discharge_results = discharge_results.reset_index(drop=True)
        batt_cap_factor_results = charge_results + discharge_results
        optim_results_sheet.range('D7').options(index=False, header=False).value = batt_cap_factor_results
        optim_results_sheet.range('L7').options(index=False, header=False).value = battery_SoC_results.iloc[(backup_res_count-1)*number_of_time_steps:backup_res_count*number_of_time_steps,0]
        energy_start_BEV = BEV_kWh*(BEV_initial_SoC - BEV_min_SoC)
        energy_start_all.append(energy_start_BEV)
        energy_end_BEV = BEV_kWh*(battery_SoC_results.iloc[(backup_res_count*number_of_time_steps)-1,0] - BEV_min_SoC)
        energy_end_all.append(energy_end_BEV)

    if backup_source_config_entered.loc[6,2] > 0: # plug-in hybrid electric vehicle battery
        backup_res_count = backup_res_count + 1
        charge_results = batt_charge_CF_results.iloc[(backup_res_count-1)*number_of_time_steps:backup_res_count*number_of_time_steps,0]      
        discharge_results = x_results.iloc[(backup_res_count-1)*number_of_time_steps:(backup_res_count*number_of_time_steps),0]
        charge_results = charge_results.reset_index(drop=True)
        discharge_results = discharge_results.reset_index(drop=True)
        batt_cap_factor_results = charge_results + discharge_results
        optim_results_sheet.range('E7').options(index=False, header=False).value = batt_cap_factor_results
        optim_results_sheet.range('M7').options(index=False, header=False).value = battery_SoC_results.iloc[(backup_res_count*number_of_time_steps)-number_of_time_steps:(backup_res_count*number_of_time_steps),0]
        energy_start_PHEV = PHEV_battery_kWh*(PHEV_initial_SoC - PHEV_min_SoC)
        energy_start_all.append(energy_start_PHEV)
        energy_end_PHEV = PHEV_battery_kWh*(battery_SoC_results.iloc[(backup_res_count*number_of_time_steps)-1,0] - PHEV_min_SoC)
        energy_end_all.append(energy_end_PHEV)

    if backup_source_config_entered.loc[15,2] > 0: # behind the meter storage
        backup_res_count = backup_res_count + 1
        charge_results = batt_charge_CF_results.iloc[(backup_res_count-1)*number_of_time_steps:backup_res_count*number_of_time_steps,0]      
        discharge_results = x_results.iloc[(backup_res_count-1)*number_of_time_steps:(backup_res_count*number_of_time_steps),0]
        charge_results = charge_results.reset_index(drop=True)
        discharge_results = discharge_results.reset_index(drop=True)
        batt_cap_factor_results = charge_results + discharge_results
        optim_results_sheet.range('G7').options(index=False, header=False).value = batt_cap_factor_results
        optim_results_sheet.range('N7').options(index=False, header=False).value = battery_SoC_results.iloc[(backup_res_count*number_of_time_steps)-number_of_time_steps:(backup_res_count*number_of_time_steps),0]
        energy_start_BTM = BTM_kWh*(BTM_initial_SoC - BTM_min_SoC)
        energy_start_all.append(energy_start_BTM)
        energy_end_BTM = BTM_kWh*(battery_SoC_results.iloc[(backup_res_count*number_of_time_steps)-1,0] - BTM_min_SoC)
        energy_end_all.append(energy_end_BTM)

    if backup_source_config_entered.loc[6,2] > 0: # plug-in hybrid electric vehicle ICE
        backup_res_count = backup_res_count + 1
        PHEV_gen_CF_results = x_results.iloc[(backup_res_count-1)*number_of_time_steps:backup_res_count *number_of_time_steps,0]
        optim_results_sheet.range('F7').options(index=False, header=False).value = PHEV_gen_CF_results
        energy_start_PHEV_gen = PHEV_fuel_kWh
        energy_start_all.append(energy_start_PHEV_gen)
        energy_end_PHEV_gen = PHEV_fuel_kWh - PHEV_kW*sum(PHEV_gen_CF_results)*time_step_interval
        energy_end_all.append(energy_end_PHEV_gen)

    if backup_source_config_entered.loc[21,2] > 0: # portable generator
        backup_res_count = backup_res_count + 1
        port_gen_CF_results = x_results.iloc[(backup_res_count-1)*number_of_time_steps:backup_res_count*number_of_time_steps,0]
        optim_results_sheet.range('H7').options(index=False, header=False).value = port_gen_CF_results
        energy_start_port_gen = port_gen_kWh
        energy_start_all.append(energy_start_port_gen)
        energy_end_port_gen = port_gen_kWh - port_gen_kW*sum(port_gen_CF_results)*time_step_interval
        energy_end_all.append(energy_end_port_gen)

    if backup_source_config_entered.loc[25,2] > 0: # standby generator
        backup_res_count = backup_res_count + 1
        standby_CF_results = x_results.iloc[(backup_res_count-1)*number_of_time_steps:backup_res_count*number_of_time_steps,0]
        optim_results_sheet.range('I7').options(index=False, header=False).value = standby_CF_results
        energy_start_standby = standby_kWh
        energy_start_all.append(energy_start_standby)
        energy_end_standby = standby_kWh - standby_kW*sum(standby_CF_results)*time_step_interval
        energy_end_all.append(energy_end_standby)

    if backup_source_config_entered.loc[14,2] > 0: # photovoltaic generation
        backup_res_count = backup_res_count + 1
        PV_to_load_CF_results = pd.DataFrame(x_results.iloc[(backup_res_count-1)*number_of_time_steps:backup_res_count*number_of_time_steps,0])
        PV_to_load_CF_results = PV_to_load_CF_results.reset_index(drop=True)
        PV_to_batt_CF_results = PV_to_batt_results
        PV_to_batt_CF_results = PV_to_batt_CF_results.reset_index(drop=True)
        PV_CF_results = PV_to_load_CF_results + PV_to_batt_CF_results
        optim_results_sheet.range('J7').options(index=False, header=False).value = PV_CF_results
        solved_PV_power_output = PV_CF_results * PV_profile
        calc_values_sheet.range('J7').options(index=False, header=False).value = solved_PV_power_output


    # print the load factor/profile results to Excel

    # for SMM mode: appliances selected
    if count_actual_load_profile == 0:
        all_appliances_load_factors = pd.DataFrame()
        all_solved_appliances_running_watts = pd.DataFrame()
        all_solved_appliances_starting_watts = pd.DataFrame()
        target_energy_all = pd.DataFrame()
        for d in range(number_temp_dep_appliances_selected):
            name_of_appliance = names_temp_dep_appliances_running_all[d]
            load_factor_results = pd.DataFrame(x_results.iloc[(backup_res_count + d)*number_of_time_steps:(backup_res_count + d + 1)*number_of_time_steps,0])
            load_factor_results = load_factor_results.reset_index(drop=True)
            all_appliances_load_factors = pd.concat([all_appliances_load_factors, load_factor_results], axis=1, ignore_index=True)

            solved_appliance_running_watts = load_factor_results.multiply(temp_dep_running_watts_all.iloc[:,d], axis=0)
            all_solved_appliances_running_watts = pd.concat([all_solved_appliances_running_watts, solved_appliance_running_watts], axis=1, ignore_index=True)

            str_watts_factor_results = pd.DataFrame(starting_watts_td_constr_results.iloc[d*number_of_time_steps:(d + 1)*number_of_time_steps,0])
            str_watts_factor_results = str_watts_factor_results.reset_index(drop=True)
            solved_appliance_starting_watts = str_watts_factor_results * temp_dep_add_starting_watts_all[d]
            solved_appliance_starting_watts[solved_appliance_starting_watts < 0] = 0
            all_solved_appliances_starting_watts = pd.concat([all_solved_appliances_starting_watts, solved_appliance_starting_watts], axis=1, ignore_index=True)

            target_energy = pd.DataFrame(temp_dep_running_watts_all.iloc[:,d].sum()*time_step_interval, index=range(1), columns=range(1))
            target_energy_all = pd.concat([target_energy_all, target_energy], axis=1, ignore_index=True)

        for m in range(number_appliances_selected):
            name_of_appliance = names_appliances_running_all[m]
            
            load_factor_results = pd.DataFrame(x_results.iloc[(backup_res_count + number_temp_dep_appliances_selected + m)*number_of_time_steps:(backup_res_count + number_temp_dep_appliances_selected + m + 1)*number_of_time_steps,0])
            load_factor_results = load_factor_results.reset_index(drop=True)
            all_appliances_load_factors = pd.concat([all_appliances_load_factors, load_factor_results], axis=1, ignore_index=True)
            
            solved_appliance_running_watts = load_factor_results * running_watts_all[m]
            all_solved_appliances_running_watts = pd.concat([all_solved_appliances_running_watts, solved_appliance_running_watts], axis=1, ignore_index=True)
            
            str_watts_factor_results = pd.DataFrame(starting_watts_constr_results.iloc[m*number_of_time_steps:(m + 1)*number_of_time_steps,0])
            str_watts_factor_results = str_watts_factor_results.reset_index(drop=True)
            solved_appliance_starting_watts = str_watts_factor_results * add_starting_watts_all[m]
            solved_appliance_starting_watts[solved_appliance_starting_watts < 0] = 0
            all_solved_appliances_starting_watts = pd.concat([all_solved_appliances_starting_watts, solved_appliance_starting_watts], axis=1, ignore_index=True)
            
            target_energy = pd.DataFrame(hours_running_per_day_all[m]*running_watts_all[m]*number_of_days, index=range(1), columns=range(1))
            target_energy_all = pd.concat([target_energy_all, target_energy], axis=1, ignore_index=True)
        
        all_appliances_load_factors.columns = names_temp_dep_appliances_running_all + names_appliances_running_all
        all_appliances_load_factors.columns = all_appliances_load_factors.columns.str.replace("_", " ")
        all_solved_appliances_running_watts.columns = names_temp_dep_appliances_running_all + names_appliances_running_all
        all_solved_appliances_running_watts.columns = all_solved_appliances_running_watts.columns.str.replace("_", " ")
        all_solved_appliances_starting_watts.columns =names_temp_dep_appliances_running_all + names_appliances_running_all
        all_solved_appliances_starting_watts.columns = all_solved_appliances_starting_watts.columns.str.replace("_", " ")

        comp_tool_xl_wb.sheets['Optimization Results']["S6"].options(pd.DataFrame, index=False, expand='table').value = all_appliances_load_factors
        comp_tool_xl_wb.sheets['Calculation Values']["S6"].options(pd.DataFrame, index=False, expand='table').value = all_solved_appliances_running_watts
        comp_tool_xl_wb.sheets['Calculation Values']["BY6"].options(pd.DataFrame, index=False, expand='table').value = all_solved_appliances_starting_watts
        comp_tool_xl_wb.sheets['Calculation Values']["S2"].options(pd.DataFrame, header=False , index=False, expand='table').value = target_energy_all
        

    # for NMM mode: actual home load profile selected
    if count_actual_load_profile > 0:
        actual_load_profile_load_factor_results = pd.DataFrame(x_results.iloc[backup_res_count*number_of_time_steps:(backup_res_count + 1)*number_of_time_steps,0])
        actual_load_profile_load_factor_results = actual_load_profile_load_factor_results.reset_index(drop=True)
        optim_results_sheet.range('S6').options(index=False, header=False).value = 'Actual Home Load Profile [W]'
        optim_results_sheet.range('S7').options(index=False, header=False).value = actual_load_profile_load_factor_results
        calc_values_sheet.range('S2').options(index=False, header=False).value = actual_home_load_profile_day.sum()*time_step_interval
        calc_values_sheet.range('S6').options(index=False, header=False).value = 'Actual Home Load Profile [W]'
        solved_actual_load_profile_results = actual_load_profile_load_factor_results * actual_home_load_profile
        calc_values_sheet.range('S7').options(index=False, header=False).value = solved_actual_load_profile_results
        calc_values_sheet.range('BY6').options(index=False, header=False).value = 'Actual Home Load Profile [W]'
        calc_values_sheet.range('BY7').options(index=False, header=False).value = starting_watts_actual_home_constr_results * actual_home_add_starting_watts



    # calculating backup duration
    user_interface.range("Q13").clear_contents()
    total_energy_start = sum(energy_start_all)
    total_energy_end = sum(energy_end_all)
    change_total_energy = total_energy_start - total_energy_end
    backup_duration = (total_energy_start/change_total_energy)*number_of_days
    if backup_duration < 0 or backup_duration > 365:
        user_interface.range('Q13').options(index=False, header=False).value = 'Infinite'
    elif backup_duration > 0:
        user_interface.range('Q13').options(index=False, header=False).value = backup_duration
    elif backup_duration == 0:
        user_interface.range('Q13').options(index=False, header=False).value = 0
