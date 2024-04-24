import pandas as pd
import numpy as np 
import os
import matplotlib.pyplot as plt
import sys
import time
import math
import tkinter as tk
from tkinter import filedialog
from bisect import bisect_left

plt.rcParams.update({'figure.max_open_warning': 0})


# ------------------------------ FUNCTIONS ------------------------------
def get_filename_from_user(message): # ----------- Opens a file selection box for user to select input files
  
    root = tk.Tk()
    root.withdraw()
    filename = filedialog.askopenfilename(title=message)

    return filename


def timetable_management(df_Timetable): # ----------- Formats timetable dataframe headers, values etc

    # rename courseID header
    df_Timetable.rename(columns = {"// courseID ": "courseID"},  inplace = True) 

    # remove any spaces from column headers
    df_Timetable.columns = df_Timetable.columns.str.rstrip(' ')

    # remove first row of df (empty row)
    df_Timetable.drop(df_Timetable.index[0], inplace=True)

    # re-adjust row index
    df_Timetable.reset_index(drop=True, inplace=True)

    # remove all rows after timetable (connections table)
    for i in range(len(df_Timetable)):
        if df_Timetable.loc[i, 'courseID'] == '//':
            endOfTimetable = i
            break

    df_Timetable.drop(df_Timetable.index[endOfTimetable:], inplace=True)

    # reformat arrival and departure time to be 00:00:00
    for i in range(len(df_Timetable['arrTime'])):
        if len(df_Timetable.loc[i, 'arrTime']) < 8:
            df_Timetable.loc[i, 'arrTime'] = '0'+df_Timetable.loc[i, 'arrTime']

    for i in range(len(df_Timetable['depTime'])):
        if len(df_Timetable.loc[i, 'depTime']) < 8:
            df_Timetable.loc[i, 'depTime'] = '0'+df_Timetable.loc[i, 'depTime']

    return df_Timetable


def filter_rows(df_Timetable, yards, courseID_and_trainTypes): # ----------- Filter timetable dataframe to only include entries with stoppage information for specified layovers/yards

    # remove rows where stopAtStation is not equal to 1
        # Identify indices where stopAtStation is not equal to 1
    drop_non_stops = df_Timetable.index[df_Timetable['stopAtStation'] != '1']
        # Drop these indices from the dataframe and reset row indices
    df_Timetable.drop(drop_non_stops, inplace=True)
    df_Timetable.reset_index(drop=True, inplace=True)

    # remove rows where stationSign is not in yards list
        # Identify indices where stationSign is not in yards list
    stations_to_drop = df_Timetable[~df_Timetable['stationSign'].isin(yards)].index
        # Drop these indices from the dataframe and reset row indices
    df_Timetable.drop(stations_to_drop, inplace=True)
    df_Timetable.reset_index(drop=True, inplace=True)

    # remove rows where courseID contain CN or VIA
        # Identify indices where courseID contains "CN" 
    drop_CN = df_Timetable['courseID'][df_Timetable['courseID'].str.contains('CN')].index.tolist()
        # Drop these indices from the dataframe and reset row indices
    df_Timetable.drop(drop_CN, inplace=True)
    df_Timetable.reset_index(drop=True, inplace=True)

    # remove rows where courseID contain VIA
        # Identify indices where courseID contains "VIA" 
    drop_VIA = df_Timetable['courseID'][df_Timetable['courseID'].str.contains('VIA')].index.tolist()
        # Drop these indices from the dataframe and reset row indices
    df_Timetable.drop(drop_VIA, inplace=True)
    df_Timetable.reset_index(drop=True, inplace=True)

    # find the train type for each course in the newly filtered timetable using the courseID and TrainType list of arrays
    types = []
    for i in df_Timetable['courseID']:
        for j in courseID_and_trainTypes:
            if i == j[0]:
                types.append(j[1])

    # insert a column of the train types into the timetable dataframe
    df_Timetable.insert(1, 'trainType', types)

    return df_Timetable


def sort_rows(df_Timetable): # ----------- Sort timetable dataframe rows by specific order to make it easier to find information for each layover/yard

    # sort the rows of the filtered timetable based on station, then train type, then departure day, then departure time, then arrival day, then arrival time (and then reset the row indices)
    df_Timetable.sort_values(by = ['stationSign', 'trainType', 'depTimeDayOffset', 'depTime', 'arrTimeDayOffset', 'arrTime'], inplace=True)
    df_Timetable.reset_index(drop=True, inplace=True)

    return df_Timetable


def graph_yard_activity(df_Timetable, yards): # ----------- function where all train arrival and departure information is gathered for each train type in each yard and added to a list of arrays for plotting

    plots = [] # list for arrays containing train information (sorted by each train type in each yard)

    # set beginning x and y values
    x = []
    y_businessDay = [0]
    y_weekendDay = [0]

    # loop through each yard in the list of yards, and define a section of the timetable where all rows of the section have stationSign equal to the specific yard
    for i in yards:
        # if i == 'Glencrest Loop': # -- this can be uncommented to break the program at any specific yard in the list of yards
        #     break
        section = df_Timetable[df_Timetable['stationSign'] == i]
        types = list(dict.fromkeys(section['trainType'])) # creates a list of traintypes from the section with no repeats
        # loop through each type in the list of train types, and define a subsection of the section where all rows of the subsection have trainType equal to the specific type
        for j in types:
            sub_section = section[section['trainType'] == j]
            sub_section.reset_index(drop=True, inplace=True)
            # loop through each row of the subsection, identify if the row lists a train arriving, departing, or both and append the time of the arrival and/or departure to the x values list 
            for line in range(len(sub_section)):
                if sub_section.loc[line, 'arrTime'] != 'HH:MM:SS' and sub_section.loc[line, 'depTime'] == 'HH:MM:SS':
                    x.append((sub_section.loc[line, 'arrTime'], sub_section.loc[line, 'arrTimeDayOffset']))
                    x.append((sub_section.loc[line, 'arrTime'], sub_section.loc[line, 'arrTimeDayOffset']))   # (the repeats are to ensure the line plots have a rectangular pattern)
                elif sub_section.loc[line, 'depTime'] != 'HH:MM:SS' and sub_section.loc[line, 'arrTime'] == 'HH:MM:SS':
                    x.append((sub_section.loc[line, 'depTime'], sub_section.loc[line, 'depTimeDayOffset']))
                    x.append((sub_section.loc[line, 'depTime'], sub_section.loc[line, 'depTimeDayOffset']))
                elif sub_section.loc[line, 'arrTime'] != 'HH:MM:SS' and sub_section.loc[line, 'depTime'] != 'HH:MM:SS':
                    x.append((sub_section.loc[line, 'arrTime'], sub_section.loc[line, 'arrTimeDayOffset']))
                    x.append((sub_section.loc[line, 'arrTime'], sub_section.loc[line, 'arrTimeDayOffset']))  
                    x.append((sub_section.loc[line, 'depTime'], sub_section.loc[line, 'depTimeDayOffset']))
                    x.append((sub_section.loc[line, 'depTime'], sub_section.loc[line, 'depTimeDayOffset']))

            x.sort(key=lambda x: x[1])

            day1 = -1
            day2 = -1
            for index in range(len(x)):
                x_val = x[index]
                if x_val[1] == '1':
                    if day1 == -1:
                        day1 = index
                if x_val[1] == '2':
                    if day2 == -1:
                        day2 = index
                if day1 != -1 and day2 != -1: #if both a day 1 and day 2 entry have been found
                    break
                elif day1 == -1 and day2 != -1: #if there are no day 1 entries, but a day 2 entry was found, break the loop set the day 1 value to be the day 2 value
                    day1 = day2
                    break
            if day1 == -1 and day2 == -1: # if there are no day 1 or day 2 entries in the subsection, the end of day 0 should be the last entry of the list of x values
                day1 = len(x)
            elif day1 != -1 and day2 == -1: # if there is a day 1 entry but no day 2 entries in the subsection, the end of day 2 should be the last entry of the list of x values
                day2 = len(x)
            
            x[0:day1] = sorted(x[0:day1])
            x[day1:day2] = sorted(x[day1:day2])
            x[day2:] = sorted(x[day2:])

            x_businessDay = ['03:00:00']
            x_weekendDay = ['03:00:00']

            for index in range(len(x)):
                x_val = x[index]
                x_time = x_val[0]
                x_day = x_val[1]
                if x_day == '0':
                    x_businessDay.append(x_time)
                elif x_day == '1' and int(x_time[:x_time.find(':')]) < 3:
                    x_businessDay.append(x_time)
                elif x_day == '1' and int(x_time[:x_time.find(':')]) > 2:
                    x_weekendDay.append(x_time)
                elif x_day == '2':
                    x_weekendDay.append(x_time)

            # loop through each x value for a given subsection and call the function that creates an array of y values for business day and weekend day 
            for index in range(1, len(x_businessDay), 2):
                for line in range(len(sub_section)):
                    x_time = x_businessDay[index]
                    if sub_section.loc[line, 'arrTime'] == x_time or sub_section.loc[line, 'depTime'] == x_time: # keep this line to avoid timetable entries where both arrTime and depTime are HH:MM:SS
                        if sub_section.loc[line, 'arrTime'] == x_time:
                            if sub_section.loc[line, 'arrTimeDayOffset'] == '0' or (sub_section.loc[line, 'arrTimeDayOffset'] == '1' and int(x_time[:x_time.find(':')]) < 3): # make sure it is only looking at business day times
                                y_businessDay.append(y_businessDay[len(y_businessDay)-1]+1)
                                y_businessDay.append(y_businessDay[len(y_businessDay)-1])
                        elif sub_section.loc[line, 'depTime'] == x_time:
                            if sub_section.loc[line, 'depTimeDayOffset'] == '0' or (sub_section.loc[line, 'depTimeDayOffset'] == '1' and int(x_time[:x_time.find(':')]) < 3):
                                y_businessDay.append(y_businessDay[len(y_businessDay)-1]-1)
                                y_businessDay.append(y_businessDay[len(y_businessDay)-1])

            for index in range(1, len(x_weekendDay), 2):
                for line in range(len(sub_section)):
                    x_time = x_weekendDay[index]
                    if sub_section.loc[line, 'arrTime'] == x_time or sub_section.loc[line, 'depTime'] == x_time: # keep this line to avoid timetable entries where both arrTime and depTime are HH:MM:SS
                        if sub_section.loc[line, 'arrTime'] == x_time:
                            if sub_section.loc[line, 'arrTimeDayOffset'] == '2' or (sub_section.loc[line, 'arrTimeDayOffset'] == '1' and int(x_time[:x_time.find(':')]) > 2): # make sure it is only looking at weekend day times
                                y_weekendDay.append(y_weekendDay[len(y_weekendDay)-1]+1)
                                y_weekendDay.append(y_weekendDay[len(y_weekendDay)-1])
                        elif sub_section.loc[line, 'depTime'] == x_time:
                            if sub_section.loc[line, 'depTimeDayOffset'] == '2' or (sub_section.loc[line, 'depTimeDayOffset'] == '1' and int(x_time[:x_time.find(':')]) > 2):
                                y_weekendDay.append(y_weekendDay[len(y_weekendDay)-1]-1)
                                y_weekendDay.append(y_weekendDay[len(y_weekendDay)-1])

            # define what yard and train type the subsection is currently looking through
            yardName = sub_section.loc[line, 'stationSign']
            trainType = sub_section.loc[line, 'trainType']

            # cut out only the important part of the train type
            trainType_start_corridor = 0
            trainType_end_corridor = trainType.find('_')

            trainType_start_num = (trainType.find('_', 5, len(trainType))) + 1
            if trainType.find('_', trainType_start_num, len(trainType)) < 0: # -- this if / else is used to work around different train type naming conventions
                trainType_end_num = len(trainType) + 1
            else:
                trainType_end_num = trainType.find('_', trainType_start_num, len(trainType))

            trainType = trainType[trainType_start_corridor:trainType_end_corridor] + '_' + trainType[trainType_start_num:trainType_end_num]

            # find the minimum value of each y array and add that value to each input in each array (so you can see how many trains start at each yard because there can never be less than 0 trains at a yard)
            if min(y_businessDay) < 0:
                min_val = abs(min(y_businessDay))
                for j in range(len(y_businessDay)):
                    y_businessDay[j] += min_val

            if min(y_weekendDay) < 0:
                min_val = abs(min(y_weekendDay))
                for j in range(len(y_weekendDay)):
                    y_weekendDay[j] += min_val

            # add the last value of the day to the x values array and add a 0 at the start of the y value arrays (this is ensures the nice square line plotting)
            x_businessDay.append('03:00:00')  
            x_weekendDay.append('03:00:00')
            y_businessDay.insert(0,y_businessDay[0])
            y_weekendDay.insert(0,y_weekendDay[0])

            # add the arrays for the specific subsection as well as the yard and train type to the list of plots
            plots.append([x_businessDay,x_weekendDay,y_businessDay,y_weekendDay,yardName,trainType])

            # reset x and y values
            x = []
            y_businessDay = [0]
            y_weekendDay = [0]

    return plots


def num_trains_plots(plots, num_trains_plots_dir, type_lengths): # ----------- this function creates the number of trains plots for each yard

    # initialize variables
    max_y_bday = 0
    max_y_wday = 0
    plot_pngs = []
    x_labels_int = []
    length_plots_info = []

    # develop a list of time string values from 3:00:00 to 28:00:00 incrementing by 10 seconds
    x_labels_str = [f'{i:02d}:{j:02d}:00' for i in range(3,28) for j in range(0, 60, 10)]

    # the string ends at the last increment before 28 which is 27:50:00, delete the last 5 values in the list so that the last entry is 27:00:00
    del x_labels_str[-5:-1]
    del x_labels_str[-1]

    # create a list of decimal values from 0 to 27 by turning each value in the list of strings into decimals
    for time_val in x_labels_str:
        x_labels_int.append(convert_time_to_decimal(time_val))

    # convert all string values in the list of strings after 23:50:00 to 24 hours earlier (so 24:00:00 is 00:00:00 and so on)
    for label in range(len(x_labels_int)):
        if x_labels_int[label] > 23.5:
            x_labels_str[label] = convert_decimal_to_time(convert_time_to_decimal(x_labels_str[label]) - 24)

    # bug fix (some values are not exactly 10 minutes after or before the next value, this section fixes those)
    for time in range(len(x_labels_str)):
        timevalue = x_labels_str[time]
        if timevalue[4] != '0':
            timevalue = timevalue[:3] + str(int(timevalue[3])+1) + '0' + timevalue[5:6] + '0' + timevalue[7:]
            x_labels_str[time] = timevalue

    # replacing every 10, 20, 40, and 50 minute interval with a blank value
    for time in range(len(x_labels_str)):
        timevalue = x_labels_str[time]
        if timevalue[3] != '0' and timevalue[3] != '3':
            x_labels_str[time] = ''

    # creating the actual plots starts here:
    print('Outputting first batch of plots:')

    # new variables to pass to length plots function
    plot_info_list = []
    
    for i in plots: # 'plots' being the list of arrays containing the information for each subsection (yard and train type)
                
        x_bday_vals = i[0]
        x_wday_vals = i[1]
        y_bday = i[2]
        y_wday = i[3]
        station = i[4]
        type = i[5]
        type_num = type[type.find('_')+1:len(type)]

        # convert each x value into a decimal value and add 24 to the day 1 and further values (so that the values are decimals up to 28.0 which is 3am)
        for value in range(len(x_bday_vals)):
            x_bday_vals[value] = convert_time_to_decimal(x_bday_vals[value])
        for value in range(len(x_bday_vals)):
            if value != 0 and x_bday_vals[value-1] > x_bday_vals[value]:
                for index in range(value,len(x_bday_vals)):
                    x_bday_vals[index] += 24
        if len(x_bday_vals) == 2: # this makes the second value 28.0 if there are only 2 x values (meaning there were no entries found in the timetable so its just ['03:00:00', '03:00:00'])
            x_bday_vals[1] += 24
        
        for value in range(len(x_wday_vals)):
            x_wday_vals[value] = convert_time_to_decimal(x_wday_vals[value])
        for value in range(len(x_wday_vals)):
            if value != 0 and x_wday_vals[value-1] > x_wday_vals[value]:
                for index in range(value,len(x_wday_vals)):
                    x_wday_vals[index] += 24
        if len(x_wday_vals) == 2: # this makes the second value 28.0 if there are only 2 x values (meaning there were no entries found in the timetable so its just ['03:00:00', '03:00:00'])
            x_wday_vals[1] += 24

        # if the current 'plot' index is not the last entry in the list of plots:
        if plots.index(i) != len(plots)-1:

            # find the next plot's station and the previous plots station (to see if they are the same and if so then the plots should be graphed on the same figure, this means they should be different train type plots from the same yard)
            next = plots[plots.index(i)+1]
            next_station = next[4]
            last = plots[plots.index(i)-1]
            last_station = last[4]

            # if the next or last station is the same as the current station, the max y value for business day and weekend day considers all values from every train type in the same yard)
            if station == next_station and station != last_station:
                max_y_bday = max(y_bday)
                max_y_wday = max(y_wday)
            elif station == next_station and station == last_station:
                if max(y_bday) > max_y_bday:
                    max_y_bday = max(y_bday)
                if max(y_wday) > max_y_wday:
                    max_y_wday = max(y_wday)
            elif station != next_station and station == last_station:
                if max(y_bday) > max_y_bday:
                    max_y_bday = max(y_bday)
                if max(y_wday) > max_y_wday:
                    max_y_wday = max(y_wday)
            else:
                max_y_bday = max(y_bday)
                max_y_wday = max(y_wday)

            # info for length plots
            # new variables to pass to length plots function
            y_bday_len = []
            y_wday_len = []
            value = 0
            
            # converting all of the y values for business and weekend days into length values dependant on the train type
            for value in y_bday:
                y_bday_len.append(value*type_lengths[type_num])
            value = 0
            for value in y_wday:
                y_wday_len.append(value*type_lengths[type_num])
            
            plot_info = [x_bday_vals, x_wday_vals, x_labels_int, x_labels_str, y_bday_len, y_wday_len, station, type]
            plot_info_list.append(plot_info)

            # new variables to plot a label of every other y value
            xbs = []
            xws = []
            ybs = []
            yws = []
            for x1 in range(1,len(np.array(x_bday_vals)),2):
                xbs.append(x_bday_vals[x1])
            for x2 in range(1,len(np.array(x_wday_vals)),2):
                xws.append(x_wday_vals[x2])
            for y1 in range(1,len(np.array(y_bday)),2):
                ybs.append(y_bday[y1])
            for y2 in range(1,len(np.array(y_wday)),2):
                yws.append(y_wday[y2])

            # business day plotting info
            plt.figure(1, figsize = (20, 7), dpi = 100)
            ax = plt.gca()
            plt.plot(np.array(x_bday_vals),np.array(y_bday), label = type)
            color = ax.get_lines()[-1].get_color()
            plt.plot(np.array(x_bday_vals),np.array(y_bday), marker='o', ms = 5, color = color)
            for x, y in zip(xbs, ybs):
                label = f"{y}"
                plt.annotate(label, (x, y), color = color, textcoords="offset points", xytext=(5, 0))
            plt.xticks(x_labels_int, labels = x_labels_str, rotation = 90)
            plt.yticks(np.arange(0, max(max_y_bday, max_y_wday)+3, step = 1))
            plt.axis([min(x_labels_int), max(x_labels_int), 0, max(max_y_bday, max_y_wday)+2])
            plt.grid(visible = True, linewidth=0.5)
            plt.title(station + ' Business Day Train Count')

            # for the last plot for a given yard, add the max capacity line and legend to the plots, adjust the layout view, and save the figures
            if station != next_station:
                plt.figure(1, figsize = (20, 7), dpi = 100)
                plt.legend()
                plt.tight_layout(h_pad=2)
                plt.savefig(num_trains_plots_dir + station + '_Business_Day_Plot.png')
                plt.clf()
                plot_pngs.append(num_trains_plots_dir + station + '_Business_Day_Plot.png') # saving the plot figure names to a list to be referred to later

            # weekend day plotting info
            plt.figure(2, figsize = (20, 7), dpi = 100)
            ax = plt.gca()
            plt.plot(np.array(x_wday_vals),np.array(y_wday), label = type)
            color = ax.get_lines()[-1].get_color()
            plt.plot(np.array(x_wday_vals),np.array(y_wday), marker='o', ms = 5, color = color)
            for x, y in zip(xws, yws):
                label = f"{y}"
                plt.annotate(label, (x, y), color = color, textcoords="offset points", xytext=(5, 0))
            plt.xticks(x_labels_int, labels = x_labels_str, rotation = 90)
            plt.yticks(np.arange(0, max(max_y_bday, max_y_wday)+3, step = 1))
            plt.axis([min(x_labels_int), max(x_labels_int), 0, max(max_y_bday, max_y_wday)+2])
            plt.grid(visible = True, linewidth=0.5)
            plt.title(station + ' Weekend Day Train Count')

            # for the last plot for a given yard, add the max capacity line and legend to the plots, adjust the layout view, and save the figures
            if station != next_station:
                plt.figure(2, figsize = (20, 7), dpi = 100)
                plt.legend()
                plt.tight_layout(h_pad=2)
                plt.savefig(num_trains_plots_dir + station + '_Weekend_Day_Plot.png')
                plt.clf()
                plot_pngs.append(num_trains_plots_dir + station + '_Weekend_Day_Plot.png') # saving the plot figure names to a list to be referred to later

                print(str(int(100*((plots.index(i)+1)/len(plots))))+'%') # this just gives the user an update of how far the program is
                length_plots_info.append(plot_info_list)
                y_bday_len = []
                y_wday_len = []
                plot_info_list = []
        
        else:   # for the last plot in the list of plots (same as last plot for a given yard)
            last = plots[plots.index(i)-1]
            last_station = last[4]
            if station == last_station:
                if max(y_bday) > max_y_bday:
                    max_y_bday = max(y_bday)
                if max(y_wday) > max_y_wday:
                    max_y_wday = max(y_wday)

            elif station != last_station:
                max_y_bday = max(y_bday)
                max_y_wday = max(y_wday)

                # info for length plots
                # new variables to pass to length plots function
                y_bday_len = []
                y_wday_len = []
                value = 0

                # converting all of the y values for business and weekend days into length values dependant on the train type
                for value in y_bday:
                    y_bday_len.append(value*type_lengths[type_num])
                value = 0
                for value in y_wday:
                    y_wday_len.append(value*type_lengths[type_num])
                
                plot_info = [x_bday_vals, x_wday_vals, x_labels_int, x_labels_str, y_bday_len, y_wday_len, station, type]
                plot_info_list.append(plot_info)

                # new variables to plot a label of every other y value
                xbs = []
                xws = []
                ybs = []
                yws = []
                for x1 in range(1,len(np.array(x_bday_vals)),2):
                    xbs.append(x_bday_vals[x1])
                for x2 in range(1,len(np.array(x_wday_vals)),2):
                    xws.append(x_wday_vals[x2])
                for y1 in range(1,len(np.array(y_bday)),2):
                    ybs.append(y_bday[y1])
                for y2 in range(1,len(np.array(y_wday)),2):
                    yws.append(y_wday[y2])
                    
                # business day plotting info
                plt.figure(1, figsize = (20, 7), dpi = 100)
                ax = plt.gca()
                plt.plot(np.array(x_bday_vals),np.array(y_bday), label = type)
                color = ax.get_lines()[-1].get_color()
                plt.plot(np.array(x_bday_vals),np.array(y_bday), marker='o', ms = 5, color = color)
                for x, y in zip(xbs, ybs):
                    label = f"{y}"
                    plt.annotate(label, (x, y), color = color, textcoords="offset points", xytext=(5, 0))
                plt.xticks(x_labels_int, labels = x_labels_str, rotation = 90)
                plt.yticks(np.arange(0, max(max_y_bday, max_y_wday)+3, step = 1))
                plt.axis([min(x_labels_int), max(x_labels_int), 0, max(max_y_bday, max_y_wday)+2])
                plt.grid(visible = True, linewidth=0.5)
                plt.title(station + ' Business Day Train Count') 
                plt.legend()
                plt.tight_layout(h_pad=2)
                plt.savefig(num_trains_plots_dir + station + '_Business_Day_Plot.png')
                plt.clf()
                plot_pngs.append(num_trains_plots_dir + station + '_Business_Day_Plot.png') # saving the plot figure names to a list to be referred to later

                # weekend day plotting info
                plt.figure(2, figsize = (20, 7), dpi = 100)
                ax = plt.gca()
                plt.plot(np.array(x_wday_vals),np.array(y_wday), label = type)
                color = ax.get_lines()[-1].get_color()
                plt.plot(np.array(x_wday_vals),np.array(y_wday), marker='o', ms = 5, color = color)
                for x, y in zip(xws, yws):
                    label = f"{y}"
                    plt.annotate(label, (x, y), color = color, textcoords="offset points", xytext=(5, 0))
                plt.xticks(x_labels_int, labels = x_labels_str, rotation = 90)
                plt.yticks(np.arange(0, max(max_y_bday, max_y_wday)+3, step = 1))
                plt.axis([min(x_labels_int), max(x_labels_int), 0, max(max_y_bday, max_y_wday)+2])
                plt.grid(visible = True, linewidth=0.5)
                plt.title(station + ' Weekend Day Train Count')
                plt.legend()
                plt.tight_layout(h_pad=2)
                plt.savefig(num_trains_plots_dir + station + '_Weekend_Day_Plot.png')
                plt.clf()
                plot_pngs.append(num_trains_plots_dir + station + '_Weekend_Day_Plot.png') # saving the plot figure names to a list to be referred to later

                print(str(int(100*((plots.index(i)+1)/len(plots))))+'%') # this just gives the user an update of how far the program is
                length_plots_info.append(plot_info_list)
                y_bday_len = []
                y_wday_len = []
                plot_info_list = []

    return plot_pngs, length_plots_info


def len_trains_plots(length_plots_info, len_trains_plots_dir, max_train_capacities_len, yards): # ----------- this function creates the total train length activity plots for each yard
    
    # initialize variables
    plot_pngs = []
    sum_x_bday_vals = []
    sum_x_wday_vals = []
    peaks = []

    print('\nOutputting second batch of plots:')
    # loop through each of the indices of the length_plots_info array and retrieve the array of info for each yard
    for i in length_plots_info:
        yard_info = i

        # loop through each index in the yard info array to retrieve the array of info for each train type at the yard
        for j in yard_info:
            # retrieve the specific graph information from the array
            x_bday_vals = j[0]
            x_wday_vals = j[1]
            x_labels_int = j[2]
            x_labels_str = j[3]
            station = j[6]
            max_capacity = int(max_train_capacities_len[yards.index(station)])

            # when creating the list of all x values across all train types at a station, this deletes the 3:00:00 and 27:00:00 values to make sure there are no duplicates
            if len(sum_x_bday_vals) > 0:
                del sum_x_bday_vals[0]
                del sum_x_bday_vals[-1]
            
            if len(sum_x_wday_vals) > 0:
                del sum_x_wday_vals[0]
                del sum_x_wday_vals[-1]
            
            # create an extended list that has all x_values for all train types through a station
            for val in range(len(x_bday_vals)):
                sum_x_bday_vals.append(x_bday_vals[val])
            sum_x_bday_vals.sort()

            for val in range(len(x_wday_vals)):
                sum_x_wday_vals.append(x_wday_vals[val])
            sum_x_wday_vals.sort()


        # initializing extended y value lists 
        sum_y_bday_len = [0]*len(sum_x_bday_vals)
        sum_y_wday_len = [0]*len(sum_x_wday_vals)
        
        # loop through each train type array again once the extended x values list is complete to create the extended y values lists
        for k in yard_info:
            x_bday_vals = k[0]
            x_wday_vals = k[1]
            y_bday_len = k[4]
            y_wday_len = k[5]
            y_bday_interpolated = np.interp(sum_x_bday_vals, x_bday_vals, y_bday_len)
            y_wday_interpolated = np.interp(sum_x_wday_vals, x_wday_vals, y_wday_len)

            for x in range(len(y_bday_interpolated)):
                sum_y_bday_len[x] += y_bday_interpolated[x]
            for x in range(len(y_wday_interpolated)):
                sum_y_wday_len[x] += y_wday_interpolated[x]

        # gets rid of duplicates and adds y value at 3:00:00
        del sum_y_bday_len[-1]
        sum_y_bday_len.insert(0, sum_y_bday_len[0])
        del sum_y_wday_len[-1]
        sum_y_wday_len.insert(0, sum_y_wday_len[0])
        
        # new variables to plot a label of every other y value
        xbs = []
        xws = []
        ybs = []
        yws = []
        for x1 in range(1,len(np.array(sum_x_bday_vals)),2):
            xbs.append(sum_x_bday_vals[x1])
        for x2 in range(1,len(np.array(sum_x_wday_vals)),2):
            xws.append(sum_x_wday_vals[x2])
        for y1 in range(1,len(np.array(sum_y_bday_len)),2):
            ybs.append(sum_y_bday_len[y1])
        for y2 in range(1,len(np.array(sum_y_wday_len)),2):
            yws.append(sum_y_wday_len[y2])

        # business day plotting info
        plt.figure(3, figsize = (20, 7), dpi = 100)
        plt.plot(np.array(sum_x_bday_vals),np.array(sum_y_bday_len), color = 'green')
        plt.plot(np.array(sum_x_bday_vals),np.array(sum_y_bday_len), marker='o', color = 'green', ms = 5)
        for x, y in zip(xbs, ybs):
            label = f"{y}m"
            plt.annotate(label, (x, y), textcoords="offset points", xytext=(5, 0), color = 'black')
        plt.xticks(x_labels_int, labels = x_labels_str, rotation = 90)
        plt.yticks(np.arange(0, max(max(sum_y_bday_len), max(sum_y_wday_len), max_capacity)+400, step = 200))
        plt.axis([min(x_labels_int), max(x_labels_int), 0, max(max(sum_y_bday_len), max(sum_y_wday_len), max_capacity)+400])
        plt.grid(visible = True, linewidth=0.5)
        plt.title(station + ' Business Day Total Occupancy Length (meters)') 
        plt.plot(np.array(x_bday_vals),len(np.array(x_bday_vals))*[max_capacity], color = 'red', linestyle = 'dashed')
        label = f"Maximum Capacity = {max_capacity}m"
        plt.annotate(label, (23.0, max_capacity), textcoords="offset points", xytext=(0, 5), color = 'red')
        plt.tight_layout(h_pad=2)
        plt.savefig(len_trains_plots_dir + station + '_Business_Day_Plot.png')
        plt.clf()
        plot_pngs.append(len_trains_plots_dir + station + '_Business_Day_Plot.png') # saving the plot figure names to a list to be referred to later

        # weekend day plotting info
        plt.figure(4, figsize = (20, 7), dpi = 100)
        plt.plot(np.array(sum_x_wday_vals),np.array(sum_y_wday_len), color = 'green')
        plt.plot(np.array(sum_x_wday_vals),np.array(sum_y_wday_len), marker='o', color = 'green', ms = 5)
        for x, y in zip(xws, yws):
            label = f"{y}m"
            plt.annotate(label, (x, y), textcoords="offset points", xytext=(5, 0), color = 'black')
        plt.xticks(x_labels_int, labels = x_labels_str, rotation = 90)
        plt.yticks(np.arange(0, max(max(sum_y_bday_len), max(sum_y_wday_len), max_capacity)+400, step = 200))
        plt.axis([min(x_labels_int), max(x_labels_int), 0, max(max(sum_y_bday_len), max(sum_y_wday_len), max_capacity)+400])
        plt.grid(visible = True, linewidth=0.5)
        plt.title(station + ' Weekend Day Total Occupancy Length (meters)')
        plt.plot(np.array(x_wday_vals),len(np.array(x_wday_vals))*[max_capacity], color = 'red', linestyle = 'dashed')
        label = f"Maximum Capacity = {max_capacity}m"
        plt.annotate(label, (23.0, max_capacity), textcoords="offset points", xytext=(0, 5), color = 'red')
        plt.tight_layout(h_pad=2)
        plt.savefig(len_trains_plots_dir + station + '_Weekend_Day_Plot.png')
        plt.clf()
        plot_pngs.append(len_trains_plots_dir + station + '_Weekend_Day_Plot.png') # saving the plot figure names to a list to be referred to later

        print(str(int(100*((length_plots_info.index(i)+1)/len(length_plots_info))))+'%') # this just gives the user an update of how far the program is
        
        # this is used for a calculation later on
        peaks.append([sum_x_bday_vals[sum_y_bday_len.index(max(sum_y_bday_len))], sum_x_wday_vals[sum_y_wday_len.index(max(sum_y_wday_len))]])
        
        # resetting variable
        sum_x_bday_vals = []
        sum_x_wday_vals = []

    return plot_pngs, peaks


def convert_time_to_decimal(time_str): # ----------- Function to convert time strings (8:00:01) to hour as float (8.0002778...)

    hour, minute, second = map(int, time_str.split(':'))
    decimal_hour = hour + minute / 60 + second / 3600

    return float(decimal_hour)


def convert_decimal_to_time(decimal_hour): # ----------- Function to convert float value (8.0002778...) to time strings (8:00:01)
    day, second = divmod(decimal_hour * 3600, 86400)
    hour, second = divmod(second, 3600)
    minute, second = divmod(second, 60)
    second = round(second)

    return "%02d:%02d:%02d" % (hour, minute, second)


def take_closest(myList, myNumber): # ----------- This function is used to find the closest value in a list to a new given value, if two values are equal distance from the new value, the function returns the smaller value
    
    pos = bisect_left(myList, myNumber)
    if pos == 0:
        return myList[0]
    if pos == len(myList):
        return myList[-1]
    before = myList[pos - 1]
    after = myList[pos]
    if after - myNumber < myNumber - before:
        return after
    else:
        return before


def createChecksFile(yardBalancingChecks,plots,yards, peaks): # ----------- this function creates the results output sheet and conducts all of the necessary checks on yard data

    for j in range(len(plots)):
        # set initial given information
        info = plots[j]
        x_bday_vals = info[0]
        x_wday_vals = info[1]
        y_bday = info[2]
        y_wday = info[3]
        station = info[4]
        type = info[5]
        if j % 2 == 0: # checks if j is an even number so that it cycles through yard_peaks index as follows: 0, 0, 1, 1, 2, 2, 3, 3, ...
            yard_peaks = peaks[int(j/2)]
            business_day_peak = yard_peaks[0]
            weekend_day_peak = yard_peaks[1]


        yardBalancingChecks.loc[j,'stationSign'] = station
        yardBalancingChecks.loc[j, 'Train Type'] = type
        yardBalancingChecks.loc[j, 'Number of Trains at Beginning of Business Day'] = y_bday[0]
        yardBalancingChecks.loc[j, 'Number of Trains at End of Business Day'] = y_bday[len(y_bday)-1]
        yardBalancingChecks.loc[j, 'Number of Trains at Beginning of Weekend Day'] = y_wday[0]
        yardBalancingChecks.loc[j, 'Number of Trains at End of Weekend Day'] = y_wday[len(y_wday)-1]
        yardBalancingChecks.loc[j, 'Number of Trains at Yard During Peak Hour on Business Day'] = y_bday[x_bday_vals.index(take_closest(x_bday_vals, business_day_peak))]
        yardBalancingChecks.loc[j, 'Number of Trains at Yard During Peak Hour on Weekend Day'] = y_wday[x_wday_vals.index(take_closest(x_wday_vals, weekend_day_peak))]
        yardBalancingChecks.loc[j, 'Peak Hour on Business Day'] = convert_decimal_to_time(business_day_peak)
        yardBalancingChecks.loc[j, 'Peak Hour on Weekend Day'] = convert_decimal_to_time(weekend_day_peak)


    # conducting compliance checks 
    for row in range(len(yardBalancingChecks)):
        if yardBalancingChecks.loc[row, 'Number of Trains at Beginning of Business Day'] == yardBalancingChecks.loc[row, 'Number of Trains at End of Business Day']:
            yardBalancingChecks.iloc[row, 4] = 'Complies'
        else:
            yardBalancingChecks.iloc[row, 4] = 'Does Not Comply'

    for row in range(len(yardBalancingChecks)):
        if yardBalancingChecks.loc[row, 'Number of Trains at Beginning of Weekend Day'] == yardBalancingChecks.loc[row, 'Number of Trains at End of Weekend Day']:
            yardBalancingChecks.iloc[row, 7] = 'Complies'
        else:
            yardBalancingChecks.iloc[row, 7] = 'Does Not Comply'

    stations_list = []
    for i in plots:
        station = i[4]
        stations_list.append(station)
    stations_list = list(dict.fromkeys(stations_list))

    # create 'total' rows for each yard
    for line in range(len(stations_list)):
        total_bday_BoD = 0
        total_bday_EoD = 0
        total_bday_Compliance = ''
        total_wday_BoD = 0
        total_wday_EoD = 0
        total_wday_Compliance = ''
        total_bday_peak = 0
        total_wday_peak = 0

        section = yardBalancingChecks[yardBalancingChecks['stationSign'] == stations_list[line]]
        section.reset_index(drop=True, inplace=True)

        for row in range(len(section)):
            total_bday_BoD += section.loc[row, 'Number of Trains at Beginning of Business Day']
            total_bday_EoD += section.loc[row, 'Number of Trains at End of Business Day']
            total_wday_BoD += section.loc[row, 'Number of Trains at Beginning of Weekend Day']
            total_wday_EoD += section.loc[row, 'Number of Trains at End of Weekend Day']
            total_bday_peak += section.loc[row, 'Number of Trains at Yard During Peak Hour on Business Day']
            total_wday_peak += section.loc[row, 'Number of Trains at Yard During Peak Hour on Weekend Day']

        if total_bday_BoD == total_bday_EoD:
            total_bday_Compliance = 'Complies'
        else:
            total_bday_Compliance = 'Does Not Comply'

        if total_wday_BoD == total_wday_EoD:
            total_wday_Compliance = 'Complies'
        else:
            total_wday_Compliance = 'Does Not Comply'

        station_index = yardBalancingChecks.index[yardBalancingChecks['stationSign']==stations_list[line]].tolist()

        if line % 2 == 0:
            yard_peaks = peaks[int(line/2)]
            business_day_peak = convert_decimal_to_time(yard_peaks[0])
            weekend_day_peak = convert_decimal_to_time(yard_peaks[1])

        yardBalancingChecks.loc[station_index[-1]+0.5] = [stations_list[line]+' Total', '', total_bday_BoD, total_bday_EoD, total_bday_Compliance, total_wday_BoD, total_wday_EoD, total_wday_Compliance, business_day_peak, total_bday_peak, weekend_day_peak, total_wday_peak]
        yardBalancingChecks = yardBalancingChecks.sort_index().reset_index(drop=True)

    return yardBalancingChecks


# ------------------------------ MAIN FUNCTION ------------------------------

def main():

    # import layovers and yards file as df
    yards_list_input = get_filename_from_user('Select the list of yards file you would like to use (must be xlsx format)')
    df_list_of_yards = pd.read_excel(yards_list_input, dtype = str)

    # create a list of yards from the input df
    yards = []
    for i in range(len(df_list_of_yards)):
        yards.append(df_list_of_yards.iloc[i, 0])

    # create a list of max train capacities from the input df
    max_train_capacities_len = []
    for i in range(len(df_list_of_yards)):
        max_train_capacities_len.append(df_list_of_yards.iloc[i, 1])

    # create a dictionary of train types and their respecitve lengths from the input df
    type_lengths = {}
    for i in range(df_list_of_yards[df_list_of_yards.columns[3]].count()):
        type = str(df_list_of_yards.iloc[i, 2])
        length = int(df_list_of_yards.iloc[i, 3])
        type_lengths[type] = length


    # import course_xml file as df
    course_xml_input = get_filename_from_user('Select the course_xml file you would like to use (must be csv format)')
    df_course_xml = pd.read_csv(course_xml_input, dtype = str)

    # create a list of arrays of courseIDs and their respective TrainTypes from the course_xml input df
    courseID_and_trainTypes = []
    for i in range(len(df_course_xml)):
        courseID_and_trainTypes.append([df_course_xml.loc[i, 'CourseID'], df_course_xml.loc[i, 'TrainType']])


    # ask user to input timetable
    timetable_input = get_filename_from_user('Select the timetable you would like to use (must be csv format)')
    df_Timetable = pd.read_csv(timetable_input, header = 11, dtype = str)

    # run functions to filter and sort timetable
    df_Timetable = timetable_management(df_Timetable)
    df_Timetable = filter_rows(df_Timetable, yards, courseID_and_trainTypes)
    df_Timetable = sort_rows(df_Timetable)

    # run function that creates a list of information for each plot 
    plots_array = graph_yard_activity(df_Timetable, yards)


    # create a new folder named 'Number of Trains Plots' in the same location as the script, if a folder with this name already exists, mention it and kill the program (this is because we don't want to accidentally overwrite old plots and also want to make sure that the script outputs the new plots in the excel file)
    script_dir = os.path.dirname(__file__)
    num_trains_plots_dir = os.path.join(script_dir, 'Number of Trains Plots/')
    if not os.path.isdir(num_trains_plots_dir):
        os.makedirs(num_trains_plots_dir)
    else:
        print("'Number of Trains Plots' folder already exists in directory, delete or move folder and re-run program")
        sys.exit()

    # output plots / get list of plot pngs
    num_trains_plots_pngs, length_plots_info = num_trains_plots(plots_array, num_trains_plots_dir, type_lengths)

    # create a new folder named 'Length of Trains Plots' in the same location as the script, if a folder with this name already exists, mention it and kill the program (this is because we don't want to accidentally overwrite old plots and also want to make sure that the script outputs the new plots in the excel file)
    script_dir = os.path.dirname(__file__)
    len_trains_plots_dir = os.path.join(script_dir, 'Length of Trains Plots/')
    if not os.path.isdir(len_trains_plots_dir):
        os.makedirs(len_trains_plots_dir)
    else:
        print("'Length of Trains Plots' folder already exists in directory, delete or move folder and re-run program")
        sys.exit()

    # output plots / get list of plot pngs
    len_trains_plots_pngs, peaks = len_trains_plots(length_plots_info, len_trains_plots_dir, max_train_capacities_len, yards)


    # create dataframe for yard checks
    yardBalancingChecks = pd.DataFrame(columns=['stationSign',
                                                'Train Type',
                                                'Number of Trains at Beginning of Business Day', 'Number of Trains at End of Business Day', 'Compliance',
                                                'Number of Trains at Beginning of Weekend Day', 'Number of Trains at End of Weekend Day', 'Compliance',
                                                'Peak Hour on Business Day', 'Number of Trains at Yard During Peak Hour on Business Day', 'Peak Hour on Weekend Day', 'Number of Trains at Yard During Peak Hour on Weekend Day'])

    #'Max Number of Trains at Yard on Business Day', 'Max Number of Trains at Yard on Weekend Day', 'Max Number of Trains Allowed at Yard', 'Compliance',


    # run functions for yard checks
    yardBalancingChecks = createChecksFile(yardBalancingChecks, plots_array, yards, peaks)

    # find the index of each 'total' row in the checks to later make these rows have a red font
    totals_row_index =[]
    for row in range(len(yardBalancingChecks)):
        if yardBalancingChecks.loc[row, 'stationSign'].find('Total') > 0:
            totals_row_index.append(row+1)

    # output excel file with multiple sheets and formatting
    with pd.ExcelWriter('Yard Balance Checks.xlsx') as writer:
        df_list_of_yards.to_excel(writer, sheet_name = 'Input', startrow = 1, index = False, header = False)
        workbook = writer.book
        worksheet_input = writer.sheets['Input']
        blue_header_format = workbook.add_format({'bg_color': '#61CBF3'})
        purple_header_format = workbook.add_format({'bg_color': '#D86DCD'})
        blue_cells_format = workbook.add_format({'bg_color': '#CAEDFB'})
        purple_cells_format = workbook.add_format({'bg_color': '#F2CEEF'})
        
        for col_num, value in enumerate(df_list_of_yards.columns.values):
            if col_num < 2:
                worksheet_input.write(0, col_num, value, blue_header_format)
            elif col_num < len(df_list_of_yards.columns):
                worksheet_input.write(0, col_num, value, purple_header_format)

        for col in range(len(df_list_of_yards.columns)):
            for row in range(df_list_of_yards[df_list_of_yards.columns[col]].count()):
                value = df_list_of_yards.iloc[row,col]
                if col < 2:
                    worksheet_input.write(row+1, col, value, blue_cells_format)
                else:
                    worksheet_input.write(row+1, col, value, purple_cells_format)

        df_Timetable.to_excel(writer, sheet_name = 'Sorted and Filtered Timetable', index = False)

        yardBalancingChecks.to_excel(writer, sheet_name = 'Results', index = False)
        workbook = writer.book
        worksheet_results = writer.sheets['Results']
        bold_and_bottom_border_format = workbook.add_format({'bold': True, 'bottom': 1})
        green_font_and_highlight_format = workbook.add_format({'bg_color':   '#C6EFCE',
                               'font_color': '#006100'})
        red_font_and_highlight_format = workbook.add_format({'bg_color':   '#FFC7CE',
                               'font_color': '#9C0006'})

        headers = list(yardBalancingChecks.columns.values)

        for col in range(len(headers)):
            if headers[col] == "Compliance":
                worksheet_results.conditional_format(0, col, len(yardBalancingChecks), col, {'type' : 'cell', 'criteria' : 'equal to', 'value' : '"Does Not Comply"', 'format' : red_font_and_highlight_format})
                worksheet_results.conditional_format(0, col, len(yardBalancingChecks), col, {'type' : 'cell', 'criteria' : 'equal to', 'value' : '"Complies"', 'format' : green_font_and_highlight_format})

        for row in totals_row_index:
            worksheet_results.set_row(row, None, bold_and_bottom_border_format)

        pd.DataFrame().to_excel(writer, sheet_name = 'Plots')
        worksheet_plots = writer.sheets['Plots']
        row1 = 1
        row2 = 18
        col = 'A'
        image_width = 30
        image_height = 34
        cell_width = 14
        cell_height = 17
        x_scale = cell_width/image_width
        y_scale = cell_height/image_height
        for i in range(len(len_trains_plots_pngs)):
            worksheet_plots.insert_image(col+str(row1),len_trains_plots_pngs[i], {'x_scale': x_scale, 'y_scale': y_scale})
            worksheet_plots.insert_image(col+str(row2),num_trains_plots_pngs[i], {'x_scale': x_scale, 'y_scale': y_scale})
            if col == 'A':
                col = 'P'
            elif col == 'P':
                col = 'A'
                row1 += 36
                row2 += 36

    return


# ============================== MAIN ==============================

if __name__ == "__main__":

    start_time = time.time() # store start time
    
    main()

    time_passed = time.time()-start_time
    minutes = math.trunc(time_passed/60)
    seconds = time_passed - (math.trunc(time_passed/60))*60
    print("\nProcess finished ---", minutes, "minutes and", round(seconds, 2), "seconds ---") # Print end time
