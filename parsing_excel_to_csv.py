# -*- coding: utf-8 -*-
'''
Find the time and value of max load for each of the regions
COAST, EAST, FAR_WEST, NORTH, NORTH_C, SOUTHERN, SOUTH_C, WEST
and write the result out in a csv file, using pipe character | as the delimiter.

An example output can be seen in the "example.csv" file.
'''

import xlrd
import os
import csv
from zipfile import ZipFile

datafile = "2013_ERCOT_Hourly_Load_Data.xls"
outfile = "2013_Max_Loads.csv"


def open_zip(datafile):
    with ZipFile('{0}.zip'.format(datafile), 'r') as myzip:
        myzip.extractall()


def parse_file(datafile):
    workbook = xlrd.open_workbook(datafile)
    sheet = workbook.sheet_by_index(0)
    data = []

    for i in range(1,9):

    ### take all data from the Column i
        cv = sheet.col_values(i, start_rowx=1, end_rowx= None)
        station_name = sheet.cell_value(0,i)
        print station_name
        #station_name = [str(station_name)]
    ### identify the minimum and maximum values from the i column list
        maxval = max(cv)
    ## use the columns list index to find out the position of the max and min values in the i column value list
    ## add by 1 to all for the reason that the list index starts at zero and tis is a header record
        maxpos = cv.index(maxval) + 1
    # from the min time postition source value from column 0 which is the Min Time
        maxtime = sheet.cell_value(maxpos, 0)
    # convert that value to an time tuple
        realdate = xlrd.xldate_as_tuple(maxtime,0)
        #print realmaxtime
    # record results in data list
        data.append([station_name,realdate[0],realdate[1], realdate[2], realdate[3], maxval])
    #unicode to ASCII

        print data
    return data



def save_file(data, filename):
    with open(filename, "w") as f:
        w = csv.writer(f, delimiter='|')
        w.writerow(["Station", "Year", "Month", "Day", "Hour", "Max Load"])
        w.writerows(data)

def test():
    open_zip(datafile)
    print '--------Start of parse_file process--------'
    data = parse_file(datafile)
    print '--------End of parse_file process ---------'

    print '--------Start of the save file ------------'
    save_file(data, outfile)
    print '--------End of the save file --------------'


    number_of_rows = 0
    stations = []

    ans = {'FAR_WEST': {'Max Load': '2281.2722140000024',
                        'Year': '2013',
                        'Month': '6',
                        'Day': '26',
                        'Hour': '17'}}
    correct_stations = ['COAST', 'EAST', 'FAR_WEST', 'NORTH',
                        'NORTH_C', 'SOUTHERN', 'SOUTH_C', 'WEST']
    fields = ['Year', 'Month', 'Day', 'Hour', 'Max Load']

    with open(outfile) as of:
        csvfile = csv.DictReader(of, delimiter="|")
        for line in csvfile:
            station = line['Station']
            if station == 'FAR_WEST':
                for field in fields:
                    # Check if 'Max Load' is within .1 of answer
                    if field == 'Max Load':
                        max_answer = round(float(ans[station][field]), 1)
                        max_line = round(float(line[field]), 1)
                        assert max_answer == max_line

                    # Otherwise check for equality
                    else:
                        assert ans[station][field] == line[field]

            number_of_rows += 1
            stations.append(station)

        # Output should be 8 lines not including header
        assert number_of_rows == 8

        # Check Station Names
        assert set(stations) == set(correct_stations)



if __name__ == "__main__":
    test()
