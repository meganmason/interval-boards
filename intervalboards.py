import datetime
import glob
import os
import shutil
import numpy as np
import pandas as pd
from csv import writer
import textwrap
from pathlib import Path
import utm
from openpyxl import load_workbook

'''
Some workflow to get started with the 2020 Interval Board dataset. Script was copy/pasted from Main script to process snow pit data.
A lot of irrelevant script was removed, but it has not be thoroughly edited for variables that do not pertain to the interval board data set.

Eventually the input files should be the published Time Series snow pit (.xlsx) files.
For right now, grabbing files from the timeseries_pitbook_sheets_EDIT folder is fine (however, please make a copy and do not work from that main folder).

What's the difference?  - the published files will have fixed coordinates for the Snow Pits (same coords will be used for the interval board for 2020),
whereas the files to start with the coordinates will vary, but that is OKAY to get started.

That parameter named assigned to the Interval Board dataset = newSnow
'''

def readIntervalBoard(path_in, filename, version, path_out):

    # filename variables
    fname = os.path.basename(filename)  # get just the filename (could write with PathLib, but not a priority now)
    pitID = filename.stem.split('_')[0] # splits the filename with delimeter "_" and saves the first
    dateString = filename.stem.split('_')[1] # splits filename and saves the second
    subString = fname.rstrip(".xlsx")
    boardPath = path_out + 'boards/' + subString + '/'
    if not os.path.exists(boardPath):
        os.makedirs(boardPath)

    # open excel file
    xl = pd.ExcelFile(filename)

    # create individual output filename
    fname_newSnow = boardPath + '/SNEX20_TS_IB_' + dateString + '_' + pitID + '_newSnow_' + version +'.csv'

    # location / pit name
    d = pd.read_excel(xl, sheet_name=0, usecols='B')
    Location = d['Location:'][0]
    Site = d['Location:'][2]  # grab site name
    PitID = d['Location:'][4][:6]  # grab pit name
    d = pd.read_excel(xl, sheet_name=0, usecols='L')
    UTME = d['Surveyors:'][2]
    d = pd.read_excel(xl, sheet_name=0, usecols='Q')
    UTMN = d['Unnamed: 16'][2]
    d = pd.read_excel(xl, sheet_name=0, usecols='X')
    UTMzone = d['Unnamed: 23'][2] #for northern hemisphere
    # convert to Lat/Lon
    LAT = round(utm.to_latlon(UTME, UTMN, UTMzone, "Northern")[0], 5) #tuple output, save first
    LON = round(utm.to_latlon(UTME, UTMN, UTMzone, "Northern")[1], 5) #tuple output, save second

    # total depth
    d = pd.read_excel(xl, sheet_name=0, usecols='G')
    TotalDepth = d['Unnamed: 6'][4]  # grab total depth

    # date and time
    d = pd.read_excel(xl, sheet_name=0, usecols='X')
    pit_time = d['Unnamed: 23'][4]
    d = pd.read_excel(xl, sheet_name=0, usecols='S')
    pit_date = d['Unnamed: 18'][4]

    # combine date and time into one datetime variable, and format
    pit_datetime=datetime.datetime.combine(pit_date, pit_time)
    pit_datetime_str=pit_datetime.strftime('%Y-%m-%dT%H:%M')
    s=str(Site) + ',' + str(UTMzone)+hsphere + ',' +  pit_datetime_str + ',' + str(UTME) + ',' + str(UTMN) + ',' + str(TotalDepth)

    # create minimal header info for other files
    index = ['# Location', '# Site', '# PitID', '# Date/Local Time',
         '# UTM Zone', '# Easting', '# Northing', '# Latitude', '# Longitude']
    column = ['value']
    df = pd.DataFrame(index=index, columns=column)
    df['value'][0] = Location
    df['value'][1] = str(Site)
    df['value'][2] = str(PitID)
    df['value'][3] = pit_datetime_str
    df['value'][4] = str(UTMzone)+hsphere # Only for Grand Mesa 2020!! - '12N'
    df['value'][5] = UTME
    df['value'][6] = UTMN
    df['value'][7] = LAT
    df['value'][8] = LON

    # add minimal header to each data file
    df.to_csv(fname_newSnow, sep=',', header=False)

    # get new snow (interval board data)
    d = pd.read_excel(xl, sheet_name=0, usecols='B:E')
    rIx = (d.iloc[:,0] == 'Interval board measurements\nUse SWE tube').idxmax() #
    d = d.iloc[rIx+4:, 2:].reset_index(drop=True) #four down from the interval board section
    d.columns = ['HN (cm)', 'SWE (mm)']
    d2=d.iloc[0:3].values
    d2=np.array(d2.flatten(order='F'))#, dtype=float) #flatten array to match csv style
    d3=d['HN (cm)'].iloc[3] #evidence of melt (y/n)
    d4= np.append(d2, d3) #combine HN and SWE array with Evidence of Melt
    columns = ['# HN (cm) A', 'HN (cm) B', 'HN (cm) C',
                'SWE (mm) A', 'SWE (mm) B', 'SWE (mm) C',
                'Evidence of Melt']
    # ~~~ ADD CODE BLOCK HERE ~~~: grab weather box, split at IB and save the last, with if statement for any PD comments -- Megan is still implementing this to the main script.
    newSnow=pd.DataFrame(d4.reshape(-1, len(d4)), columns=columns)
    newSnow.to_csv(fname_newSnow, sep=',', index=False, mode='a', na_rep='NaN')
    print('wrote: .../' + fname_newSnow.split('/')[-1])

#~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
if __name__ == "__main__":
    # static variables
    version = 'v01'
    hsphere = 'N' # northern hemisphere
    # paths
    path_in = Path('/Users/meganmason491/Google Drive/SnowEx-2020-timeseries-pits/timeseries_pitbook_sheets_EDIT/FOR_EDIT/')
    path_out = '/Users/meganmason491/Google Drive/SnowEx-2020-timeseries-pits/parameter_output_csv/'

    # loop over all pit sheets
    for filename in path_in.rglob('*.xlsx'):
        print(filename.name)
        r = readIntervalBoard(path_in, filename, version, path_out)

    print('..... Script is Complete !  .....')
