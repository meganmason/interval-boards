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
import matplotlib.pyplot as plt
import seaborn as sns
# sns.set_theme(style="whitegrid")

'''
Some workflow to get started with the 2020 Interval Board dataset. Script was copy/pasted from ts_pits_xls2csv.py script to process snow pit data.
A lot of irrelevant script was removed, but it has not be thoroughly edited for variables that do not pertain to the interval board data set.

Eventually the input files should be the published Time Series snow pit (.xlsx) files.
For right now, grabbing files from the timeseries_pitbook_sheets_EDIT folder is fine (however, please make a copy and do not work from that main folder).

What's the difference?  - the published files will have fixed coordinates for the Snow Pits (same coords will be used for the interval board for 2020),
whereas the files to start with the coordinates will vary, but that is OKAY to get started.

The parameter named assigned to the Interval Board dataset = newSnow
'''

def readIntervalBoard(path_in, filename, version, path_out, fname_newSnow):


    # open excel file
    xl = pd.ExcelFile(filename) # opens xlxs workbook

    # get new snow (interval board data)
    d = pd.read_excel(xl, sheet_name=0, usecols='B:E')
    rIx = (d.iloc[:,0] == 'Interval board measurements\nUse SWE tube').idxmax() #
    d = d.iloc[rIx+4:, 2:].reset_index(drop=True) #four down from the interval board section
    d.columns = ['HN (cm)', 'SWE (mm)']

    if d['HN (cm)'][0] >=0: # includes zeros!

        hgtA = d['HN (cm)'][0] # height of Sample A
        hgtB = d['HN (cm)'][1] # height of Sample B
        hgtC = d['HN (cm)'][2] # height of Sample C
        sweA = d['SWE (mm)'][0] # swe of Sample A
        sweB = d['SWE (mm)'][1] # swe of Sample B
        sweC = d['SWE (mm)'][2] # swe of Sample C

        # compute depth and swe means
        avg = round(d.iloc[0:3].mean(skipna=True), 1) # averages 3 samples, skips all nans, and rounds to 1 decimal place

        # evidence of melt
        Melt = d['HN (cm)'].iloc[3] # Yes / No string

        # comments: pull from 'weather comments' box
        d = pd.read_excel(xl, sheet_name=0, usecols='B:M')
        rIx = (d.iloc[:,0] == 'Weather:').idxmax() #locate 'Weather:' cell in spreadsheet (row Index)
        d = d.loc[rIx:,:].reset_index(drop=True) # subset dataframe from 'Weather:' cell down to bottom, and reset index (not always fixed due to extra rows in above measurements)
        Weather = str(d['Location:'][1]) # grab Weather comment box

        if "IB: " in Weather:
            Comments = Weather.split('IB: ')[1] # splits at 'IB: ' and saves 2nd return
            if "PD: " in Comments:
                Comments = Comments.split('PD: ')[0] # splits at 'PD: ' and saves first (removes any perimeter depth comments)
        else:
            Comments = None

        # date and time
        d = pd.read_excel(xl, sheet_name=0, usecols='S')
        pit_date = d['Unnamed: 18'][4] # grab snow pit date
        d = pd.read_excel(xl, sheet_name=0, usecols='X')
        pit_time = d['Unnamed: 23'][4] # grab snow pit time


        # location / pit name
        d = pd.read_excel(xl, sheet_name=0, usecols='B')
        Location = d['Location:'][0]
        Site = d['Location:'][2]  # grab site name
        if isinstance(d['Location:'][4], float):
            PitID = filename.stem
        else:
            PitID = d['Location:'][4][:6]  # grab pit name
        d = pd.read_excel(xl, sheet_name=0, usecols='L')
        UTME = int(d['Observers:'][2])# grab UTME
        d = pd.read_excel(xl, sheet_name=0, usecols='Q')
        UTMN = int(d['Unnamed: 16'][2]) # grab UTMN
        d = pd.read_excel(xl, sheet_name=0, usecols='X')
        UTMzone = int(d['Unnamed: 23'][2]) # for northern hemisphere

        # combine date and time into one datetime variable, and format
        pit_datetime=datetime.datetime.combine(pit_date, pit_time)

        # create dictionary to store data
        dict = {}
        dict = {'Easting': UTME,
                'Northing': UTMN,
                'Zone': UTMzone,
                'Date/Local Standard Time': pit_datetime,
                'Week No.': pit_datetime.isocalendar()[1]-1, #returns tuple, want 2nd item (-1 because campaign started in teh 2nd week?)
                'PitID': PitID,
                'State': PitID[:2],
                'Location': Location,
                'Site': Site,
                'HN (cm) A': hgtA,
                'HN (cm) B': hgtB,
                'HN (cm) C': hgtC,
                'AVG HN (cm)': avg['HN (cm)'],
                'SWE (mm) A': sweA,
                'SWE (mm) B': sweB,
                'SWE (mm) C': sweC,
                'AVG SWE (mm)': avg['SWE (mm)'],
                'Melt': Melt,
                'Comments': Comments,}

        # grow dictionary
        dict.update(dict)

        # grow df rows
        rows_list.append(dict)

#~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
if __name__ == "__main__":

    # static variables
    version = 'v01' # working version
    hsphere = 'N' # northern hemisphere

    # paths (path1 = all boards that have snow pits, path2 is for sites with only boards (i.e. no published snow pits))
    path1 = Path('/Users/mamason6/Google Drive/My Drive/SnowEx-2020/SnowEx-2020-timeseries-pits/timeseries_pitbook_sheets_EDIT/COMPLETE') # interval boards that have a corisponding pit
    path2 = Path('/Users/mamason6/Google Drive/My Drive/SnowEx-2020/SnowEx-2020-timeseries-interval-boards/interval-boards-only-sheets') # interval boards independent of pit
    path_out = Path('/Users/mamason6/Documents/snowex/core-datasets/ground/interval-boards/output')

    # create output filename
    fname_newSnow = path_out.joinpath('SNEX20_TS_IB_Summary_newSnow_' + version +'.csv')
    fname_newSnow_sites_avg = path_out.joinpath('SNEX20_TS_IB_Summary_newSnow_sites_avg_' + version +'.csv')
    fname_newSnow_locations_avg = path_out.joinpath('SNEX20_TS_IB_Summary_newSnow_locations_avg_' + version +'.csv')

    # empty DataFrame
    df = pd.DataFrame()

    # empty rows list
    rows_list = []

    # loop over all pit sheets
    for i, filename in enumerate(sorted(path1.rglob('*.xlsx'))):
        print(i, 'running file: .../', filename.name)
        r = readIntervalBoard(path1, filename, version, path_out, fname_newSnow)

    for i, filename in enumerate(sorted(path2.rglob('*.xlsx'))):
        print(i, 'running file: .../', filename.name)
        r = readIntervalBoard(path2, filename, version, path_out, fname_newSnow)

    # create dataframe from rows
    df = pd.DataFrame(rows_list)

    '''
    Any necessary corrections based on Interval Board Comments handled here
    (i.e. adjusting dataframe if something doesn't align from pit sheet)
    '''

    # issue 1 ~~~~
    # Cameron Pass Joe Wright, coords are 1km from snow pit site, fix is below
    df.loc[df['PitID']=='COCPJW','Easting'] =  424807
    df.loc[df['PitID']=='COCPJW','Northing'] = 4487254


    # issue 2 ~~~~
    # Mores Creek summit - MCS site is by snotel, 2/12 and 3/11 grab pit location which is not acurate.
    df.loc[df['PitID']=='IDBRMC','Easting'] =  607094
    df.loc[df['PitID']=='IDBRMC','Northing'] = 4865196



#~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
# convert and add Lat / Lon to dataframe (better done after 'issues' section)
def getLat(row): # use utm python package to convert, returns a tuple, save the first
    return round(utm.to_latlon(row['Easting'], row['Northing'], row['Zone'], "Northern")[0], 5)

def getLon(row): # returns a tuple, save the second
    return round(utm.to_latlon(row['Easting'], row['Northing'], row['Zone'], "Northern")[1], 5)


df['Latitude'] = df.apply(getLat, axis=1)
df['Longitude'] = df.apply(getLon, axis=1)

# reorder df columns (put Lat/Lon in front [lat overwrite's lon's position])
cols = list(df)
cols.insert(0, cols.pop(cols.index('Longitude')))
cols.insert(0, cols.pop(cols.index('Latitude')))
df = df.loc[:, cols]



#~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
# average coords. use groupby
#Groupby one column and return the mean of only particular column in the group.
avg_site_df = df.groupby('Site')[['Latitude', 'Longitude']].mean().reset_index()
avg_loc_df = df.groupby('Location')[['Latitude', 'Longitude']].mean().reset_index()

#~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
# saving outputs

# convert to csv and save
df.to_csv(fname_newSnow, sep=',', index=False, na_rep=' ')
avg_site_df.to_csv(fname_newSnow_sites_avg, sep=',', index=False, na_rep=' ')
avg_loc_df.to_csv(fname_newSnow_locations_avg, sep=',', index=False, na_rep=' ')

    # save pickel df
    # df.to_pickle(path_out + '/SNEX20_TS_IB_newSnow_' + version + '.pkl')

print('..... Script is Complete !  .....')
