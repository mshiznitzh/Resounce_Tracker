#!/usr/bin/env python3
"""
Module Docstring
"""

__author__ = "MiKe Howard"
__version__ = "0.1.0"
__license__ = "MIT"



# Importing the matplotlb.pyplot
import matplotlib.pyplot as plt
import logging
import os
import pandas as pd
from datetime import date, datetime, timedelta
from dateutil.relativedelta import *
import numpy as np


#OS Functions
def filesearch(word=""):
    """Returns a list with all files with the word/extension in it"""
    logger.info('Starting filesearch')
    file = []
    for f in glob.glob("*"):
        if word[0] == ".":
            if f.endswith(word):
                file.append(f)

        elif word in f:
            file.append(f)
            #return file
    logger.debug(file)
    return file

def Change_Working_Path(path):
    # Check if New path exists
    if os.path.exists(path):
        # Change the current working Directory
        try:
            os.chdir(path)  # Change the working directory
        except OSError:
            logger.error("Can't change the Current Working Directory", exc_info = True)
    else:
        print("Can't change the Current Working Directory because this path doesn't exits")

#Pandas Functions
def Excel_to_Pandas(filename):
    logger.info('importing file ' + filename)
    df=[]
    try:
        df = pd.read_excel(filename)
    except:
        logger.error("Error importing file " + filename, exc_info=True)

    df=Cleanup_Dataframe(df)
    logger.debug(df.info(verbose=True))
    return df

def Cleanup_Dataframe(df):
    logger.info('Starting Cleanup_Dataframe')
    logger.debug(df.info(verbose=True))
    # Remove whitespace on both ends of column headers
    df.columns = df.columns.str.strip()

    # Replace whitespace in column header with _
    df.columns = df.columns.str.replace(' ', '_')

    return df



def make_gnat(y_min, y_max, x_min, x_max, df):

    df['Planned_Start']= df['Planned_Start'].dt.date
    df['Planned_Finish'] = df['Planned_Finish'].dt.date

    df = df.sort_values(by=['Planned_Start'])

    df = df[(df['Activity_Other_Resources'] == ' Ft. Worth P&amp;C Crews' )&
            (df['Planned_Finish'] > pd.Timestamp(datetime.now()) ) &
            (df['Planned_Finish'] < pd.Timestamp(datetime.now() + relativedelta(months=+6)) )]


    labels=df.Project_ID.apply(str) + ' - ' + df.Activity_Description.apply(str)
    length=len(df.index)
    ticks=[]

    for x in range(length):
        ticks.append((x+1)/length*100)
    #plt.rcParams['ytick.major.pad'] = ticks[0]
    # Declaring a figure "gnt"
    fig, gnt = plt.subplots(figsize=(18, 20))

    # Setting Y-axis limits
    gnt.set_ylim(y_min, y_max)

    # Setting X-axis limits

    #gnt.set_xlim(df['Planned_Start'].min(), '2020-12-31')
    gnt.set_xlim(date.today(), datetime.now() + relativedelta(months=+6))
    #gnt.set_xlim('2020-08-01', df['Planned_Start'].max())

    # Setting labels for x-axis and y-axis
    #gnt.set_xlabel('')
    #gnt.set_ylabel('Date')
    #gnt.autoscale(enable=True, axis='x')
    #gnt.autoscale(enable=True, axis='y')
    # Setting ticks on y-axis
    gnt.set_yticks(ticks)
    #gnt.rcParams['ytick.major.pad'] = ticks[0]
    # Labelling tickes of y-axis
    gnt.set_yticklabels(labels)

    # Setting graph attribute
    gnt.grid(True)
    gnt.xaxis_date()

    # Declaring a bar in schedule
    #gnt.barh([ticks[0]-ticks[0]/2, (df.Planned_Finish.values[0] - df.Planned_Start.values[0]), left=df.Planned_Start.values[0], height=ticks[0], align='center', color='orange', alpha = 0.8)
    for x in range(len(df.index)):
        if pd.notnull(df.Actual_Start.values[x]):
            gnt.barh(ticks[x], (df.Planned_Finish.values[x] - df.Planned_Start.values[x]), left=df.Planned_Start.values[x], height=ticks[0]/2, align='center', color='maroon', alpha = 0.8)
        else:
            gnt.barh(ticks[x], (df.Planned_Finish.values[x] - df.Planned_Start.values[x]), left=df.Planned_Start.values[x],
                 height=ticks[0] / 2, align='center', color='red', alpha=0.8)

    # Declaring multiple bars in at same level and same width
    #gnt.broken_barh([(110, 10), (150, 10)], (10, 9),
                   # facecolors='tab:blue')

    #gnt.broken_barh([(10, 50), (100, 20), (130, 10)], (20, 9),
                 #   facecolors=('tab:red'))
    fig.autofmt_xdate()
    plt.tight_layout()
    plt.savefig("Fort Worth.png")

def main():
    Project_Data_Filename='PETE Schedules Export - 08-13-2020.xlsx'

    """ Main entry point of the app """
    logger.info("Starting Pete Maintenance Helper")
    Change_Working_Path('./Data')
    try:
        Project_Data_df = Excel_to_Pandas(Project_Data_Filename)
    except:
        logger.error('Can not find Project Data file')
        raise

    make_gnat(0, 100, 180, 180, Project_Data_df)

if __name__ == "__main__":
    """ This is executed when run from the command line """
    # Setup Logging
    logger = logging.getLogger('root')
    FORMAT = "[%(filename)s:%(lineno)s - %(funcName)20s() ] %(message)s"
    logging.basicConfig(format=FORMAT)
    logger.setLevel(logging.DEBUG)

    main()