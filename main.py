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
import xlsxwriter
#from slugify import slugify
from sanitize_filename import sanitize


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



def make_gnat(df, title):

    labels=df.PETE_ID.apply(str) + ' - ' + df.Grandchild.apply(str)
    length=len(df.index)
    ticks=[]

    for x in range(length):
        ticks.append((x+1)/length*100)

    # Declaring a figure "gnt"
    fig, gnt = plt.subplots(figsize=(19, 15))

    fig.suptitle(title, fontsize=16)

    # Setting Y-axis limits
    gnt.set_ylim(0, ticks[-1]+ticks[0])

    # Setting X-axis limits

    gnt.set_xlim(date.today(), datetime.strptime('2020-12-31',"%Y-%m-%d"))

    # Setting labels for x-axis and y-axis

    # Setting ticks on y-axis
    gnt.set_yticks(ticks)

    # Labelling tickes of y-axis
    gnt.set_yticklabels(labels)

    # Setting graph attribute
    gnt.grid(True)
    gnt.xaxis_date()

    # Declaring a bar in schedule
    #gnt.barh([ticks[0]-ticks[0]/2, (df.Finish_Date.values[0] - df.Start_Date.values[0]), left=df.Start_Date.values[0], height=ticks[0], align='center', color='orange', alpha = 0.8)
    for x in range(len(df.index)):
        if df['Start_Date_Planned\Actual'].values[x] == 'A':
            gnt.barh(ticks[x], (df.Finish_Date.values[x] - df.Start_Date.values[x]),
                     left=df.Start_Date.values[x], height=ticks[0]/2, align='center', color='maroon', alpha = 0.8)
        else:
            gnt.barh(ticks[x], (df.Finish_Date.values[x] - df.Start_Date.values[x]),
                     left=df.Start_Date.values[x],
                 height=ticks[0] / 2, align='center', color='red', alpha=0.8)

    fig.autofmt_xdate()
    plt.tight_layout()
    plt.savefig('Output/'+sanitize(title)+'.png')

def identify_date_with_over_allocation(df, Resource, list):

    daterange = pd.date_range(datetime.today(), datetime.strptime('2020-12-31',"%Y-%m-%d"))
    # Lets find days that we have over allocated
    column_names = ['Date', 'Count']
    Utilizationdf = pd.DataFrame(columns=column_names)
    for single_date in daterange:
        holddf = df[(df['Start_Date'] <= single_date.date()) &
                        (df['Finish_Date'] >= single_date.date())]
        Utilizationdf = Utilizationdf.append({'Date': single_date.date(), 'Count': len(holddf)}, ignore_index=True)
    mean=round(np.mean(Utilizationdf['Count'], 0))
    print(' '.join(['minimum = ', str(round(np.amin(Utilizationdf['Count'], 0)))]))
    print(' '.join(['mean =' , str(mean)]))
    print(' '.join(['maximum = ', str(round(np.amax(Utilizationdf['Count'], 0)))]))
    std=round(np.std(Utilizationdf['Count'], 0))
    print(' '.join(['Standard Deviation =',str(std)]))
    Overutilizationdf=Utilizationdf[(Utilizationdf['Count'] > mean + std) ]
    underutilizationdf = Utilizationdf[(Utilizationdf['Count'] < mean - std)]

    print('Allocation is higher than mean + 1 standard deviation for '+ df.iloc[0].Other_Activity_Resource + ' on the following dates:')
    print(Overutilizationdf)
    Overutilizationdf.to_csv(sanitize('_'.join([str(Resource), str(list), 'Overutilizationdf.csv'])))
    print('Allocation is lower than mean - 1 standard deviation for ' + df.iloc[
        0].Other_Activity_Resource + ' on the following dates:')
    underutilizationdf.to_csv(sanitize('_'.join([str(Resource), str(list), 'underutilizationdf.csv'])))
    print(underutilizationdf)


def print_district_stats(planneddf, resourcemissingdf, district):
    print('Currently ' + district + ' has ' + str(len(planneddf)) + ' planned activities on the district schedule with '
          + str(len(resourcemissingdf)) + ' activities missing a resource. ' + 'Thus ' +
          str(round((1 - len(resourcemissingdf) / len(planneddf)) * 100,
                    0)) + '% of activities have a resource selected in PETE.')

def filter_Prject_Data_By_Schedule(district ,Project_Data_df):
    planneddf = Project_Data_df[(Project_Data_df['Schedule_Function'] == 'District') &
                                (Project_Data_df['Work_Center_Name'] == district) &
                                (Project_Data_df['Grandchild'] != 'D - Construction Ready') &
                                (Project_Data_df['Grandchild'] != 'Project Energization') &
                                # (Project_Data_df['Start_Date_Planned\Actual'] != 'A') &
                                (Project_Data_df['Finish_Date_Planned\Actual'] != 'A')]

    resourcemissingdf = Project_Data_df[(Project_Data_df['Schedule_Function'] == 'District') &
                                        (Project_Data_df['Work_Center_Name'] == district) &
                                        (Project_Data_df['Grandchild'] != 'D - Construction Ready') &
                                        (Project_Data_df['Grandchild'] != 'Project Energization') &
                                        # (Project_Data_df['Start_Date_Planned\Actual'] != 'A') &
                                        (Project_Data_df['Finish_Date_Planned\Actual'] != 'A') &
                                        (pd.isnull(Project_Data_df['Other_Activity_Resource']))]

    return planneddf, resourcemissingdf;


def find_the_counstrunction_Season(date_input):

    if pd.Timestamp(date_input).quarter < 3:
        Last_Day_of_Season = date(datetime.today().year, 6, 30)

    else:
        Last_Day_of_Season = date(datetime.today().year, 12, 31)

    return Last_Day_of_Season

def main():
    Project_Data_Filename='Metro West PETE Schedules.xlsx'

    """ Main entry point of the app """
    logger.info("Starting Resource Tracker")
    Change_Working_Path('./Data')
    try:
        Project_Data_df = Excel_to_Pandas(sanitize(Project_Data_Filename))
    except:
        logger.error('Can not find Project Data file')
        raise

    Project_Data_df['Start_Date'] = Project_Data_df['Start_Date'].dt.date
    Project_Data_df['Finish_Date'] = Project_Data_df['Finish_Date'].dt.date

    Project_Data_df = Project_Data_df.sort_values(by=['Start_Date'])

    #Stats about Other Activity Resource

    districts=['AMARILLO', 'BIG SPRING','FORT WORTH', 'GRAHAM','ODESSA', 'SWEETWATER', 'WICHITA FALLS']
    #districts = ['FORT WORTH']
    #districts = planneddf.Other_Activity_Resource.dropna().unique()
    #list=['Without Assumptions', 'With Assumptions']
    list=['']
    for item in list:
        for district in districts:

            if item == 'With Assumptions':
                Project_Data_df.loc[((Project_Data_df['Grandchild'] == 'Electrical Job Planning') &
                                    (pd.isnull(Project_Data_df['Other_Activity_Resource'])) &
                                    (Project_Data_df['Work_Center_Name'] == district)), ['Other_Activity_Resource']] = 'Ft. Worth P&C Crews'

                Project_Data_df.loc[(Project_Data_df['Grandchild'] == 'Electrical Construction') &
                                    (pd.isnull(Project_Data_df['Other_Activity_Resource'])) &
                                    (Project_Data_df['Work_Center_Name'] == district), ['Other_Activity_Resource']] = 'Ft. Worth P&C Crews'

            planneddf, resourcemissingdf = filter_Prject_Data_By_Schedule(district, Project_Data_df)


            #if item == 'Without Assumptions':
            resourcemissingdf.to_excel(' '.join([district,'Activities missing resources.xlsx']), district, index=False, engine='xlsxwriter')

            print_district_stats(planneddf, resourcemissingdf, district)

            for Resource in planneddf.Other_Activity_Resource.dropna().unique():
                print(Resource)
                DATADF = planneddf[(planneddf['Other_Activity_Resource'] == Resource) &
                    (planneddf['Finish_Date'] > pd.Timestamp(datetime.now())) &
                    (planneddf['Finish_Date'] <= pd.Timestamp(datetime.strptime(str(find_the_counstrunction_Season(datetime.today())) ,"%Y-%m-%d")))]

                if DATADF.size > 0:
                    make_gnat(DATADF, ' '.join([DATADF['Other_Activity_Resource'].values[0],'Utilization', item]))
                    identify_date_with_over_allocation(DATADF, Resource, item)




if __name__ == "__main__":
    """ This is executed when run from the command line """
    # Setup Logging
    logger = logging.getLogger('root')
    FORMAT = "[%(filename)s:%(lineno)s - %(funcName)20s() ] %(message)s"
    logging.basicConfig(format=FORMAT)
    logger.setLevel(logging.DEBUG)

    main()