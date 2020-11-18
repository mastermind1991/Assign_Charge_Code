#!/usr/bin/python3

import snapicall
import requests
import sys
import os
import json
import xlsxwriter
import pandas as pd
from openpyxl import load_workbook
from datetime import datetime


## Search this document for "testing" to find all testing variable switches.

### Variables ###
# Default variable for later use if needed to keep log files organized    
today = datetime.now()
timestamp = today.strftime("%m/%d/%y")
filedate = today.strftime("%Y%m%d")
month = today.strftime("%m") 
day = today.strftime("%d")
year = today.strftime("%Y")
time = today.strftime("%H.%M")

##### Change the following to fit your needs #####
##Live Environment
base_api_url = "https://umchealthsystem.service-now.com/api/"

##Testing - Test Environment
#base_api_url = "https://umchealthsystemtest.service-now.com/api/"


# Temp file output directory
tempdir = sys.path[0]
#tempdir = r"C:\temp\\"

# Spreadsheet output
xldir = os.path.join(tempdir, "spreadsheets")
#xldir = r"C:\temp\assign_charge_code\spreadsheets"

# Log directory
logdir = os.path.join(tempdir, "logs")
# Log file output
logfile = os.path.join(logdir, filedate + "_" + time + "assign_charge_code.txt")
#logfile = r"C:\temp\assign_charge_code\logs"

#####################################################################

# Count if anything was done on the tickets return in the query
action_count = 0

## Class creation for bill item details:
class bill_item(object):
    def __init__(self, description, hcpcs_cpt, cdm, lawson_number, price_schedule, price, effective_date, charge_schedule, charge_level, charge_point):
        self.description = description
        self.hcpcs_cpt = hcpcs_cpt
        self.cdm = cdm
        self.lawson_number = lawson_number
        self.price_schedule = price_schedule
        self.price = price
        self.effective_date = effective_date
        self.charge_schedule = charge_schedule
        self.charge_level = charge_level
        self.charge_point = charge_point
class detail(bill_item):
    pass

# Function to itterate over the sc_description_list. The list will be created by using the newline variable with either \n or \r\n
def pd_df_creation(sc_task_number, newline, action_count):
    # Create empty list to add classes to
    charge_code_list = []
    df_description = []
    df_hcpcs_cpt = []
    df_cdm = []
    df_lawson_number = []
    df_price_schedule = []
    df_price = []
    df_effective_date = []
    df_charge_schedule = []
    df_charge_level = []
    df_charge_point = []
    # Create a list based on the new line delineator
    sc_description_list = sc_description.split(newline + newline)
    # Open a spread sheet with the SCTASK number in the name and the correct column headers. Here is where you would change the column headers.
    xldoc =  "charge_services_data_" + sc_task_number + ".xlsx"
    xlfile = os.path.join(xldir, xldoc)
    df = pd.DataFrame(data={'Description':[], 'Short Description(Lawson Number)':[], 'Schedule':[], 'Charge Level':[], 'Charge Point':[],'HCPCS':[],'CDM':[],'Price Schedule':[], 'Price':[], 'Begin Date':[]})
    writer = pd.ExcelWriter(xlfile, engine='xlsxwriter')
    df.to_excel(writer, index=False)
    writer.save()
    # Account for list change
    for item in sc_description_list[1:]:
        if "Charge Processing will be group complete on Charge Point - UMC Schedule." in item:
            break
        else:
            ## Include a potential null for all items
            item_list = item.split(newline)
            # Lawson Item number(?)
            lawson_number = item_list[4]
            try:
                lawson_number = lawson_number.split(': ',1)[1]
            except:
                lawson_number = ' '
            # Item/Orderable Description
            description = item_list[0]
            # Charge code/CDM
            cdm = item_list[3]
            try:
                cdm = cdm.split(': ',1)[1]
            except:
                cdm = ' '
            # CPT/HCPCS Code
            hcpcs_cpt = item_list[1]
            try: 
                hcpcs_cpt = hcpcs_cpt.split(': ',1)[1]
            except:
                hcpcs_cpt = ' '
            # Price Schedule - Always "Base Technical"
            price_schedule = "Base Technical"
            # Always $0
            price = "0"
            #start date
            effective_date = item_list[2]
            try:
                effective_date = effective_date.split(': ',1)[1]
            except:
                effective_date = ' '
            # Charge processing
            charge_schedule = "Charge Point - UMC"
            charge_level = "Group|Complete"
            charge_point = "Complete"
            charge_code_list.append(detail(description, hcpcs_cpt, cdm, lawson_number, price_schedule, price, effective_date, charge_schedule, charge_level, charge_point))
            # Append the spreadsheet created outside of the loop
    for n in range(len(charge_code_list)):
        df_description.append(charge_code_list[n].description)
        df_hcpcs_cpt.append(charge_code_list[n].hcpcs_cpt)
        df_cdm.append(charge_code_list[n].cdm)
        df_lawson_number.append(charge_code_list[n].lawson_number)
        df_price_schedule.append(charge_code_list[n].price_schedule)
        df_price.append(charge_code_list[n].price)
        df_effective_date.append(charge_code_list[n].effective_date)
        df_charge_schedule.append(charge_code_list[n].charge_schedule)
        df_charge_level.append(charge_code_list[n].charge_level)
        df_charge_point.append(charge_code_list[n].charge_point)
    writer = pd.ExcelWriter(xlfile, engine='openpyxl')
    writer.book = load_workbook(xlfile)
    writer.sheets = dict((ws.title, ws) for ws in writer.book.worksheets)
    reader = pd.read_excel(xlfile)
    # Append Data Frame
    #df = pd.DataFrame(data={'Description':[charge_code_list[n].description], 'Short Description(Lawson Number)':[charge_code_list[n].lawson_number], 'Schedule':[charge_code_list[n].price_schedule], 'Charge Level':[charge_code_list[n].charge_level], 'Charge Point':[charge_code_list[n].charge_point],'HCPCS':[charge_code_list[n].hcpcs_cpt],'CDM':[charge_code_list[n].cdm],'Price Schedule':[charge_code_list[n].price_schedule], 'Price':[charge_code_list[n].price], 'Begin Date':[charge_code_list[n].effective_date]})
    df = pd.DataFrame(data={'Description':df_description, 'Short Description(Lawson Number)':df_lawson_number, 'Schedule':df_price_schedule, 'Charge Level':df_charge_level, 'Charge Point':df_charge_schedule,'HCPCS':df_hcpcs_cpt,'CDM':df_cdm,'Price Schedule':df_price_schedule, 'Price':df_price, 'Begin Date':df_effective_date})
    df.to_excel(writer,index=False,header=False,startrow=len(reader)+1)
    writer.close()
    #writer.save()
    # Attachment Post
    snapicall.api_xlsx_post(content = open(xlfile, 'rb').read(), url = base_api_url + 'now/attachment/file?table_name=sc_task&table_sys_id=' + sc_task_sysid + '&file_name='+ xldoc)
    ## Removed; The spreadsheet has been compile in this location, so moving does not have to occur
    # Move the spreadsheet to a 'spreadsheet' file of somekind to keep track of the uploads
    #if not os.path.exists(xldir):
    #    os.makedirs(xldir)
    #newxlfile = xldir + "\\" + xldoc
    #if os.path.exists(newxlfile):
    #    os.remove(newxlfile)
    #os.rename(xlfile, xldir + "\\" + xldoc )
    # Log action
    action_count = action_count + 1
    with open(logfile, "a+") as f:
        f.write("Spreadsheet added to: " + sc_task_number + " at " + time + "\n")
        f.close()
    # Return the iteration number
    return action_count

# API Request to retrieve the SCTASK with the description "Assign Charge Code to Item in Cerner PROD" and sctask state of open, pending, or work in progress; Note: %20 dilineates a space in a url
snreply = snapicall.api_get(url = base_api_url + str('now/table/sc_task?sysparm_query=short_description=Assign%20Charge%20Code%20to%20Item%20in%20Cerner%20PROD^stateIN-5,2,1'))
# Testing Uncommment to test with SCTASK0141977 (Don't forget to flip the ticket state, this ticket is closed)
#snreply = snapicall.api_get(url = base_api_url + str('now/table/sc_task?sysparm_query=number=SCTASK0141977'))
# Count the number of tickets in the reply result for itteration number
ticketcnt = len(snreply["result"])
# if the result list has more than "0" entrys; Prevents error in script due to empty response
if ticketcnt >= 1:
    for i in range(ticketcnt):
        # SCTASK sysid; Table primary key
        sc_task_sysid = snreply["result"][i]["sys_id"]
        # SCTASK State; Note: 1 = Open, 2 = Work in Progress, 3 = Closed Complete, 4 = Closed Incomplete, -5 = Pending, -9 = Validation
        sc_task_state = snreply["result"][i]["state"]
        # Actual SCTASK Number
        sc_task_number = snreply["result"][i]["number"]
        # SCTASK description
        sc_description = snreply["result"][i]["description"]
        # Check to make sure the ticket is open
        # Add an "or" statement
        if sc_task_state == '1' or sc_task_state == '2' or sc_task_state == '-5':
        # Testing - Uncomment for testing
        #if sc_task_state == '3':
            # Check to make sure there is something useful in the short description
            if sc_description not in 'View the request variables for the details of this request.':             
                #### Look for attachments ####
                attachmentreply = snapicall.api_get(url = base_api_url + 'now/attachment?sysparm_query=table_name=sc_task^table_sys_id=' + sc_task_sysid)
                # Count attachments for itteration number
                attachment_cnt = len(attachmentreply["result"])
                # If there are attachments check to see if they are spreadsheets. If a spreadsheet exists, do nothing. If a spreadsheet does not exist, create one
                tempfile = tempdir + "assignchargetemp" + str(i) + ".xlsx"
                
                # Create Pandas Data Frame    
                if attachment_cnt >= 1:
                    for a in range(attachment_cnt):
                        # Change the attachment check to look for the exact spreadsheet name
                        attachmentname = attachmentreply["result"][a]["file_name"]
                        if 'charge_services_data_' in attachmentname:           
                            attachmentbln = True
                            if not os.path.exists(logdir):
                                os.makedirs(logdir)
                            with open(logfile, "a+") as f:
                                f.write("Spread sheet already exists on: " + sc_task_number + "\n")
                                f.close()
                            #print("Already has spread sheet")
                        else:
                            attachmentbln = False
                else:
                    attachmentbln = False
                if attachmentbln == False:
                    #### Excel Data Frame Build ####
                    # Build class for each bill item; leverage double new line to create item list, then create class
                    # Account for manual edits that may change the coded "new line" from \n and \n\n to \r\n and \r\n\r\n
                    if '\r\n' in sc_description:
                        pd_df_creation(sc_task_number = sc_task_number, newline = '\r\n', action_count = action_count)
                    elif '\n\n' in sc_description:
                        pd_df_creation(sc_task_number = sc_task_number, newline = '\n', action_count = action_count)              
                    else:
                        # Write 'No bill items on sc_task_number' to a text file
                        # Open text file and write logs to it
                        if not os.path.exists(logdir):
                            os.makedirs(logdir)
                        with open(logfile, "a+") as f:
                            f.write("No bill items found on: " + sc_task_number + "\n")
                            f.close()
else:
    if not os.path.exists(logdir):
        os.makedirs(logdir)
    with open(logfile, "a+") as f:
        f.write("Could not find any tickets with that description\n")
        f.close()
if action_count == 0:
    if not os.path.exists(logdir):
        os.makedirs(logdir)
    with open(logfile, "a+") as f:
        f.write("No tickets are open or on hold with that description\n")
        f.close()
