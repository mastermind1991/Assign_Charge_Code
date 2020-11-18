#!/usr/bin/python3

from datetime import datetime
import requests
import json
import yaml

####### Variables ######################################
# Time stamp
today = datetime.now()
date = today.strftime("%Y%m%d")
timestamp = today.strftime("%Y%m%d_%H.%M")
# Log File Location
jsonlogfile = r'C:\temp\\' + date + "jsonlog.txt"
# API Key
## Place the location of the API password file here:
u_info = yaml.load(open(r"C:\temp\script_info\sn_api.yml"), Loader=yaml.FullLoader)
username = u_info['user']['username']
password = u_info['user']['password'] # Use the following format: '(username,password)'

# Headers
headers = {'Accept' : 'application/json', 'Content-Type' : 'application/json'}


########################################################
## Uncomment the following for large JSON files:
#def api_post(sctask_json, ritm_json):
#    r = requests.post(api_endpoint, data=open(ritm_json, 'rb'), headers=headers)
#######################################################

#######################################################
## Uncomment the following for smaller json files:
def api_xlsx_post(content, url):
    ## Request ##
    xlsx_headers = {'Accept' : 'application/json', 'Content-Type' : 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'}
    response = requests.post(url, auth=(username, password), headers=xlsx_headers, data=content)
    # Check for HTTP codes other than 200
    if response.status_code != 200: 
        print('Status:', response.status_code, 'Headers:', response.headers, 'Error Response:',response.json())

    # Decode the JSON response into a dictionary and use the data
    data = json.loads(response.text)
    print(data)
    print("\n")
    # Log it
    #with open(jsonlogfile, "a+") as j:
    #    j.write("API Post at " +  timestamp + "\n")
    #    j.write(data + "\n")
    #    j.close()
    return data["result"]["sys_id"]
def api_post(content, url):
    ## Request ##
    response = requests.post(url, auth=(username, password), headers=headers, data=content)
    # Check for HTTP codes other than 200
    if response.status_code != 200: 
        print('Status:', response.status_code, 'Headers:', response.headers, 'Error Response:',response.json())

    # Decode the JSON response into a dictionary and use the data
    data = json.loads(response.text)
    print(data)
    print("\n")
    # Log it
    #with open(jsonlogfile, "a+") as j:
    #    j.write("API Post at " +  timestamp + "\n")
    #    j.write(data + "\n")
    #    j.close()
    return data["result"]["sys_id"]

def api_put(content, url):
    ## Request ##
    response = requests.put(url, auth=(username, password), headers=headers, data=content)
    # Check for HTTP codes other than 200
    if response.status_code != 200: 
        print('Status:', response.status_code, 'Headers:', response.headers, 'Error Response:',response.json())

    # Decode the JSON response into a dictionary and use the data
    data = json.loads(response.text)
    print(data)
    print("\n")
    # Log it
    #with open(jsonlogfile, "a+") as j:
    #    j.write("API Put at " +  timestamp + "\n")
    #    j.write(data + "\n")
    #    j.close()

def api_get(url):
    ## Request ##
    response = requests.get(url, auth=(username, password), headers=headers)
    # Check for HTTP codes other than 200
    if response.status_code != 200: 
        print('Status:', response.status_code, 'Headers:', response.headers, 'Error Response:',response.json())

    # Decode the JSON response into a dictionary and use the data
    data = json.loads(response.text)
    print(data)
    print("\n")
    # Log it
    #with open(jsonlogfile, "a+") as j:
    #    j.write("API Get at " +  timestamp + "\n")
    #    j.write(data + "\n")
    #    j.close()
    return data
    # Original query
    #return data["result"][0]["sys_id"]


########################################################

#if __name__ == "__main__":
#    sctask_json = r"C:\temp\snticket.json"
#    ritm_json = r"C:\temp\ritmticket.json"
#    api_post(sctask_json = sctask_json, ritm_json = ritm_json)
