'''
Raya Rabu Automation V2
- Generates an excel file with the contents of all tickets created within a week of the day before this script is runned. (ex, today is 25th, script will get tickets from 17th to 24th)

Jira API credential and token - Available at the following link https://id.atlassian.com/manage-profile/security/api-tokens
- Username
- API Token

Dependencies
- python (duh)
- requests
- xlsxwriter
- API Token from a JIRA account.

Notes:
- Following code is made to provide data collection and format automation using Jira API and python. Requires no input and can be modified as required.
- Script can be run manually or by using Windows Task scheduler (should run on Linux too, but no promises)
- i know the code looks ass but it works okay.

TODO:
- Set checks for missing data
- cleanup code
- pound sand

'''
import requests
from requests.auth import HTTPBasicAuth
import json
from datetime import timedelta, date, datetime
import xlsxwriter
from collections import Counter

def data_acq():
    # Jira instance details
    jira_url = "<REPLACE ME>"

    # Authentication - using an API token (create one from your Jira account settings)
    email = "<REPLACE ME>"  # Replace with your jira account email
    api_token = "<REPLAC ME>" # Replace with your API token

    # URL for Jira search endpoint with JQL query
    url = f"{jira_url}/rest/api/3/search"

    # Date range for filtering issues (use format YYYY-MM-DD)
    start_date, end_date = getTimeRange()

    # Project key - Replace as required if uneeded then remove from "jql_query" variable
    project_key = "<REPLAC ME>"

    # Set JQL query to find the correct project within a certain time range.
    jql_query = f"project = {project_key} AND created >= '{start_date}' AND created <= '{end_date}'"

    # Set up the headers for the request
    headers = {
        "Accept": "application/json",
        "Content-Type": "application/json",
    }

    # Define all required fields. Following fields are taken from Jira, replace as required.
    # Note: This is not used, but can be integrated to reduce size of data from Jira. 
    fields = [
        "created",           # Timestamp
        "key",               # SOC Key
        "reporter",          # Reporter 
        "customfield_10681", # PIC
        "summary",           # Summary
        "customfield_10592", # Source IP
        "customfield_10593", # Destination IP
        "customfield_10906", # Details
        "priority",          # Priority 
        "status",            # Status 
        "labels",            # Labels          
        "customfield_10892", # Category
    ]

        # Define the parameters for the search
    params = {
        "jql": jql_query,           # Use the JQL query to filter by created date
        #"fields": ",".join(fields),
        "maxResults": 1000,         # Maximum number of issues to return (adjust as needed)
    }

    # Make the GET request to Jira API
    response = requests.get(url, headers=headers, auth=HTTPBasicAuth(email, api_token), params=params)

    # Check if the request was successful
    if response.status_code == 200:
        issues_data = response.json()  # Parse the JSON response
        issues = issues_data.get("issues", [])
        
        #print(issues_data) #Note, if you are unsure what fields may exist in your Jira project uncomment this print statement and find the required fields from the resulting JSON.
        #exit()

        if issues:
            print(f"Found {len(issues)} issues between {start_date} and {end_date - timedelta(1)}:")

            full_data_array = []

            # Below is a list of checks to make sure resulting fields are not empty. 
            # Note: As this field is REALLY dependent on the received JSON, manual modification is required. goodluck.
            for issue in issues:
                curr_data_array = []

                # Time Created 
                #print(f"TimeStamp: {issue['fields']['created']}")
                timeCreated = datetime.strptime(str(issue['fields']['created']), "%Y-%m-%dT%H:%M:%S.%f+0700")   #date format, change as required
                curr_data_array.append(timeCreated.strftime("%d/%m/%Y %H:%M"))

                # Issue Key
                #print(issue["key"])  # Print each issue key
                curr_data_array.append(issue["key"])
                
                # Reporter
                reporter_itsec = issue['fields']['reporter']
                #print(reporter_itsec['displayName'])
                curr_data_array.append(reporter_itsec['displayName'])
                
                # PIC 
                PIC_raya = issue['fields']['customfield_10681']
                if PIC_raya == None:
                    curr_data_array.append("N/A")   #Append the string NA if field is null.
                else:
                    #print(PIC_raya[0]["displayName"])
                    curr_data_array.append(PIC_raya[0]['displayName'])

                # Summary
                details_attack = issue['fields']['summary']
                #print(details_attack)
                curr_data_array.append(details_attack)

                # Source IP
                source_ip = issue['fields']['customfield_10592']
                #print("Source IP: " + str(source_ip))
                curr_data_array.append(str(source_ip))

                # Destination IP
                destination_ip = issue['fields']['customfield_10593']
                #print("Destination IP: " + str(destination_ip))
                curr_data_array.append(str(destination_ip))

                # Details
                details_manual = issue['fields']['customfield_10906']
                if details_manual == None:
                    #print("N/A")
                    curr_data_array.append("N/A")
                else:
                    #print(details_manual["content"][0]["content"][0]["text"])
                    curr_data_array.append(details_manual["content"][0]["content"][0]["text"])
          
                # Priority
                priority = issue['fields']['priority']['name']
                #print(priority)
                curr_data_array.append(priority)

                # Status
                status = issue['fields']['status']['name']
                #print(status)
                curr_data_array.append(status)

                # Labels
                status = issue['fields']['labels']
                if len(status) !=0:
                    #print(status[0])
                    curr_data_array.append(status[0])
                else:
                    #print("N/A")
                    curr_data_array.append("N/A")

                # Category
                status = issue['fields']['customfield_10892']['value']
                #print(status)
                curr_data_array.append(status)

                #NOTE make checks on all acqquiry just in case

                full_data_array.append(curr_data_array)
        else:
            print(f"No issues found between {start_date} and {end_date}.")
    else:
        print(f"Failed to retrieve issues. Status code: {response.status_code}")
        print(response.text)  # Print error message if any

    #Output acquired data into an excel file.
    excelOutput(list(reversed(full_data_array)))

# Helper function to get date ranges
def getTimeRange():
    currRabu = date.today() #- timedelta(1)
    prevRabu = date.today() - timedelta(8)
    
    return prevRabu, currRabu

# Helper function to generate Excel file with required format (i know it looks ass, but it works and clients are happy with it so iiwis)
def excelOutput(dataIn):
    # Init Var
    listMonths = ["Januari", "Februari", "Maret", "April", "Mei", "Juni", "Juli", "Augustus", "September", "Oktober", "November", "Desember"]

    # Get time range
    prevDay, currDay = getTimeRange()

    excelName = "Report SOC tgl " + str(prevDay.day) + " " + str(listMonths[prevDay.month - 1]) + " - " + str(currDay.day - 1) + " " + str(listMonths[currDay.month - 1]) + ".xlsx"
    
    excelBook = xlsxwriter.Workbook(excelName)

    # Sheet 1 - Full Ticket List #
    excelSheet = excelBook.add_worksheet()

    excelSheet.write('A1', "Created")
    excelSheet.write('B1', "Issue Key")
    excelSheet.write('C1', "Reporter")
    excelSheet.write('D1', "PIC")
    excelSheet.write('E1', "Summary")
    excelSheet.write('F1', "Source IP")
    excelSheet.write('G1', "Destination IP")
    excelSheet.write('H1', "Comment")
    excelSheet.write('I1', "Priority")
    excelSheet.write('J1', "Status")
    excelSheet.write('K1', "Labels")
    excelSheet.write('L1', "Category")

    index = 0

    for data_i in dataIn:
        index += 1
        for data_j in range(len(data_i)):
            excelSheet.write(index, data_j, data_i[data_j])

    # Sheet 2 - Medium and Low Ticket split #
    excelSheet_2 = excelBook.add_worksheet()
    
    lowTickets = []
    mediumTickets = []

    for data_k in dataIn:
        if data_k[8] == "Low":
            lowTickets.append(data_k)
        else:
            mediumTickets.append(data_k)

    excelSheet_2.write('A1', "LOW TICKETS")
    excelSheet_2.write('A2', "Created")
    excelSheet_2.write('B2', "Issue Key")
    excelSheet_2.write('C2', "Reporter")
    excelSheet_2.write('D2', "PIC")
    excelSheet_2.write('E2', "Summary")
    excelSheet_2.write('F2', "Source IP")
    excelSheet_2.write('G2', "Destination IP")
    excelSheet_2.write('H2', "Comment")
    excelSheet_2.write('I2', "Priority")
    excelSheet_2.write('J2', "Status")
    excelSheet_2.write('K2', "Labels")
    excelSheet_2.write('L2', "Category")

    index_2 = 1

    for lows_i in lowTickets:
        index_2 += 1
        for lows_j in range(len(lows_i)):
            excelSheet_2.write(index_2, lows_j, lows_i[lows_j])

    index_2 += 3
    excelSheet_2.write('A'+str(index_2), "MEDIUM TICKETS")

    for meds_i in mediumTickets:
        index_2 += 1
        for meds_j in range(len(meds_i)):
            excelSheet_2.write(index_2, meds_j, meds_i[meds_j])

    # Sheet 3 - Ticket Data
    excelSheet_3 = excelBook.add_worksheet()

    excelSheet_3.write(0,0, "Description")
    excelSheet_3.write(1,0, "All Tickets")
    excelSheet_3.write(2,0, "All Tickets Close")
    excelSheet_3.write(3,0, "All Tickets Open")

    excelSheet_3.write(0,1, "High")
    excelSheet_3.write(0,2, "Medium")
    excelSheet_3.write(0,3, "Low")
    excelSheet_3.write(0,4, "Total")

    excelSheet_3.write("C2", str(len(mediumTickets)))                       # All tickets Medium
    excelSheet_3.write("D2", str(len(lowTickets)))                          # All Tickets Low
    excelSheet_3.write("E2", str(len(dataIn)))                              # All Tickets Total Tickets

    excelSheet_3.write("B2", "0")
    excelSheet_3.write("B3", "0")
    excelSheet_3.write("B4", "0")
    excelSheet_3.write("D4", "0")

    closed_meds_counter_var = 0
    for meds_k in mediumTickets:
        if meds_k[7] is not None:
            closed_meds_counter_var += 1

    excelSheet_3.write("C3", str(closed_meds_counter_var))                         # All Tickets Close Medium
    excelSheet_3.write("C4", str(len(mediumTickets) - closed_meds_counter_var))    # All Tickets Open Medium
    excelSheet_3.write("D3", str(len(lowTickets)))                                 # All Tickets Close Low

    excelSheet_3.write("E3", str(closed_meds_counter_var + len(lowTickets)))       # All Tickets Close Total
    excelSheet_3.write("E4", str(len(mediumTickets) - closed_meds_counter_var))    # All Tickets Open Total

    # Sheet 4 - "Types of attacks - grouped"
    excelSheet_4 = excelBook.add_worksheet()

    summary_list = []
    for data_l in dataIn:
        summary_list.append(data_l[4])

    item_counts = Counter(summary_list)

    index_l = 0;
    for item, count in sorted(item_counts.items(), key=lambda x: x[1], reverse=True):
        #print(f"{item} : {count}")
        excelSheet_4.write(index_l, 0, item)
        excelSheet_4.write(index_l, 1, count)
        index_l += 1;

    # Close book to save
    excelBook.close()

# Test function, do whatever here
def test():
    # Fill me

if __name__ == "__main__":
    
    print("--- Weekly Automation V2.3 ---")
    print("")
    data_acq()

    #test()  


    


