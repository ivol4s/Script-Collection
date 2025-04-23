'''
Raya Rabu Automation V2
- Generates an excel file with the contents of all tickets created within a week of the day before this script is runned. (ex, today is 25th, script will get tickets from 17th to 24th)

Jira API credential and token (Cuman sampe 31 Desember 2025)
- email
- API TOKEN

Dependencies
- requests
- xlsxwriter
- API Token 1 account. bisa dapet / generate dari link: https://id.atlassian.com/manage-profile/security/api-tokens
    - Note: API token langsung di save disini habis copy, website akan nunjuk token hanya 1 kali

Notes:
- Raya Rabu Automation V2 BABY
    - 30X faster and no need to configure dashboard
    - only requires an API token and not a hint of selenium #bless
- Pastikan di jira sudah complete datanya, ini script tidak cek untuk missing data. 

TODO:
- Set checks for missing data
- cleanup code

'''
import requests
from requests.auth import HTTPBasicAuth
import json
from datetime import timedelta, date, datetime
import xlsxwriter
from collections import Counter
from colorama import init, Fore, Style

def data_acq():
    # Jira instance details
    jira_url = "https://your.site.here"

    # Authentication - using an API token (create one from your Jira account settings)
    email = "coolguy@someDomain.com"  # Replace with your email
    api_token = "abcdefghijkLMAOasIfop.."  # Replace with your API token

    # URL for Jira search endpoint with JQL query
    url = f"{jira_url}/rest/api/3/search"

    #print(url)

    # Date range for filtering issues (use format YYYY-MM-DD)
    start_date, end_date = getTimeRange()
    #print(start_date)
    #print(end_date)

    # Project key SOC - Jira Raya ada banyak project hati-hati pas mau ambil data. 
    project_key = "le_key" # Replace with your key 

    # Set JQL query buat cari ticket yang di created pada time range tertentu
    jql_query = f"project = {project_key} AND created >= '{start_date}' AND created <= '{end_date}'"

    # Set up the headers for the request
    headers = {
        "Accept": "application/json",
        "Content-Type": "application/json",
    }

    # Define all required fields
    fields = [
        "created",           # Timestamp
        "key",               # SOC Key
        "reporter",          # Reporter (note perlu di ulik) 
        "customfield_10681", # PIC 
        "summary",           # Summary
        "customfield_10592", # Source IP
        "customfield_10593", # Destionation IP
        "customfield_10906", # Details
        "priority",          # Priority (note ulik)
        "status",            # Status (ulik di name)
        "labels",            # Labels (bentuk array ini)          
        "customfield_10892", # Category
    ]

        # Define the parameters for the search
    params = {
        "jql": jql_query,          # Use the JQL query to filter by created date
        #"fields": ",".join(fields),
        "maxResults": 1000,         # Maximum number of issues to return (adjust as needed)
    }

    # Make the GET request to Jira API
    response = requests.get(url, headers=headers, auth=HTTPBasicAuth(email, api_token), params=params)

    # Check if the request was successful
    if response.status_code == 200:
        issues_data = response.json()  # Parse the JSON response
        issues = issues_data.get("issues", [])
        
        #print(issues_data)

        #exit()
        if issues:
            print(f"Found {len(issues)} issues between {start_date} and {end_date - timedelta(1)}")
            print(Fore.GREEN + "To get a better experience consider subscribing for the PREMIUM DELUXE EXTRA MAX PRO tier pack to get awesome benefits!" + Style.RESET_ALL)

            full_data_array = []

            for issue in issues:
                curr_data_array = []

                #print(issue)
                #print("\n")

                # Time Created 
                #print(f"TimeStamp: {issue['fields']['created']}")
                timeCreated = datetime.strptime(str(issue['fields']['created']), "%Y-%m-%dT%H:%M:%S.%f+0700")
                curr_data_array.append(timeCreated.strftime("%d/%m/%Y %H:%M"))

                # Issue Key
                #print(issue["key"])  # Print each issue key
                curr_data_array.append(issue["key"])
                
                # Reporter
                reporter_itsec = issue['fields']['reporter']
                #print(reporter_itsec['displayName'])
                curr_data_array.append(reporter_itsec['displayName'])
                
                # PIC Bank Raya
                PIC_client = issue['fields']['customfield_10681']
                if PIC_client == None:
                    curr_data_array.append("INPUT PIC")
                else:
                    #print(PIC_client[0]["displayName"])
                    curr_data_array.append(PIC_client[0]['displayName'])

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
                if destination_ip == None:
                    curr_data_array.append("")
                else:
                    curr_data_array.append(str(destination_ip))

                # Details
                details_manual = issue['fields']['customfield_10906']
                
                #need to check if the customfield_10906 field is populated or not, then check if the inside is empty or not
                # update, code works now as i have found that customfield_10906 (details) have 3 different states 
                #   1) filled - field with will be populated by multiple subfields and will contain a text / point / whatever
                #   2) unfilled - field will not contain anything, simply will present as "customfield_10906: None"
                #   3) filled but deleted - if customfield_10906 has been edited before, it will now generate subfields with empty content
                # Code below checks if customfield_10906 is empty, then checks if customfield_10906["content"] is empty or not (by counting how many variables are in the array empty will be 0 and anythign else is 1)
                # if both is not empty then get the text/
                # NOTE: only works if customfield_10906 is filled with pure text, if it has a point (formatting) this will break, i can put checks so that it will always work but... i'll leave it to you readers as homework :D
                if details_manual != None:
                    if len(details_manual["content"]) != 0:
                        #print(details_manual["content"][0]["content"][0]["text"])
                        curr_data_array.append(details_manual["content"][0]["content"][0]["text"])
                    else:
                        curr_data_array.append("")
                        #print("Tis but an empty place sire!")
                else:
                    curr_data_array.append("")
                    #print("DW my brain is emptier")
          
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
                    curr_data_array.append("")

                # Category
                status = issue['fields']['customfield_10892']['value']
                #print(status)
                curr_data_array.append(status)

                
                #NOTE make checks on all acqquiry just in case

                #print('--------------------')

                full_data_array.append(curr_data_array)
        else:
            print(f"No issues found between {start_date} and {end_date}.")
    else:
        print(f"Failed to retrieve issues. Status code: {response.status_code}")
        print(response.text)  # Print error message if any

    #print(len(full_data_array))

    excelOutput(list(reversed(full_data_array)))

# Helper function to get date ranges
def getTimeRange():
    prevRabu = date.today() - timedelta(8)
    currRabu = date.today() + timedelta(1) # Script bakal jalan pada hari Kamis, jadinya hari dikurang 1
    
    return prevRabu, currRabu

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
    excelSheet.write('D1', "PIC Client")
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
    excelSheet_2.write('D2', "PIC Client")
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

    # Sheet 3 - Jumlah Tiket
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
        if meds_k[7]:
            closed_meds_counter_var += 1

    excelSheet_3.write("C3", str(closed_meds_counter_var))                         # All Tickets Close Medium
    excelSheet_3.write("C4", str(len(mediumTickets) - closed_meds_counter_var))    # All Tickets Open Medium
    excelSheet_3.write("D3", str(len(lowTickets)))                                 # All Tickets Close Low

    excelSheet_3.write("E3", str(closed_meds_counter_var + len(lowTickets)))       # All Tickets Close Total
    excelSheet_3.write("E4", str(len(mediumTickets) - closed_meds_counter_var))    # All Tickets Open Total

    # Sheet 4 - "Tipe serangan yang muncul"
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

def test():
    full_array = []
    array_1 = [1,2,3,4,5,6]
    array_2 = [7,8,9,0]

    full_array.append(array_1)
    full_array.append(array_2)

    for a in full_array:
        print(a)

if __name__ == "__main__":
    
    print("--- Weekly Automation V2 ---")
    print("")
    data_acq()



    




    # Banner
    #subprocess.call("python banner_HO.py")
    
    #print("Generating Output...")
    #print("Please Wait")
    
    #data_acq()
    #test()  

    #test_extract()
    


