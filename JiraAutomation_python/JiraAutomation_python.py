''' 
i am uploading this script to githuh for posterity and so i remember not to try this ever again.
---------------------- Weekly Automation Script (DEPRECATED) ---------------------------------
Script Intro
- This script provides weekly automation to acquire Jira tickets by means of web scraping. 
- Script can be runned from windows cmd.
- This script is now deprecated and cannot be used in Jira with the current (infinite scrolling tickets) dashboard view. You can thank the Atlassian frontend team for that completely sane decision.

Dependencies
    - PywhatKit
    - Selenium + BeautifulSoup
    - xlsxwriter
    - pip 

Dependencies Install command 
 - pip install Pywhatkit, selenium, xlswriter

TODO:
 - Change all static time.sleep() methods to Await command for faster and more reliable report generation 
 - Output Excel file to a chosen directory 
 - Add Cell colouring to empty cells
 - Cleanup funtcions

'''

from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from bs4 import BeautifulSoup
from colorama import Style, Fore, Back
import time
import xlsxwriter
from datetime import timedelta, date, datetime


def access_and_get_data():
    i = 0

    # Configuration for Firefox webdriver
    # Setting ignore certificate to bypass "Your connection is not private"
    options = webdriver.FirefoxOptions()
    options.add_argument('--headless') # make browser run in background
    options.add_argument('--ignore-ssl-errors=yes')
    options.add_argument('--ignore-certificate-errors')

    # Use Firefox webdriver
    driver = webdriver.Firefox(options=options) #webdriver placed in C:\Users\%username%\AppData\Local\Programs\Python\Python312\Scripts
    
    # Needed for webpage to be completely rendered
    wait = WebDriverWait(driver, 30)

    # Establish elastic logon session, click first option, and check input name=username & password is loaded
    driver.get("https://id.atlassian.com/login?continue=https%3A%2F%2Fhome.atlassian.com%2F&application=atlas")
  
    #input username and password
    username_field = wait.until(EC.presence_of_element_located(("name", 'username')))
    username_field.send_keys('<REPLACE ME>')                                 # Username
    username_field.send_keys(Keys.RETURN)

    time.sleep(3)

    password_field = wait.until(EC.presence_of_element_located(("name", 'password')))
    password_field.send_keys('<REPLACE ME>')                                         # Password
    password_field.send_keys(Keys.RETURN)

    time.sleep(5) #to ensure session established, add a few second before changed to another tab
    

    # Get date range
    currWeek, prevWeek = getRabuTime()

    #filter_rabu = 'https://redacted.atlassian.net/jira/software/projects/SOC/issues/?jql=project = "SOC" AND "timestamp[time stamp]" >= "' + str(prevWeek) + ' 00:00" AND "timestamp[time stamp]" <= "' + str(currWeek) + ' 23:59" ORDER BY created DESC'

    filter_rabu = '<REPLACE ME WITH URL>' #sample above

    # Navigate to the dashboard URL
    driver.get(filter_rabu)

    # Wait for some time to ensure the page fully loads (adjust the sleep duration a sneeded)
    time.sleep(30)

    # Begin Main Sequence
    data_list = []

    # Below is old cold for when jira still used pagination to show tickets.
    # Check if there are multiple pages
    nextPageVariable = checkForNextPage(driver)

    # If only one page then get data and go to excel output.
    if not nextPageVariable:
        #False (No pages)
        scraped_data = scrapeCurrentPage(driver)

        # Get Data from current page
        parsedData_noPages = getFormattedData(scraped_data) 

        # Output to excel
        excelOutput(parsedData_noPages)
    else:
        #True (Multiple Pages)
        nextPageVariable = True
        while nextPageVariable is not False:
            scraped_data_CurrPage = scrapeCurrentPage(driver)

            # Get Data from current page
            parsedData_currPage = getFormattedData(scraped_data_CurrPage)

            # Add to global list
            data_list = data_list + parsedData_currPage

            # Determine if there is a next page and get button
            nextPageVariable = checkForNextPage(driver)

            print(nextPageVariable)

            # Click to next page if exist
            if nextPageVariable is False:
                time.sleep(3)
                break
            else:
                nextPageVariable.click()

            # Let next page load
            time.sleep(3)
        else:
            # No more pages, currently on last page.
            scraped_data_CurrPage = scrapeCurrentPage(driver)
            
            # Get Data from current page
            parsedData_currPage = getFormattedData(scraped_data_CurrPage)

            # Add to global list
            data_list = data_list + parsedData_currPage

        # Output full list to excel
        excelOutput(data_list)

def scrapeCurrentPage(driver):
    # Scrape data from dashboard
    try:
        # Find Main Table
        #main_table_container = driver.find_element(By.XPATH, '//div[@data-vc="issue-table-main-container"]')
        main_table_container = driver.find_element(By.XPATH, '<REPLACE ME>') # get XPATH from browser, sample above

        return main_table_container
    except Exception as e:
        print(f"An error occurred: {str(e)}")
        exit()

def checkForNextPage(driver):
    # Helper Function to determine if there are pages or not
    try:  
        next_page_button = driver.find_element(By.XPATH, '//button[@data-testid="native-issue-table.ui.footer-renderer.pagination-controls.pagination-controls--right-navigator"]')
        soupedButton = BeautifulSoup(next_page_button.get_attribute('outerHTML'), 'html.parser')
    except Exception as e:
        print(f"An error has occured: {str(e)}")
        exit()

    # Method: Find the next page button and check if there is a "disabled" attribute.
    button_disabled_checker = soupedButton.find('button', disabled=True)

    # If "next" button can be clicked then return the button element, else False boolean
    if button_disabled_checker == None:
        return next_page_button
    else:
        return False


def getFormattedData(raw_Data):
    inner_html = raw_Data.get_attribute('innerHTML')

    # Parse HTML with beutiful soup
    soupedHTML = BeautifulSoup(inner_html, 'html.parser')

    # Get all rows in the table
    all_rows = soupedHTML.find_all('tr')
    all_rows.pop(0)     # First element is empty due to the program taking the table titles

    #print(len(all_rows))

    full_data = []

    for rowData in all_rows:
        full_data.append(getCurrRowData(rowData))

    return(full_data)

def getCurrRowData(parsedHTML_soup):
    # Find all columns within this row
    all_columns_in_row = parsedHTML_soup.find_all('td')

    #Note: 
    # "hey why is there so much code just to get datetime here ? seems overkill no ?"
    # Let me introduce you to the second mind boggling Atlassian decision that I encountered while making this - removing date time from the columns to be replaced with... relative time!
    # Now Jira shows all tickets within a 7 day period as - 1 day ago, 2 days ago, and so on. 
    # So this code is to extract the datetime WHICH IS STILL IN THE HTML. GOD i swear the atlassian frontend team just want to make everything slightly more annoying to use. Thankyou backend team for the API.
    # If any frontend atlassian team reads this (apologies) go pound sand and lookup the term "innovation" in a dictionary.
    try:
        title_span = parsedHTML_soup.find('time', datetime=True)

        dt = datetime.strptime(str(title_span['datetime']), "%Y-%m-%dT%H:%M:%S%z")
        formatted_timestamp = dt.strftime("%d/%m/%Y %I:%M %p")

        all_columns_in_row[0] = formatted_timestamp
    except:
        all_columns_in_row[0] = "n/a"

    try:
        all_columns_in_row[2] = getName(all_columns_in_row[2].text)
    except:
        all_columns_in_row[2] = "n/a"

    # Get all data on this row. Last 2 elements (from jira columns) are uneeded so remove from generated list
    currRowData = []
    for i in range(len(all_columns_in_row) - 2):
        if i == 0 or i == 2:
            currRowData.append(all_columns_in_row[i])
        else:
            text = all_columns_in_row[i].text.strip()
            currRowData.append(text if text else "n/a")

    return(currRowData)

# Helper function to format string that is given weirdly.
def getName(reporterName):
    actualName = ""

    # If name is empty return warning string
    if len(reporterName) == 0:
        return("NAME EMPTY CHECK AGAIN")

    # Check if current name has spaces in it. Modify as required.
    # Note: 
    # Scraping from Jira results in duplicate names (only for reporter column)
    #   - Example = "Ivan Christian Halim" Becomes "Ivan Christian HalimIvan Christian Halim" 
    # Code below is to fix the issue
    #
    # This.. is incredibly stupid but during development I was unable (read: did not bother) to find another way. apologies for the eye strain.
    if " " in reporterName:                             # Check if name contains multiple words
        splittedName = reporterName.split()             # Split names into single words and put into array              
        splittedName.pop(len(splittedName)//2)          # remove duplicate name (always in the middle element of the array)
        finalName = list(dict.fromkeys(splittedName))   # Remove duplicate strings

        seperator = ' '                         
        stringifiedName = seperator.join(finalName)     # turn array into string
        actualName = stringifiedName
    else:
        actualName = reporterName[:len(reporterName)//2]

    return(actualName)

# Helper function to generate excel output
def excelOutput(dataIn):
    # Init Var
    listMonths = ["Januari", "Februari", "Maret", "April", "Mei", "Juni", "Juli", "Augustus", "September", "Oktober", "November", "Desember"]

    # Get time range
    currDay, prevDay = getRabuTime()

    # Check if currRabu and prevRabu are wednesday
    if currDay and prevDay != 2:
        print("WARNING: Tanggal dan Hari bukan rabu ke rabu")

    excelName = "Report SOC tgl " + str(currDay.day) + " " + str(listMonths[currDay.month - 1]) + " - " + str(prevDay.day) + " " + str(listMonths[prevDay.month - 1]) + ".xlsx"
    
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
        if meds_k[7] is not None:
            closed_meds_counter_var += 1

    excelSheet_3.write("C3", str(closed_meds_counter_var))                         # All Tickets Close Medium
    excelSheet_3.write("C4", str(len(mediumTickets) - closed_meds_counter_var))    # All Tickets Open Medium
    excelSheet_3.write("D3", str(len(lowTickets)))                                 # All Tickets Close Low

    excelSheet_3.write("E3", str(closed_meds_counter_var + len(lowTickets)))       # All Tickets Close Total
    excelSheet_3.write("E4", str(len(mediumTickets) - closed_meds_counter_var))    # All Tickets Open Total

    # Sheet 4 - "Tipe serangan yang muncul"
        #TODO 

    # Close book to save
    excelBook.close()

# Helper function to get date ranges
def getRabuTime():
    currRabu = date.today() - timedelta(1)
    prevRabu = date.today() - timedelta(8)
    
    return currRabu, prevRabu

# Test function, try stuff out here
def test_extract():
    excelBook = xlsxwriter.Workbook("test.xlsx")

    # Sheet 1 - Full Ticket List
    excelSheet = excelBook.add_worksheet()
    excelSheet.write('A1', "Created")
    

    excelSheet_2 = excelBook.add_worksheet()
    excelSheet_2.write(1,1, "askd")

    excelBook.close()


if __name__ == "__main__":
  
    # Banner
    print("Generating Output...")
    print("Please Wait")
    
    access_and_get_data()

    #test_extract()
    
