import json
import csv
from datetime import datetime
from colorama import Style, Fore, Back

JSON_basename = "formatted_output_"
JSON_Count = 1
file_count = 1

while True:
    currJSONFile = f"{JSON_basename}{JSON_Count}.json"

    with open(currJSONFile, 'r') as f:
        lists_of_lists = json.load(f)
        print(f"Converting {JSON_basename}{JSON_Count}.json to CSV File")

    flattened_data = []
    for json_list in lists_of_lists:
        for entry in json_list:
            flat_entry = {}
            source = entry.get('source', {})
            destination = entry.get('destination', {})   # Correct key for destination.ip
            destinatio = entry.get('destinatio', {})    # As per your original data for port

            flat_entry['@timestamp'] = entry.get('@timestamp', '')
            flat_entry['user'] = entry.get('user', '')
            flat_entry['source.ip'] = source.get('ip', '')
            flat_entry['source.port'] = source.get('port', '')
            flat_entry['destination.ip'] = destination.get('ip', '')
            flat_entry['destinatio.port'] = destinatio.get('port', '')
            flat_entry['usingpolicy'] = entry.get('usingpolicy', '')
            flat_entry['utmaction'] = entry.get('utmaction', '')
            flat_entry['utmevent'] = entry.get('utmevent', '')
            flat_entry['url'] = entry.get('url', '')
            flat_entry['srcname'] = entry.get('srcname', '')
            flat_entry['service'] = entry.get('service', '')

            flattened_data.append(flat_entry)

    fieldnames = ['@timestamp', 'user', 'source.ip', 'source.port', 'destination.ip', 'destinatio.port',
                  'usingpolicy', 'utmaction', 'utmevent', 'url', 'srcname', 'service']

    before_Time = ((datetime.fromisoformat(flattened_data[0]["@timestamp"])).ctime())[4:-5]
    after_time = ((datetime.fromisoformat(flattened_data[-1]["@timestamp"])).ctime())[4:-5]

    
    if before_Time[5] == after_time[5]:
        filename = f"{file_count}. Forti Analyzer ~ {before_Time.replace(":",".")} - {after_time[-8:].replace(":",".")} 2025.csv"
    else:
        filename = f"{file_count}. Forti Analyzer ~ {before_Time.replace(":",".")} - {after_time.replace(":",".")} 2025.csv"

    #print(datetime.fromisoformat(flattened_data[0]["@timestamp"]))
    #print(datetime.fromisoformat(flattened_data[-1]["@timestamp"]))   
    #print(filename)


    with open(filename, 'w', newline='') as csvfile:
        writer = csv.DictWriter(csvfile, fieldnames=fieldnames)
        writer.writeheader()
        writer.writerows(flattened_data)
        csvfile.close()
        print(f"File Converted to CSV with filename {Fore.CYAN}{filename}{Style.RESET_ALL}")

    JSON_Count += 1
    file_count += 1

    f.close()
