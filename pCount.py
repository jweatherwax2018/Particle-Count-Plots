import os
import matplotlib.pyplot as plt
#from gsheets import Sheets
from datetime import date
from apiclient import discovery
from google.oauth2 import service_account
import datetime

SHEET_ID = "1MKaU-G_wmMnCOLTe292c1wrlWi8A_MyVMYt-9Fn1UFo"
WORKSHEETS = ["Current Month(s)"]
LEGEND = {".3":["Loc.1", "Loc.2", "Loc.3", "ISO6"], ".5":["Loc.1", "Loc.2", "Loc.3", "ISO6", "ISO7"], "5":["Loc.1", "Loc.2", "Loc.3", "ISO6", "ISO7"]}

#flag for ignoring zeroes in data
IGNORE_ZEROS = False

def get_data(service, range, data_range):

    request = service.spreadsheets().values().get(spreadsheetId=SHEET_ID, range=range)
    response = request.execute()
    

    dates = []
    counts = []
    for col in data_range[1:]:
        counts.append([])
    
    for data in response['values'][2:]:
        # skip invalid data
        if len(data) < 16:
            continue
        
        has_null_values = False
        for col in data_range:
            if len(data[col]) < 1 or (IGNORE_ZEROS and (col!=0 and float(data[col]) <= 0.01)):
                has_null_values = True
                
        if has_null_values:
            continue
        date = None
        try:
            date = datetime.datetime.strptime(data[0],"%m/%d/%Y")
        except:
            pass
        
        if data is None:
            continue        
        
        # Append Data to list
        dates.append(date)
        for i, col in enumerate(data_range[1:]):
            counts[i].append(float(data[col]))
            
    return (dates, counts)
    

    
    
def parse_data(service):
    dates = []
    counts = []
    #Ask user what graph they wanr to show
    inp = input("Particle Size to graph? (options are '.3', '.5', '5'): ")
    inp = inp[:2]
   
    #Take relevant data from google sheet
    print(f"Downloading Data for {WORKSHEETS[0]}...")
    if inp == ".3":
        dates, counts = get_data(service, range = f"{WORKSHEETS[0]}!A:P", data_range=[0,1,2,3,4])
    elif inp == ".5":
        dates, counts = get_data(service, range = f"{WORKSHEETS[0]}!A:P", data_range=[0,6,7,8,9,10])
        pass
    elif inp == "5":
        dates, counts = get_data(service, range = f"{WORKSHEETS[0]}!A:P", data_range=[0,11,12,13,14,15])
        pass
    else:
        print("INVALID INPUT")
        exit(-1)
    print(f"Finished Downloading Data for {WORKSHEETS[0]}.")
    return (dates, counts, inp)

def main():
    SCOPE = ["https://spreadsheets.google.com/feeds",'https://www.googleapis.com/auth/spreadsheets',"https://www.googleapis.com/auth/drive.file","https://www.googleapis.com/auth/drive"]

    print("Connecting...")
    creds = service_account.Credentials.from_service_account_file("C:\\Users\\Joseph\\Documents\\Research\\Particle Count Plot\\client_secret.json",scopes=SCOPE);
    service = discovery.build('sheets','v4', credentials=creds)

    dates, counts, inp = parse_data(service)
    m = 0
    X = []
    Y = []
    for i, _ in enumerate(counts):
        Y.append([])
    
    #If date column is empty, move to next date
    for i, date in enumerate(dates):
        if date is None:
            continue
        
        X.append(date)
        for j,_ in enumerate(counts):
            Y[j].append(counts[j][i])
    
    #Plotting
    for line in Y:
        plt.plot(X, line)
        m = max(m, max(line))
    plt.yscale("log")
    plt.xlim(dates[0] + datetime.timedelta(days=-5),dates[-1] + datetime.timedelta(days=+5))
    
    plt.legend(LEGEND[inp],loc='upper right',bbox_to_anchor=(0,1.05,1.11,0))
    plt.xlabel("Date")
    plt.ylabel("Particle Count")
    plt.title('{}μm'.format(inp), fontdict = {'fontsize' : 20})
    plt.show()

if __name__ == "__main__":
    main()
