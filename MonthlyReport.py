#!/usr/bin/env python3

__auth__ = 'RWG'

import xlsxwriter
import datetime
import requests
import json
import time
import sys

def WriteToSpreadsheet(tickets):
        #
        month = datetime.datetime.now().month
        #
        year = datetime.datetime.now().year
        #
        report_string = "JTC-TRS Field Service Report"+str(month)+"-"+str(year)
        #
        monthly_report = xlsxwriter.Workbook('MonthlyReport.xlsx')
        #
        monthly_sheet = monthly_report.add_worksheet()
        #
        bold = monthly_report.add_format({'bold': True})
        #
        monthly_sheet.insert_image('A1','logo.PNG')
        #
        monthly_sheet.write('B1',report_string,bold)
        #
        monthly_sheet.set_column('A:A',5)
        monthly_sheet.set_column('B:B',10)
        #
        monthly_sheet.write('A2','Ticket ID',bold)
        monthly_sheet.write('B2','Status',bold)
        monthly_sheet.write('C2','Discrepancy',bold)
        monthly_sheet.write('D2','Submitter',bold)
        monthly_sheet.write('E2','Unit',bold)
        monthly_sheet.write('F2','Device ID',bold)
        monthly_sheet.write('G2','Date/Time Reported',bold)
        monthly_sheet.write('H2','Date/Time Acknowledged',bold)
        monthly_sheet.write('I2','Notification Mode',bold)
        monthly_sheet.write('J2','Subject',bold)
        monthly_sheet.write('K2','Repeat CAT 1',bold)
        monthly_sheet.write('L2','Corrective Action',bold)
        monthly_sheet.write('M2','Asignee',bold)
        monthly_sheet.write('N2','Date Rsolved',bold)
        monthly_sheet.write('O2','Root Cause',bold)
        monthly_sheet.write('P2','Remarks',bold)
        #
        row_index = 3 ; col_index = 1
        #
        indeces = ['A','B','C','D','E','F','G','H','I','J','K','L','M','N','O','P']
        #
        for t in tickets:
                #
                for i in range(0,len(indeces)):
                        #
                        ptr = str(indeces[i])+str(row_index)
                        #
                        value = str(t[i])
                        #
                        monthly_sheet.write(ptr,value)
                        #
                        col_index += 1
                        #
                row_index += 1
                        #
                col_index = 1
                #
        monthly_report.close()

def IdentifyDevice(unit_value):
        #
        sites = {}
        #
        sites.update({"15 ASOS":1})
        sites.update({"17 STS":1})
        sites.update({"Quantadyn":2})
        sites.update({"14 ASOS":3})
        sites.update({"7 ASOS":4})
        sites.update({"20 ASOS":5})
        sites.update({"10 ASOS":6})
        sites.update({"19 ASOS":7})
        sites.update({"3 ASOS, Det 1":8})
        sites.update({"25 ASOS":9})
        sites.update({"15 ASOS":10})
        sites.update({"WPC":11})
        sites.update({"2 ASOS":12})
        sites.update({"12 CTS":13})
        sites.update({"20 ASOS, Det 1":14})
        sites.update({"6 CTS":15})
        sites.update({"6 CTS":16})
        sites.update({"6 CTS":17})
        sites.update({"5 ASOS":18})
        sites.update({"23 STS":19})
        sites.update({"STTS":20})
        sites.update({"13 ASOS":21})
        sites.update({"9 ASOS":22})
        sites.update({"705th DMOC":23})
        sites.update({"11 ASOS":24})
        sites.update({"22 STS":25})
        sites.update({"607 ASOG":26})
        sites.update({"24 STS":27})
        sites.update({"26 STS":28})
        sites.update({"321 STS":29})
        sites.update({"3 ASOS":30})
        #
        if(unit_value in sites):
                #
                return int(sites.get(unit_value))
                #
        else:
                #
                return 0

def SortMenu():
        #
        print('''
        *****************
        *** Month Menu ***
        *****************
        1) January
        2) February
        3) March
        4) April
        5) May
        6) June
        7) July
        8) August
        9) September
       10) October
       11) November
       12) December
              ''')
        #
        month_selection = int(input("[+] Enter the month integer value-> "))
        #
        if(month_selection > 12 or month_selection < 1):
                #
                sys.exit("[!] Invalid Selection ")
                #
        else:
                #
                return month_selection

def BubbleSort(tickets):
        #
        length = len(tickets)
        #
        for i in range(0,length):
                #
                for j in range(0,length-i-1):
                        #
                        if(tickets[j][5] > tickets[j+1][5]):
                                #
                                temp = tickets[j]
                                #
                                tickets[j] = tickets[j+1]
                                #
                                tickets[j+1] = temp
                                #
        return tickets

def SortTickets(month,tickets):
        #
        sorted_list = [] ; temp_list = []
        #
        time_value = datetime.datetime.now().year
        #
        current_year = str(time_value)
        #
        for t in tickets:
                #
                statuses = ["Solved","Open","Pending","Closed","New","Hold"]
                #
                ongoing = ["Open","New","Pending","Hold"]
                #
                date = t[6]
                #
                month_index = int(date[5:7])
                #
                if((month == month_index) or (t[1] in ongoing) or ((statuses[3] in t[1]) and month_index<month)):
                        #
                        temp_list.append(t)
                        #
        sorted_list = BubbleSort(temp_list)
        #
        return sorted_list

def GatherTickets():
    #
    tickets = []
    #
    print("[~] Gathering parameters... ")
    #
    time.sleep(1)
    #
    ticket_number = int(input("[+] Enter the total number of tickets-> ")); n = 0
    #
    api_key = "2601bc33167b42f8a074463460d6e8de"
    #
    auth_token = "df0addc1b0f54ed18aa51e541e2fd7cf"
    #
    while(n <= ticket_number):
        #
        UniformResourceLocator = "https://quantadyn.happyfox.com/api/1.1/json/ticket/"+str(n)+"/"
        #
        authorization = (api_key,auth_token)
        #
        response = requests.get(UniformResourceLocator,auth=authorization)
        #
        response_string = str(response)
        #
        if('2' in response_string):
                #
                intake = response.json()
                #
                try:
                        #
                        custom_fields = intake['custom_fields']
                        custom_field_zero = custom_fields[0]
                        custom_field_one = custom_fields[1]
                        custom_field_two = custom_fields[2]
                        custom_field_three = custom_fields[3]
                        custom_field_four = custom_fields[4]
                        custom_field_five = custom_fields[5]
                        custom_field_six = custom_fields[6]
                        category = intake['category']
                        status = intake['status']
                        updates = intake['updates']
                        #
                        ticket_id = intake['display_id'] 
                        status_value = status['name']
                        discrepancy = intake['priority']['name']
                        submitter = intake['assigned_to']['name']
                        unit = intake['user']['custom_fields'][0]['value']
                        device_id = IdentifyDevice(unit) #custom_field_four['value']#
                        date_time_reported = intake['created_at']
                        date_time_acknowledged = intake['created_at']
                        notification = custom_field_zero['value']
                        subject = intake['subject']
                        repeat_cat_one = custom_field_five['value']
                        #
                        if(repeat_cat_one == 1):
                                #
                                repeat_cat_one = 'Yes'
                                #
                        else:
                                repeat_cat_one = "No"
                        #
                        corrective_action = custom_field_three['value']
                        asignee = intake['assigned_to']['name']
                        date_resolved = intake['last_updated_at'][0:10]
                        #
                        try:
                                #
                                root_cause = custom_field_two['name']+" "+custom_field_two['value']
                                #
                        except:
                                #
                                root_cause = "Undefined"
                                #
                        remarks =  custom_field_two['value']
                        #
                        temp_list = []
                        #
                        temp_list.append(ticket_id)
                        temp_list.append(status_value)
                        temp_list.append(discrepancy)
                        temp_list.append(submitter) 
                        temp_list.append(unit) 
                        temp_list.append(device_id) 
                        temp_list.append(date_time_reported) 
                        temp_list.append(date_time_acknowledged) 
                        temp_list.append(notification) 
                        temp_list.append(subject) 
                        temp_list.append(repeat_cat_one)
                        temp_list.append(corrective_action)
                        temp_list.append(asignee) 
                        temp_list.append(date_resolved) 
                        temp_list.append(root_cause) 
                        temp_list.append(remarks)
                        #
                        tickets.append(temp_list)
                        #
                except Exception as e:
                        #
                        print("[!] Error: ", e)
                        #
        else:
                #
                pass
                #
        n += 1
        #
    return tickets

def main():
        #
        print('''
                 /$$      /$$                       /$$     /$$       /$$                                
                | $$$    /$$$                      | $$    | $$      | $$                                
                | $$$$  /$$$$  /$$$$$$  /$$$$$$$  /$$$$$$  | $$$$$$$ | $$ /$$   /$$                      
                | $$ $$/$$ $$ /$$__  $$| $$__  $$|_  $$_/  | $$__  $$| $$| $$  | $$                      
                | $$  $$$| $$| $$  \ $$| $$  \ $$  | $$    | $$  \ $$| $$| $$  | $$                      
                | $$\  $ | $$| $$  | $$| $$  | $$  | $$ /$$| $$  | $$| $$| $$  | $$                      
                | $$ \/  | $$|  $$$$$$/| $$  | $$  |  $$$$/| $$  | $$| $$|  $$$$$$$                      
                |__/     |__/ \______/ |__/  |__/   \___/  |__/  |__/|__/ \____  $$                      
                                                                        /$$  | $$                      
                                                                        |  $$$$$$/                      
                                                                         \______/                       
                 /$$$$$$$                                            /$$                                 
                | $$__  $$                                          | $$                                 
                | $$  \ $$  /$$$$$$   /$$$$$$   /$$$$$$   /$$$$$$  /$$$$$$                               
                | $$$$$$$/ /$$__  $$ /$$__  $$ /$$__  $$ /$$__  $$|_  $$_/                               
                | $$__  $$| $$$$$$$$| $$  \ $$| $$  \ $$| $$  \__/  | $$                                 
                | $$  \ $$| $$_____/| $$  | $$| $$  | $$| $$        | $$ /$$                             
                | $$  | $$|  $$$$$$$| $$$$$$$/|  $$$$$$/| $$        |  $$$$/                             
                |__/  |__/ \_______/| $$____/  \______/ |__/         \____/                               
                                    | $$                                                                 
                                    | $$                                                                 
                                    |__/                                                                 
                 /$$$$$$                                                     /$$                        
                /$$__  $$                                                   | $$                        
                | $$  \__/  /$$$$$$  /$$$$$$$   /$$$$$$   /$$$$$$  /$$$$$$  /$$$$$$    /$$$$$$   /$$$$$$ 
                | $$ /$$$$ /$$__  $$| $$__  $$ /$$__  $$ /$$__  $$|____  $$|_  $$_/   /$$__  $$ /$$__  $$
                | $$|_  $$| $$$$$$$$| $$  \ $$| $$$$$$$$| $$  \__/ /$$$$$$$  | $$    | $$  \ $$| $$  \__/
                | $$  \ $$| $$_____/| $$  | $$| $$_____/| $$      /$$__  $$  | $$ /$$| $$  | $$| $$      
                |  $$$$$$/|  $$$$$$$| $$  | $$|  $$$$$$$| $$     |  $$$$$$$  |  $$$$/|  $$$$$$/| $$      
                \_______/ \________/|__/  |__/ \_______/|__/      \_______/  \_____/  \______/ |__/      
                                                                                                        
                                                                                                        
                                                                                                        
                                                                                                        
                                                                                                        
                                                                                                        
                /$$$$$$ /$$$$$$ /$$$$$$ /$$$$$$ /$$$$$$ /$$$$$$ /$$$$$$                                 
                |______/|______/|______/|______/|______/|______/|______/                                 
                                                            
              ''')
        #
        time.sleep(3)
        #
        tickets = []
        #
        sorted_tickets = []
        #
        tickets = GatherTickets()
        #
        month_index = SortMenu()
        #
        sorted_tickets = SortTickets(month_index,tickets)
        #
        WriteToSpreadsheet(sorted_tickets)
        #
        print()
        #
        print("[*] Finished, check directory for excel product")
        #
        time.sleep(3)

if(__name__ == '__main__'):
        #
        main()

