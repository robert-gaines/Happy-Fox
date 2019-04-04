#!/usr/bin/env python3
#
import time
import json
import requests
#
def FindTimeResolved(ticket):
    #
    length = len(ticket)-1 ; i = 0
    #
    while(i < length):
        #
        i += 1
        #
    time_resolved = ticket[i]['timestamp']
    #
    return time_resolved
#
print("[~] Gathering parameters... ")
#
time.sleep(1)
#
#ticket_number = int(input("[+] Enter the total number of tickets-> ")); n = 0
#
api_key = "2601bc33167b42f8a074463460d6e8de"
#
auth_token = "df0addc1b0f54ed18aa51e541e2fd7cf"
#
#
UniformResourceLocator = "https://quantadyn.happyfox.com/api/1.1/json/ticket/1032/"
#
authorization = (api_key,auth_token)
#
response = requests.get(UniformResourceLocator,auth=authorization)
#
response_string = str(response)
#
intake = response.json()
#
temp = []
#
temp = intake['updates']
#
FindTimeResolved(temp)
#
""" custom_fields = intake['custom_fields']
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
device_id = IdentifyDevice(unit) #Value should be -> custom_field_four['value']#
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
remarks =  custom_field_two['value'] """
#
