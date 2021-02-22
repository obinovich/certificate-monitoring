import pandas as pd
import os
import datetime
import xlrd
import csv
import numpy as np
import json
import sys
import pickle
import requests
import urllib3
from shutil import copyfile
from pathlib import Path
from office365.runtime.auth.authentication_context import AuthenticationContext
from office365.sharepoint.client_context import ClientContext
from office365.sharepoint.files.file import File
from office365.runtime.http.request_options import RequestOptions

main_path=Path("./")
dst_repo=main_path / "certificate repository.xlsx"
url="https://[sharepoint site]/sites/[path to file]/certificate inventory.xlsx" #path should have the spaces

def get_secrets_from_storage():
    s_dict = dict()
    #with open('secrets_ext.obj', 'rb') as f:
    with open(main_path/'secrets.obj', 'rb') as f:
        s_dict = pickle.load(f)
        f.close()
    return s_dict

def get_off_secrets_from_storage():
    s_off_dict = dict()
    #with open('secrets_ext.obj', 'rb') as f:
    with open(main_path/'off-secrets.obj', 'rb') as g:
        s_off_dict = pickle.load(g)
        g.close()
    return s_off_dict

#Delete old sharepoint file
if os.path.exists(dst_repo):
    os.remove(dst_repo)
    print("Old file deleted")
else:
    print("The old file does not exist")



#Authenticate with Sharepoint
try:
    ctx_auth = AuthenticationContext(url)
    username = get_off_secrets_from_storage()['uname']
    password = get_off_secrets_from_storage()['pwd']
    ctx_auth.acquire_token_for_user(username, password)
    ctx = ClientContext(url, ctx_auth)
    print("Sharepoint auth successful")
except ValueError as e:
    error = 'Sharepoint auth failed : {}'.format(e)
    print(error)
    send_slack_alert(error)

# Authenticate request
options = RequestOptions(url)
ctx_auth.authenticate_request(options)
urllib3.disable_warnings()
req = requests.get(url, headers=options.headers, verify=False, allow_redirects=True)
print("File request successful")

#Download Repo file
output = open(dst_repo, 'wb')
output.write(req.content)
output.close()
print("New File copy complete")

def send_slack_alert(description):
    try:
        # secret
        key = get_secrets_from_storage()['sl_api_key']
        body_dict = dict()
        body_dict.update({'type': 'mrkdwn'})
        body_dict.update({'text': description})
        body_json = json.dumps(body_dict)

        # send to slack channel created for the post_url
        post_url = 'https://hooks.slack.com/services/XXXXXXXXX/YYYYYYYYY/{}'.format(key)  # cert_mon

        r = requests.post(post_url, data=body_json)
        if r.status_code != 200:
            print('slack not send : {} {}'.format(str(r.status_code), r.reason))
    except HTTPError as e:
        error = 'error sending Slack : {}'.format(e)
        print(error)


df = pd.read_excel (dst_repo)
#print (df)

# Output with dates converted to YYYY-MM-DD
df["Date"] = pd.to_datetime(df["Expiry date"], errors='coerce')
dfs = pd.DataFrame(df)
x = datetime.datetime.now()
today = x.strftime("%Y-%m-%d")
#print(today)

#Print message
notify_msg ="======Certificate Expiring Soon======\n"
notify_e_msg ="\n\n======Expired Certificates======\n"

for index, row in dfs.iterrows():
    row_date = pd.to_datetime(row['Date'])
    today = pd.to_datetime(today)
    diff_days = row_date - today
    row_cert = row['Name']
    row_env = str(row['Environment'])
    row_sol = str(row['Solution'])
    row_side = str(row['Client-side or server-side?'])
    row_cust = str(row['Customer-specific?'])
    row_report = str(row['Reported by'])

    days = diff_days.days

    # check for certificate expiring soon
    if (days == 45 or days == 30 or days <= 15) and (days > 0):
        if (row_cert == ""):
            notify_msg = "No Certificate"
        else:
            notify_msg = notify_msg + "Expires in " + str(days) + " days ==> "+ row_cert +".\n"\
            "\t\t--info--\n"\
            "\t\tEnvironment = "+row_env+"\n"\
            "\t\tSolution = "+row_sol+"\n" \
            "\t\tRole = "+row_side+"\n" \
            "\t\tCustomer-Specific = "+row_cust+"\n" \
            "\t\tReport by = "+row_report+"\n" \
            "\t\t--------\n"

    #check for expired certificate
    if (days < 1) and (days >= -30):
        if (row_cert==""):
            notify_e_msg="No Certificate"
        else:
            notify_e_msg = notify_e_msg + row_cert + " expired " + str(abs(days)) + " days ago.\n"\
            "\t\t--info--\n"\
            "\t\tEnvironment = "+row_env+"\n"\
            "\t\tSolution = "+row_sol+"\n" \
            "\t\tRole = "+row_side+"\n" \
            "\t\tCustomer-Specific = "+row_cust+"\n" \
            "\t\tReport by = "+row_report+"\n" \
            "\t\t--------\n"


print(notify_msg)
print(notify_e_msg)

send_slack_alert(notify_msg)
send_slack_alert(notify_e_msg)

#delete file
os.remove(dst_repo)
# -------------------------------
