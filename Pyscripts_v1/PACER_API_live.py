
import requests
import json
import ast
import urllib3
import pyodbc
import pandas as pd
import base64
import datetime
import keyring
#import modin.pandas as pd
LMS_db_config = 'DRIVER={SQL Server};SERVER=AC-PC-097;DATABASE=FinanceTeamDB;Trusted_Connection=yes;'
cnxn = pyodbc.connect(LMS_db_config)
##test response
http = urllib3.PoolManager()
url_test = http.urlopen('POST', auth_post_url,
                 headers=headers_dict,
                 body=encoded_body)
url_test.read()

req = requests.Request('POST',auth_post_url,headers=headers_dict,data=encoded_body)
prepared = req.prepare()

def pretty_print_POST(req):
    """
    At this point it is completely built and ready
    to be fired; it is "prepared".

    However pay attention at the formatting used in 
    this function because it is programmed to be pretty 
    printed and may differ from the actual request.
    """
    print('{}\n{}\r\n{}\r\n\r\n{}'.format(
        '-----------START-----------',
        req.method + ' ' + req.url,
        '\r\n'.join('{}: {}'.format(k, v) for k, v in req.headers.items()),
        req.body,
    ))

pretty_print_POST(prepared)
r2 = requests.Request('POST',pacer_auth_post_url,headers = headers_dict, data = encoded_body)
prepared2 = r2.prepare()
pretty_print_POST(prepared2)

def get_auth_token():
    auth_url = 'qa-login.uscourts.gov'
    pacer_url = 'pacer.login.uscourts.gov'
    username = 'americashloans'
    password = keyring.get_password('pacer',username)
    testclientcode = 'testclientcode'
    auth_post_url = f"https://{auth_url}/services/cso-auth"
    pacer_auth_post_url = f"https://{pacer_url}/services/cso-auth"

    login_dict = {'loginId': username,'password':password, "redactFlag":1,}

    headers_dict = {'Content-type': 'application/json', 'Accept': 'application/json'}
    encoded_body = json.dumps(login_dict,ensure_ascii=False)

    #Using requests problems with "!" character
    r = requests.post(pacer_auth_post_url,headers = headers_dict, data = encoded_body.encode('utf-8') 
                      )
    auth_dict = r.json()
    auth_token = auth_dict['nextGenCSO']
    return auth_token

#token active for 30 min without use then auth needs to be repulled
auth_token = get_auth_token()
pcl_url = 'qa-pcl.uscourts.gov'

pcl_headers = {'Content-type': 'application/json', 'Accept': 'application/json', "X-NEXT-GEN-CSO" : auth_token}
pcl_case_search = { "caseTitle": "Jacob" }

{
"lastName": "Henderson",
"firstName": "Nicholas",
"exactNameMatch": False,
"courtCase": {
"jurisdictionType": "bk",
"caseType": [
"cv", "ncrim", "misc"
],

"courtId": [
"ilcbk", "ilcdc"
],
"dateFiledFrom": "2000-01-01",
"dateFiledTo": "2020-01-01",
"dateTermedFrom": "2000-01-01",
"dateTermedTo": "2020-01-01",
"dateDismissedFrom": "2000-01-01",
"dateDismissedTo": "2020-01-01",
"dateDischargedFrom": "2000-01-01",
"dateDischargedTo": "2020-01-01",
"federalBankruptcyChapter": [
"7", "15"
],
"natureOfSuit": [
"140", "151"
]
},
"searchName": "Henderson",
"searchType": "PARTY"
}

casefind_url = f'https://{pcl_url}/pcl-public-api/rest/parties/find'

casenum = {
"lastName": "Henderson",
}
case_body = json.dumps(casenum,ensure_ascii=False)

#case = requests.post(casefind_url,headers = pcl_headers, data = case_body) 

info = case.json()
a = datetime.datetime.today()
b =  datetime.datetime.today()

print(json.dumps(info, indent = 4))

str(case)
