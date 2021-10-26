
import requests
import json
import ast
import urllib3
import pyodbc
import pandas as pd
import base64
#import modin.pandas as pd
LMS_db_config = 'DRIVER={SQL Server};SERVER=AC-PC-097;DATABASE=FinanceTeamDB;Trusted_Connection=yes;'
cnxn = pyodbc.connect(LMS_db_config)
##test response
#http = urllib3.PoolManager()
#url_test = http.urlopen('POST', auth_post_url,
#                 headers=headers_dict,
#                 body=encoded_body)
#url_test.read()

#req = requests.Request('POST',auth_post_url,headers=headers_dict,data=encoded_body)
#prepared = req.prepare()

#def pretty_print_POST(req):
#    """
#    At this point it is completely built and ready
#    to be fired; it is "prepared".

#    However pay attention at the formatting used in 
#    this function because it is programmed to be pretty 
#    printed and may differ from the actual request.
#    """
#    print('{}\n{}\r\n{}\r\n\r\n{}'.format(
#        '-----------START-----------',
#        req.method + ' ' + req.url,
#        '\r\n'.join('{}: {}'.format(k, v) for k, v in req.headers.items()),
#        req.body,
#    ))

#pretty_print_POST(prepared)

auth_url = 'qa-login.uscourts.gov'
pacer_url = 'pacer.login.uscourts.gov'
username = 'americashloans'
password = r'i4b4g63r6FVZc!f'
testclientcode = 'testclientcode'
auth_post_url = f"https://{auth_url}/services/cso-auth"
pacer_post_url = f"https://{auth_url}/services/cso-auth"

login_dict = {'loginId': username,'password':password, "redactFlag":1,}

headers_dict = {'Content-type': 'application/json', 'Accept': 'application/json'}
encoded_body = json.dumps(login_dict,ensure_ascii=False)

#Using requests problems with "!" character
r = requests.post(auth_post_url,headers = headers_dict, data = encoded_body.encode('utf-8') 
                  )
auth_dict = r.json()
auth_token = auth_dict['nextGenCSO']


