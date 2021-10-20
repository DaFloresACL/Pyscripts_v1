#google api key AIzaSyAkVbIqLQdmn0mBU7hEDBbFAyFetu3nrJ4

g_apikey = "AIzaSyAkVbIqLQdmn0mBU7hEDBbFAyFetu3nrJ4"

import requests
import json
import ast
import pyodbc
import pandas as pd
LMS_db_config = 'DRIVER={SQL Server};SERVER=AC-PC-097;DATABASE=FinanceTeamDB;Trusted_Connection=yes;'
cnxn = pyodbc.connect(LMS_db_config)

afc_custlist_query = ''' 
select a.uniqueID, a.CustomerAddress 
from AFC_CustList a
left outer join AFC_CustList_geometry b
on a.uniqueID = b.uniqueID


where CustomerAddress not like 'po%' and CustomerAddress not like 'p o%' and CustomerAddress not like 
'%box%' and CustomerAddress not like '%address%' and CustomerAddress not like '%addy%'
and b.uniqueID is null
''' 

addresses = pd.read_sql(afc_custlist_query, cnxn)  


testcode = '''
#base_url = f"https://maps.googleapis.com/maps/api/geocode/json?address=1600+Amphitheatre+Parkway,+Mountain+View,+CA&key={g_apikey}"

#response = requests.get(base_url)

#response.status_code

#response.json()
#type(response.json())

#test = response.json()

#test2 = json.dumps(test)
#test['results'][0]
#test
#new = ast.literal_eval(test)

#str_response = test['results'][0]

#str_response.split(',')

#str_response['geometry']['location']
'''
#test_address = "1312 121ST STREET WHITING IN 46394"

#test_addr = test_address.replace(' ','+')
#test_addr
#test_url = f'''https://maps.googleapis.com/maps/api/geocode/json?address={test_addr}&key={g_apikey}'''

#test_response = requests.get(test_url)

#new_test = test_response.json()

#new_response = new_test['results'][0]
#new_response['geometry']['location']['lat']
#new_response['geometry']['location']['lng']

def execProc(sql, values):
    #get the triplist for file creation
    

        try:

            import pypyodbc 
            import pandas as pd
            #SQL Server Native Client 11.0
            #ODBC Driver 13 for SQL Server
            cnxn2 = pypyodbc.connect("Driver={SQL Server Native Client 10.0};"
                                    "Server=AC-PC-097;"
                                    "Database=FinanceTeamDB;"
                                     "Trusted_Connection=yes")
            
            cnxn2.autoCommit = True

            cursor = cnxn2.cursor()
            #print(cursor.autocommit)
            cursor.execute(sql, values).commit()
            

            return 
        except:
            print('error obtaining trips list due to SQL connection or otherwise ODBC connection error-'+sql)
            pass
        return

for index,row in addresses.iterrows():
    try:
        addr = row.CustomerAddress
        api_addr = addr.replace(' ','+')
        api_url = f'''https://maps.googleapis.com/maps/api/geocode/json?address={api_addr}&key={g_apikey}'''
        response = requests.get(api_url)
        results = response.json()
        geometry = results['results'][0]
        lat = geometry['geometry']['location']['lat']
        long = geometry['geometry']['location']['lng']

        values = [row.uniqueID,row.CustomerAddress,lat,long,None]
        execProc('exec insert_AFC_CustList_geometry ?,?,?,?,?',values)
    except:
        values = [row.uniqueID,row.CustomerAddress,None,None,'ASA-TEAM']
        execProc('exec insert_AFC_CustList_geometry ?,?,?,?,?',values)
        continue

import geopy
import geopy.distance

locations_query = '''
select distinct
	LocationName
	,Number
    ,convert(varchar(100),Number + '-' + LocationName) as Location
	,HomeLocation
	,HomeLatitude
	,HomeLongitude
	from v_Location_Distance
'''


uniqueID_query = '''
select
	*
	from AFC_CustList_geometry
    where uniqueID = 'AFC-156158'
'''


locations = pd.read_sql(locations_query, cnxn) 
uniqueID = pd.read_sql(uniqueID_query, cnxn) 

for index_u,row_u in uniqueID.iterrows():
    store_dict = {}
    for index_l,row_l in locations.iterrows():

        coords_1 = (row_u['Latitude'],row_u['Longitude'])
        coords_2 = (row_l['HomeLatitude'],row_l['HomeLongitude'])
        km = round(geopy.distance.distance(coords_1, coords_2).km,2)
        store_dict[row_l.Location] = km * 0.621371
        #print(store_dict)

    store_sort = dict(sorted(store_dict.items(), key=lambda item: item[1]))
    store_sort
    closest = next(iter(store_sort))
    values = [row_u.UniqueID,closest, store_sort[closest]]
    
    execProc ('exec update_AFC_CustList_geometry_ClosestStore ?,?,?',values)



coords_1 = (38.315830,-88.923640)
coords_2 = (38.7137680,-87.7547163 )
km = round(geopy.distance.distance(coords_1, coords_2).km,2)
store_dict[row_l.Location] = km * 0.621371





