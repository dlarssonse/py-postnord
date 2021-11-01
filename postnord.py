import openpyxl
import requests
import json
import time
from pathlib import Path

xlsx_file = Path(".", "test.xlsx")

wb = openpyxl.load_workbook(xlsx_file)
sheet = wb.active

col_firstname = -1
col_lastname = -1
col_email = -1
col_address = -1
col_zip = -1
col_city = -1
authkey = "<PLEASE INSERT AUTH KEY HERE>"

# Get column position
for index, col in enumerate(sheet.iter_cols(1, sheet.max_column)):
  if col[0].value == "Firstname": col_firstname = index
  if col[0].value == "Lastname": col_lastname = index
  if col[0].value == "Email": col_email = index
  if col[0].value == "Address": col_address = index
  if col[0].value == "Zipcode": col_zip = index
  if col[0].value == "City": col_city = index


# Make dictionary
data = []
for i, row in enumerate(sheet.iter_rows(values_only=True)):
  if i > 0: 
    data.append({ "Firstname": row[col_firstname], "Lastname": row[col_lastname], "Email": row[col_email], "Address": row[col_address], "Zipcode": row[col_zip], "City": row[col_city]})

# Output persons
for person in data:
  # Make request
  headers = {
    "Accept": "application/json, text/plain, */*",
    "Accept-Encoding": "gzip, deflate, br",
    "Accept-Language": "sv,en;q=0.9,en-GB;q=0.8,en-US;q=0.7",
    "Authorization": authkey, 
    "Content-Type":"application/json;charset=UTF-8", 
    "Host": "portal.postnord.com", 
    "Origin": "https://portal.postnord.com", 
    "Referer": "https://portal.postnord.com/skickadirekt/user/receivers",
    #"Cookie": "Humany__clientId=e291a69e-adcc-0b95-23be-b0fadd7b4b96",
    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/95.0.4638.54 Safari/537.36 Edg/95.0.1020.40",
    "Sec-Fetch-Dest": "empty",
    "Sec-Fetch-Mode": "cors",
    "Sec-Fetch-Site": "same-origin",
  }
  #pload = {"name":"D L","address":{"street":"L","zipCode":"90432","city":"Umeå","countryCode":"SE"},"country":{"countryCode":"SE","sortIndex":1,"flagUrl":"flags/SE@2x.png","meta":{"euMemberState":True,"udlandZone":"Europa 1"},"callingCode":["46"],"currency":"SEK","postalCodeRegExp":"^\\d\\d\\d ?\\d\\d$","postalCodeExample":"11122","name":"Sverige"}}
  pload = {"name":person["Firstname"] + " " + person["Lastname"], "email": person["Email"],"address":{"street":person["Address"],"zipCode":person["Zipcode"],"city":person["City"],"countryCode":"SE"},"country":{"countryCode":"SE","sortIndex":1,"flagUrl":"flags/SE@2x.png","meta":{"euMemberState":True,"udlandZone":"Europa 1"},"callingCode":["46"],"currency":"SEK","postalCodeRegExp":"^\\d\\d\\d ?\\d\\d$","postalCodeExample":"11122","name":"Sverige"}}
  r = requests.post("https://portal.postnord.com/api/receiver/receivers", data = json.dumps(pload), headers=headers)
  if r.status_code == 200: 
    print("Added '" + person["Firstname"] + " " + person["Lastname"] + "' ")
  else:
    print("Failed to add '" + person["Firstname"] + " " + person["Lastname"] + "'...")
    print(r.json())
  time.sleep(10)


