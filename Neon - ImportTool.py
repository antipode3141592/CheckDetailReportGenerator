import requests
#import json

#s = requests.Session()

#login
r = requests.post('https://trial.z2systems.com/neonws/services/api/' + 'common/login' + '?login.apiKey=e320cdfe1c8a29cbb572396610e95daa&login.orgid=racctrial')
json_decoded = r.json()
sessionid = json_decoded['loginResponse']['userSessionId']

#do stuff
r = requests.post('https://trial.z2systems.com/neonws/services/api/' + 'donation/listFunds' + '?userSessionId=' + sessionid)
json_decoded = r.json()
print(json_decoded)

#logout
r = requests.post('https://trial.z2systems.com/neonws/services/api/' + 'common/logout' + '?userSessionId=' + sessionid)
print(r.text)
