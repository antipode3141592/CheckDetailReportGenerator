import requests

#login
apikey = 'e320cdfe1c8a29cbb572396610e95daa'
orgid = 'racctrial'
payload = {'login.apiKey' : apikey, 'login.orgid' : orgid}
r = requests.post('https://trial.z2systems.com/neonws/services/api/' + 'common/login', params=payload)
json_decoded = r.json()
sessionid = json_decoded['loginResponse']['userSessionId']
print('session id: {:s}'.format(sessionid))

email = 'seanvokirkpatrick@comcast.net'

#return all funds
payload = {'userSessionId' : sessionid}
r = requests.post('https://trial.z2systems.com/neonws/services/api/' + 'donation/listFunds', params = payload)
print(r.text)

#return all custom objects
payload = {'userSessionId' : sessionid,
           'objectApiName' : 'RACC_Grants_c',
            'customObjectRecord/listCustomObjectRecords' : 'Account Owner',
            'customObjectOutputFieldList.customObjectOutputField.columnName' : 'Account_Owner_c'}
r = requests.post('https://trial.z2systems.com/neonws/services/api/' + 'customObjectRecord/listCustomObjectRecords', params = payload)
print(r.text)

#account lookup by email
payload = {'userSessionId' : sessionid , 
           #'accountSearchCriteria.email': 'seanvokirkpatrick@comcast.net',
           'searches.search.key' : 'Email',
           'searches.search.searchOperator' : 'EQUAL',
           'searches.search.value' : email,
           'outputfields.idnamepair.id' : '',
           'outputfields.idnamepair.name' : 'Account ID'}
r = requests.post('https://trial.z2systems.com/neonws/services/api/' + 'account/listAccounts', params = payload)
json_decoded = r.json()
#print(json_decoded)
accounts = []
for pairs in json_decoded['listAccountsResponse']['searchResults']['nameValuePairs'][0]['nameValuePair']:
    if (pairs['name'] == 'Account ID'):
        #print('name: {:s}'.format(pairs['value']))
        accounts.append(str(pairs['value']))
    else:
        print('error!')

accountid = accounts[0]
print("account id: {0:s}".format(accountid))

#create pledge
payload = {'userSessionId' : sessionid , 
           'pledge.accountId' : accountid,
           'pledge.amount': '1234.56',
           'pledge.campaign.name' : 'Arts Impact Fund 2019',
           'pledge.fund.name' : 'RACC Unrestricted',
           'pledge.date' : '2019-01-28'}
r = requests.post('https://trial.z2systems.com/neonws/services/api/' + 'donation/createPledge', params = payload)
json_decoded = r.json()
pledgeid = json_decoded['createPledge']['pledgeId']
print('pledge created.  id: {0}'.format(pledgeid))

#create pledge payment
payload = {'userSessionId' : sessionid , 
           'pledgeId' : pledgeid, 
           'payment.amount': '100',
           'payment.tenderType.name' : 'Cash'}
r = requests.post('https://trial.z2systems.com/neonws/services/api/' + 'donation/createPledgePayment', params = payload)
json_decoded = r.json()
pledgebalance = json_decoded['createPledgePayment']['balance']
print('payment recorded.  pledge id - {0}, balance - ${1:,.2f}'.format(pledgeid, pledgebalance))

#logout
payload = {'userSessionId' : sessionid }
r = requests.post('https://trial.z2systems.com/neonws/services/api/' + 'common/logout', params = payload)
print(r.text)
