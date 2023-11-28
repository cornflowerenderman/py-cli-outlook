import lib.getToken
from pyOutlook import OutlookAccount

print("Python Outlook CLI Client")
print("")
print("1. Personal (outlook.live.com) [Not fully working]")
print("2. School/Work (outlook.office.com)")
isBusiness = int(input("Enter account type [1,2]: "))==2
cookie = ""
if(isBusiness):
    print("Business type selected")
    print()
    cookie = input("Enter your OpenIdConnect.token.v1 cookie: ")
else:
    print("Personal type selected")
    print()
    cookie = input("Enter your __Host-MSAAUTHP cookie: ")
print()

userAgent = input("Enter your user-agent (or leave blank for default): ")
if(userAgent == ""):
    userAgent = "Mozilla/5.0 (X11; Linux x86_64; rv:109.0) Gecko/20100101 Firefox/119.0"
print()
print("Getting token, this may take a few seconds")
token = lib.getToken.getOutlookToken(userAgent, cookie, isBusiness) #True for business, false for personal
print("Found token: ",token)

account = OutlookAccount(token)

print(account.inbox())
