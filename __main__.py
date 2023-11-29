import lib.getToken

try:
    from pyOutlook import OutlookAccount
except:
    raise Exception("Missing pyOutlook dependency")
try:
    import browser_cookie3
except:
    raise Exception("Missing browser-cookie3 dependency")

print("Python Outlook CLI Client")
print("")
print("1. Personal (outlook.live.com) [Not working]")
print("2. School/Work (outlook.office.com)")
isBusiness = int(input("Enter account type [1,2]: "))==2
cookie = ""
if(isBusiness):
    print("Business type selected")
    print()    
    try:
        cookieJar = browser_cookie3.load(domain_name="outlook.office.com")
        for i in cookieJar:
            if(i.name == "OpenIdConnect.token.v1"):
                cookie = i.value
                break
        if(len(cookie)<1):
            raise Exception("0-byte cookie")
        print("Cookie automatically found from browser")
    except:
        cookie = input("Enter your OpenIdConnect.token.v1 cookie: ")
else:
    print("Personal type selected")
    print()
    try:
        cookieJar = browser_cookie3.load(domain_name="outlook.live.com")
        for i in cookieJar:
            if(i.name == "__Host-MSAAUTHP"):
                cookie = i.value
                break
        if(len(cookie)<1):
            raise Exception("0-byte cookie")
        print("Cookie automatically found from browser")
    except:
        cookie = input("Enter your __Host-MSAAUTHP cookie: ")
print()
userAgent = "Mozilla/5.0 (X11; Linux x86_64; rv:109.0) Gecko/20100101 Firefox/119.0"
print("Getting token, this may take a few seconds")
token = lib.getToken.getOutlookToken(userAgent, cookie, isBusiness) #True for business, false for personal
print("Found token: ",token[0:10]+"..."+token[-10:])

account = OutlookAccount(token)

print(account.inbox())
