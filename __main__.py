import modules.getToken
import base64
import hashlib
import os.path
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
token = modules.getToken.getOutlookToken(userAgent, cookie, isBusiness) #True for business, false for personal
print("Found token: ",token[0:10]+"..."+token[-10:])

account = OutlookAccount(token)

raw_folders = account.get_folders()
folders = {}
for i in raw_folders:
    if(i.name == "Inbox"):
        folders['inbox']=i
    elif(i.name == "Junk Email"):
        folders['junk']=i
    elif(i.name == "Archive"):
        folders['archive']=i
    elif(i.name == "Deleted Items"):
        folders['deleted']=i

print()
print("1. Inbox")
print("2. Junk Mail")
print("3. Archive")
print("4. Deleted Items")
box = int(input("Enter choice [1-4]: "))
name = None
if(box==1):
    name = "inbox"
elif(box==2):
    name = "junk"
elif(box==3):
    name = "archive"
elif(box==4):
    name = "deleted"
print()
print(folders[name].unread_count,"unread,",folders[name].total_items,"total emails in "+name)
messages = folders[name].messages()
for i in messages:
    print("  Sender: "+str(i.sender))
    print("  Subject: "+i.subject)
    print("  Time sent: "+str(i.time_sent)) #Datetime
    importance = "Normal"
    if(i.importance == 0):
        importance = "Low"
    if(i.importance == 2):
        importance = "High"
    print("  Importance: "+importance+(" (Unread)" if i.is_read == False else ""))
    print("    Focused:",i.focused)
    snippet = i.body_preview.replace("\u200c", " ").replace("\n", " ").replace("\r"," ").replace("  "," ").replace("  "," ").replace("  "," ").replace("  "," ").strip() #Clean up string
    if(len(snippet)>50):
        snippet = snippet[:50]+"..."
    print(" ",snippet)
    #print(i.body)
    for a in i.attachments:
        uniqueName = hashlib.sha256(a.outlook_id.encode()).hexdigest()[0:8].upper() #Creates a unique 8-digit id from a semi-random id
        filename = "attachments/["+uniqueName+"] "+a.name
        print("  # Downloading attachment \""+a.name+"\"")
        if(os.path.isfile(filename)):
            print("  # Attachment already downloaded as \""+filename+"\"")
        else:
            raw=a.api_representation()["ContentBytes"]
            attachment = base64.b64decode(raw)
            f = open(filename,"wb")
            f.write(attachment)
            f.close()
            print("  # Attachment saved as \""+filename+"\"")
    print()
