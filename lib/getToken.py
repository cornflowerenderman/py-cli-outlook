import requests, urllib.parse, json, time

def getOutlookToken(userAgent, cookie, isBusiness):
    if(userAgent == ""):
        userAgent = None
    if(isBusiness):
        return getOutlookTokenBusiness(userAgent, cookie)
    else:
        return getOutlookTokenPersonal(userAgent, cookie)

def getOutlookTokenBusiness(userAgent,cookie): #Returns token string of length 3864
    cookies = {"OpenIdConnect.token.v1":cookie}
    headers = {
        'action': 'GetAccessTokenforResource',
        'x-owa-urlpostdata': urllib.parse.quote_plus(
            json.dumps({
                "__type": "TokenRequest:#Exchange",
                "Resource": "https://outlook.office.com/"
            }, separators=(',', ':'))
        )
    }
    if(userAgent != None):
        headers['User-Agent'] = userAgent
    url = "https://outlook.office.com/owa/service.svc?action=GetAccessTokenforResource"
    r = requests.post(url,cookies=cookies, headers=headers)
    retry = False
    if(r.status_code == 449): #Hacky solution to get x-owa-canary from OpenIdConnect.token.v1
        for i in r.cookies:     #When a valid x-owa-canary isn't sent, the server just sends a valid one back if the main auth cookie is valid
            if(i.name == "X-OWA-CANARY"):
                headers['x-owa-canary'] = i.value
                retry = True            
                break
    if(retry):
        time.sleep(1)
        r = requests.post(url,cookies=cookies, headers=headers)
    if(r.ok):
        return json.loads(r.text)['AccessToken']
    else:
        return None

def getOutlookTokenPersonal(userAgent,cookie): #Returns token string of length 1112
    cookies = {"__Host-MSAAUTHP":cookie}
    headers = {}
    if(userAgent != None):
        headers['User-Agent'] = userAgent
    url = "https://login.live.com/oauth20_authorize.srf?response_type=token&prompt=none&redirect_uri=https%3A%2F%2Foutlook.live.com%2Fowa%2Fauth%2Fdt.aspx&scope=https%3A%2F%2Faugloop.office.com%2Fv2%2FAugLoop.All&client_id=292841"
    r = requests.get(url, cookies=cookies, headers=headers, allow_redirects=False)
    print(r.headers)
    print(r.text)    
    if(r.status_code == 302):
        dest = r.headers['Location']
        return urllib.parse.unquote(dest.split("=")[1].split("&")[0])
    else:
        return None
