#I2outlook.ps1
#A quick and dirty scipt to import msignite session data into outlook
#you can contact me via hotmail (figure the address out)linkedin,Ignite Portal,twitter @d_bargna or Kaizala
#If you have GIT then please feel free to host it there and I will update my links

#please read all the comments, otherwise the script won't work.

#I'm not a coder, so feel free to do anything you like.

#if this script breaks your computer, or doesnt work, or does horrible things to you then please be aware that it is not my fault and you run this at your own risk
#always read the code and understand it before you execute, look for set- delete- etc, be safe!

#Support?
#if you can't get the data, then check the jwt token, has it expired? Connection closed? TLS issue
#if you have the data but it doesnt write to outlook, check you email address, calendar name and security settings.
#if you still cant figure it out then contact me at hotmail,Ignite portal or Kaizala, but not twitter - cant have a conversations there.


#Prereqs.
#you need to have Outlook installed and ideally running on your pc. this has been tested with outlook 2016 C2R, nothing else has been tested, but should work.
#outlook will prompt you for programatic access, just let it do so. if you are not prompted this could be disabled via GPO :-(... check your security settings in outlook.
#your powershell session must be running with the same identity as your outlook session, not with RunAS
#I'm using v5.1 of powershell, but should be fine with any version.

#How does it work?
#your session data is store in a json object hosted on https://api.myignite.techcommunity.microsoft.com/api/schedule/sessions

#to get your data you need to send a webrequest with your personal jwt(auth token) in the header.
#to get the jwt I could have written a bit of code to do the auth request, but its easier to grab it from your browser..quick and dirty remember.

#once we have this the rest is easy, we just read the json and stick it into outlook.

#IMPORTANT, you need to do this!
#How to get your auth cookie?
#Note the token seems to expire after 10 hours or so!
#logon to myignite with your creds. 
#In your prefered browser start dev tools,this F12 in edge,firefox, goto Debugger/storage, Cookies and then find the ignite.token cookie
#copy the cookie value, depending on browser remove "ignite.token" from the start and then", myignite.tech....18/09...." from the end so you are left with a base64 string - no commas or other stuff.

#when done it should look like this. dont use this one its not real!

#"eyJ0eXAiOiJKV1QiLCJllGJiOiJIUzI1NiJ9.eyQzL29wZXMiOlsibXlpZ25pdGUiLCJzY2hlZHVsZXIiLCJtZWV0aW5ncyIsImlucGVyc29uIl0sImVpZCI6IjA4NzA2YjhhLWVlNjQtNGYzNy04MDdlLThjMzgxMTllNWJhNCIsImh1YmIiOiJKc21KMHdBZDlLYzJ1S0Y5d2E3ZlVCVEprTVdTOThOTnZHWUJYc21sU01zVVBGdlcrTmVkc3lXZG4zeE80cG5LZDltYm9jUUdhM0QycjJOUFpMa1kxdHdoRzNPcGZ1Qm03eU5KUC9HT1ZOVWhLL2xrdGFTb1lQOGgyL0FPSi94QXBmVXRRVmRYSUFWSmJ3QnBJcDBHbGRTdlNhUDRJOTBDbVRNSytsOWRkbTNpbTdraURSNWFhdzVpeDZsMmttd2xpMlZMYlhPOGdTS2ROd3lpbGQ4ZDcrMGg4TzgrUkk1NUhSaEJIOFSAD3SDaWyycWd1ek9IZmFodWlMYkJORThobkp2OEUrcVArNStXWG5DSWUzTW55clRjQUQ2Uy9IRG4zWHJpZVBaNGRyRC9CTytBRmtCQm05Z0ZrSWJYdVJyek9hc0hxSlJzSHJvTUNreDM1MlFQTUE9PSIsIm5hbWUiOiJEYXZpZCBCYXJnbmEiLCJmaXJzdE5hbWUiOiJEYXZpZCIsImxhc3ROYW1lIjoiQmFyZ25hIiwiZW1haWwiOiJkYXZpZC5iYXJnbmFAZXNvLm9yZyIsInJpZCI6IjM2OTkzOSIsImFpZCI6Ik1TSSBBdHRlbmRlZSBDdXN0b21lciBcdTAwMjYgUGFydG5lciIsInZlcnNpb24iOiIyLjAuMC40IiwiaWduaXRlLnRva2VuLmRpc2N1c3Npb24iOiJUcGxKMFZWQklpa3o0WnZOR2FGbjFkOGtMclo4dm1nSFg2MmIwTDl2SjFVPSIsInRjaWQiOiI0MDIzNiIsImlzcyI6Im23SWduaXRlIiwiZXhwIjoxNTM3Mjk2OTY2LCJuYmYiOjE1MzcyNTM3NjZ9.LW-UZKl_zA6XjO2Cwfigv3GVyCbwhUSSHAPC03QWERTY"


#Last but not Least!
# you need to do fill in the bits of data in this section.
#---------YOUR DATA GOES HERE
# and you need to uncomment the line below this one further down in the script to write your data to outlook.
#--------UnComment to add to Outlook
#go can skip to fill in your details now, or read on.


# we will validate the cookie later using https://gallery.technet.microsoft.com/JWT-Token-Decode-637cf001
#once we have the cookie then we just grab the data and the rest is easy.

#This is a sample if the JSON data from the api.

#@search.score             : 1.0
#sessionId                 : 64616
#sessionInstanceId         : 64616
#sessionCode               : BRK2112
#title                     : Windows Subsystem for Linux and Enterprise
#description               : In this session, learn ways to leverage advancements in command line tools and WSL. We walk you through a variety of ways to use Linux tools on Windows 10 including 
#                            using WSL with Visual Studio Code to target a Linux environment, new integrations into the Windows system, and WSL in your enterprise environment, among other demos.
#roomId                    : 2523
#location                  : OCCC W303
#speakerIds                : {423051, 390871}
#speakerNames              : {Craig Loewen, Tara Raj}
#speakerCompanies          : {Microsoft, Microsoft}
#startDateTime             : 2018-09-28T16:30:00+00:00
#endDateTime               : 2018-09-28T17:45:00+00:00
#level                     : Intermediate (200)
#format                    : Session
#products                  : {Windows	Windows Developer, Windows}
#durationInMinutes         : 75
#sessionType               : Breakout: 75 Minute
#isMandatory               : False
#techCommunityDiscussionId : 209740
#learningPath              : {}
#contentCategory           : {WDG , WDG 	 Windows Dev}


#---------Functions
# JWT from #https://gallery.technet.microsoft.com/JWT-Token-Decode-637cf001
function Convert-FromBase64StringWithNoPadding([string]$data)
{
    $data = $data.Replace('-', '+').Replace('_', '/')
    switch ($data.Length % 4)
    {
        0 { break }
        2 { $data += '==' }
        3 { $data += '=' }
        default { throw New-Object ArgumentException('data') }
    }
    return [System.Convert]::FromBase64String($data)
}

function Decode-JWT([string]$rawToken)
{
    $parts = $rawToken.Split('.');
    $headers = [System.Text.Encoding]::UTF8.GetString((Convert-FromBase64StringWithNoPadding $parts[0]))
    $claims = [System.Text.Encoding]::UTF8.GetString((Convert-FromBase64StringWithNoPadding $parts[1]))
    $signature = (Convert-FromBase64StringWithNoPadding $parts[2])

    $customObject = [PSCustomObject]@{
        headers = ($headers | ConvertFrom-Json)
        claims = ($claims | ConvertFrom-Json)
        signature = $signature
    }

    Write-Verbose -Message ("JWT`r`n.headers: {0}`r`n.claims: {1}`r`n.signature: {2}`r`n" -f $headers,$claims,[System.BitConverter]::ToString($signature))
    return $customObject
}

function Get-JwtTokenData
{
    [CmdletBinding()]  
    Param
    (
        # Param1 help description
        [Parameter(Mandatory=$true)]
        [string] $Token,
        [switch] $Recurse
    )
    
    if ($Recurse)
    {
        $decoded = [System.Text.Encoding]::UTF8.GetString([System.Convert]::FromBase64String($Token))
        Write-Host("Token") -ForegroundColor Green
        Write-Host($decoded)
        $DecodedJwt = Decode-JWT -rawToken $decoded
    }
    else
    {
        $DecodedJwt = Decode-JWT -rawToken $Token
    }
    Write-Host("Token Values") -ForegroundColor Green
    Write-Host ($DecodedJwt | Select headers,claims | ConvertTo-Json)
    return $DecodedJwt
}
#----- end JWT from #https://gallery.technet.microsoft.com/JWT-Token-Decode-637cf001


#add stuff to outlook.
#using a com object isnt pretty but it works. if it doesnt work for you then check out the docs..
#https://docs.microsoft.com/en-us/dotnet/api/microsoft.office.interop.outlook.appointmentitemclass?view=outlook-pia#properties
function add-I2outlook{
    param ($jsonitem,$calname,$email,$outlook)
    #const for the default calendar.
    $olFolderCalendar=9
    Add-type -assembly "Microsoft.Office.Interop.Outlook" | out-null
    #declare mapi namespace
    $olNS=$outlook.getnamespace("MAPI")
    #find the mailbox to update the calendr in from the email address
    $Store = $olNS.Stores | ? {$_.displayname -eq $email}
    #not really needed!
    $calendar=$null
    #of the calname is empty then we use the default outlook calendar for the mailbox, otherwise we will try and find the named calendar.
    if ([string]::IsNullOrEmpty($calname))
      {
        $calendar=$store.GetDefaultFolder([Microsoft.Office.Interop.Outlook.OlDefaultFolders]::olFolderCalendar)
      }
    else
      {
        $calendar=$store.GetDefaultFolder([Microsoft.Office.Interop.Outlook.OlDefaultFolders]::olFolderCalendar).Folders($calname)
      }

    #outlook appointment item.   
    $olAptitem=$calendar.Items.add("IPM.Appointment")
    $olAptitem.start=[datetime]$jsonitem.startDateTime
    $olAptitem.duration=$jsonitem.durationInMinutes
    $olAptitem.subject="$($jsonitem.sessionCode) - $($jsonitem.title)"
    $olAptitem.location="$($jsonitem.roomId)"
    $olAptitem.body="$($jsonitem.description)"
    $olAptitem.Save()
}

#-END Functions


#---------YOUR DATA GOES HERE

#-Put your auth token here - between the single quotes, do not share your token with anyone, do not distribute it :-)
# the token expires, so you may need to get a new one every 12 hours
$xjwt=''

#-put the email address where your outlook calendar lives here.
$email="someone@somedomain.com"

#-put the calendar name here
#I would highly recommend using a separate calendar for the ignite stuff, duplicates will be created the more times you run the script!!!
#data will be writen to your default calendar if you comment out $calname - not recommended
#create a new calendar in outlook, call it Ignite, or whatever you want (make sure it is in the same mailbox as your email address!)
#set the value of $Calname= the calendar you created
$calname="Ignite"

#--------dont forget to uncomment the line below this line (#--------UnComment to add to Outlook), further down in the script!


## Lets Go!
#first we check the security token
$token=$null
$token=Get-JwtTokenData $xjwt



#other intersted URIsa
#meetings https://api.myignite.techcommunity.microsoft.com/api/schedule/meetings

if ($token.claims.scopes -match "myignite")
  {
    write-output "Bingo! your token looks good...lets grab the data"
    #now we grab the json data :-)
    #set TLS to v 1.2 otherwise connection will be closed!
    [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
    $uri="https://api.myignite.techcommunity.microsoft.com/api/schedule/sessions"
    $headers = New-Object "System.Collections.Generic.Dictionary[[String],[String]]"
    $headers.Add('x-jwt',$xjwt)
    try
      {
        $jsondata =convertfrom-json -InputObject $(Invoke-WebRequest -Uri $uri -Headers $headers -UseBasicParsing -erroraction Stop).content
      }
    catch 
      {
        write-output "I cannot get the JSON data from the api server."
        write-output "$($error[0])"
        break
      }
  #got some data? make sure its not empty
    if($jsondata -ne $null)
      {
      #load outlook type library and Com object
      #Add-type -assembly "Microsoft.Office.Interop.Outlook" | out-null
      $outlook = new-object -comobject "Outlook.Application"
      #got some data, display it then chuck it into outlook
        foreach ($item in $jsondata)
          {
            
            #just try to get the data first and then uncomment the below - 
            write-output $item
            
	    #--------UnComment to add to Outlook
            #uncomment the next line to try and put the data into your outlook calendar.outlook will prompt you to allow programable access, do it for 10 mins or so.
            add-I2outlook -jsonitem $item -calname $calname -email $email -outlook $outlook
            }
    }
    else
      {
        write-output "No JSON data, sorry!!"
      }
  }
    else
    {
      write-output "looks like the auth token is broken/invalid...sorry. have you removed all the junk from the start end?"
  }
