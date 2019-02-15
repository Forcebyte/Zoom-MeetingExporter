# Zoom Past Meeting Fetcher
Script that allows a report excel file to be made for a manager's users in Zoom; The report will contain everything that
the dashboard has.
Created with the original goal of allowing Managers to monitor zoom meetings occuring in their department; this will hopefully
be converted into an Executable which will allow managers to drive reports on zoom meetings

## Getting Started

- Generate a API Key and Secret via. https://developer.zoom.us/
- Ensure that you have the latest version of powershell installed, 4.0 is preferred, but anything past Powershell 3.0 will work

### Prerequisites

API Key and Token MUST be inserted into the script before running, Generate an API key and secret at https:developer.zoom.us/ and insert it on line 74/75 of the script
```
$api_key = ''
$api_secret = ''
```

You'll also want to specify the user list in a comma seperated portion
```
$userlist = @(
"useremail1",
"useremail2",
)
```

### API Limits

Due to issues with the API, Zoom only allows queries a minute at a time, and ~ 1 month at a time due to prebuilt constraints.

If you run the script before a minute, or query over a month, you will get a 403 error similar to below
```
{ "error" : { "code" : 403, "message" : "Sorry, the maximum number of api requests are already reached. Please try later." } }
```
The script attempts to circumvent this by waiting a minute using a modified start-sleep command, and automatically waits a minute after this error to avoid erroring out.

Otherwise, results will appear in a .json format, which the script converts into a readable excel format (.csv) 

## Built With

* [JWT Practices](https://jwt.io/) - The API method used for representing security claims in the script
* [u/ping_localhost's Generate-JWT Tokenmaker](https://www.reddit.com/r/PowerShell/comments/8bc3rb/generate_jwt_json_web_token_in_powershell/) - The Function used to generate a JWT Token
* [CTIGeek's Start-sleep script](https://gist.github.com/ctigeek/bd637eeaeeb71c5b17f4) - Used for a more elegant start-sleep command


## Acknowledgments

* u/ping_localhost on reddit for Powershell JSON web-token generator

## Changing Date in Excel Sheet
* Since ISO 8601 comes with a weird time format, use this excel formula to convert the time to a better DateTime object if requested: __=DATEVALUE(MID(A1,1,10))+TIMEVALUE(MID(A1,12,8))__


