$scriptPath = split-path -parent $MyInvocation.MyCommand.Definition
$ErrorActionPreference = 'SilentlyContinue' | Out-Null﻿
# Removes all Variables in case we're running from ISE
Get-Variable -Exclude PWD,*Preference | Remove-Variable -EA 0
# Clears Screen
Clear

#Start-Sleep is a graphical sleep GUI used to show a progress bar in Shell
function Start-Sleep($seconds) {
    $doneDT = (Get-Date).AddSeconds($seconds)
    while($doneDT -gt (Get-Date)) {
        $secondsLeft = $doneDT.Subtract((Get-Date)).TotalSeconds
        $percent = ($seconds - $secondsLeft) / $seconds * 100
        Write-Progress -Activity "Waiting 60 Seconds due to API Limitation" -Status "Sleeping..." -SecondsRemaining $secondsLeft -PercentComplete $percent
        [System.Threading.Thread]::Sleep(500)
    }
    Write-Progress -Activity "Sleeping" -Status "Sleeping..." -SecondsRemaining 0 -Completed
}

<#Generate JWT is a modified JWT generation function that works with the Zoom API
  In this case we need HS256, as it is a symetric algorithm that will allow
  us to respond to next_page_query returns                                       #>
function Generate-JWT (
   [Parameter(Mandatory = $True)]
   [ValidateSet("HS256", "HS384", "HS512")]
   $Algorithm = $null,
   $type = $null,
   [Parameter(Mandatory = $True)]
   [string]$Issuer = $null,
   [int]$ValidforSeconds = $null,
   [Parameter(Mandatory = $True)]
   $SecretKey = $null
   ){

   $exp = [int][double]::parse((Get-Date -Date $((Get-Date).addseconds($ValidforSeconds).ToUniversalTime()) -UFormat %s)) # Grab Unix Epoch Timestamp and add desired expiration.

   [hashtable]$header = @{alg = $Algorithm; typ = $type}
   [hashtable]$payload = @{iss = $Issuer; exp = $exp}

   $headerjson = $header | ConvertTo-Json -Compress
   $payloadjson = $payload | ConvertTo-Json -Compress

   $headerjsonbase64 = [Convert]::ToBase64String([System.Text.Encoding]::UTF8.GetBytes($headerjson)).Split('=')[0].Replace('+', '-').Replace('/', '_')
   $payloadjsonbase64 = [Convert]::ToBase64String([System.Text.Encoding]::UTF8.GetBytes($payloadjson)).Split('=')[0].Replace('+', '-').Replace('/', '_')

   $ToBeSigned = $headerjsonbase64 + "." + $payloadjsonbase64

   $SigningAlgorithm = switch ($Algorithm) {
       "HS256" {New-Object System.Security.Cryptography.HMACSHA256}
       "HS384" {New-Object System.Security.Cryptography.HMACSHA384}
       "HS512" {New-Object System.Security.Cryptography.HMACSHA512}
   }

   $SigningAlgorithm.Key = [System.Text.Encoding]::UTF8.GetBytes($SecretKey)
   $Signature = [Convert]::ToBase64String($SigningAlgorithm.ComputeHash([System.Text.Encoding]::UTF8.GetBytes($ToBeSigned))).Split('=')[0].Replace('+', '-').Replace('/', '_')

   $token = "$headerjsonbase64.$payloadjsonbase64.$Signature"
   $token
}
 
# Get the date required, and convert to datetime object so compare

# Start Date
$EnteredStartDate = read-host "Please enter a valid start date (I.e. 25 Oct 2018, 25/10/2018, etc.)"
$EnteredStartDate = $EnteredStartDate -as [datetime]
$EnteredStartDate = $EnteredStartDate | get-date -format "yyyy-MM-dd"

# End Date
$EnteredEndDate = read-host "Please enter a valid end date (I.e. 25 Dec 2018, 25/11/2018, etc.)"
$EnteredEndDate = $EnteredEndDate -as [datetime]
$EnteredEndDate = $EnteredEndDate | get-date -format "yyyy-MM-dd"

# To confirm, write the entered start date and end date to make sure the datetime conversion doesn't change anything
write-host -ForegroundColor Green "Date's entered ($EnteredStartDate to $EnteredEndDate) seem to be valid"


<# API Key and Secret used to communicate with Zoom - (Obtain from https://marketplace.zoom.us/user/build)
   Enter the API Key under api_key and api_secret under each, as it is not posted under the Gitlab Repo   #>
$APIURL = "https://api.zoom.us/v2/metrics/meetings"
$api_key = ''
$api_secret = ''

# Generate JWT key that is valid for a minute with the Key and Secret above
$JWT = Generate-JWT -Algorithm 'HS256' -type 'JWT' -Issuer $api_key -SecretKey $api_secret -ValidforSeconds 60

# Body of Invoke-webrequest
$blob = @{
   access_token = "$JWT"
   type = "past"
   from = "$EnteredStartDate"
   to = "$EnteredEndDate"
   next_page_token = ""
   page_size = "300"
}

# Test to see if api_key and api_secret exist
if ($api_key -eq ""){
    write-warning "API Key not entered, Exiting..."
    exit
}elseif($api_secret -eq ""){
    write-warning "API Secret not entered, Exiting..."
    exit
}

# Get the difference between the two date-time objects, if more than 30 days, end script with error
$DateDifference = New-timespan -start $EnteredStartDate -End $EnteredEndDate

if ($DateDifference -gt "32.00:00"){
    Write-Warning "Date difference over 32 days is hard-blocked by the API, exiting"
    exit
}elseif($DateDifference -lt "00.00:00"){
    Write-Warning "Date difference is Invalid, exiting."
    exit
}


# Posts a GET HTTP Method, containing $blob, which is just asking for past meetings;
# Also catches the Zoom API limit with a CODE 403 Write-error function

try{$WebResponse = Invoke-WebRequest -Uri "$APIURL" -Body $blob -Method Get | select-object -expand content | ConvertFrom-Json} catch{Write-Error "CODE 403, Maximum number of API Requests per minute hit"}
    $InitialMeetings = $WebResponse | select-object -Expandproperty meetings


<# If date difference is correct; move onto fetching - If there is a Next-page-token specified, loop until
   we have retrieved all meetings in the given timeframe                                                  #>

While ($WebResponse.next_page_token){
    write-host "Next Page Token is:" $WebResponse.next_page_token 
    write-host "Amount of Total Records are:" $WebResponse.total_records "This will require mutliple API calls"
    write-host "Gathering Data from"$WebResponse.from"To"$WebResponse.to

    <#Since Zoom API has a limitation of 1 API push per minute, wait a minute
    Additionally, note that that Start-sleep is a graphical function defined above#>
    write-host "Waiting 60 seconds"
    Start-sleep 60

    # regenerate JWT since it has expired by now
    $JWT = Generate-JWT -Algorithm 'HS256' -type 'JWT' -Issuer $api_key -SecretKey $api_secret -ValidforSeconds 60

    # Update Blob With new JWT token
    $blob.next_page_token = $Webresponse.next_page_token
    $blob.access_token = $JWT

    # Invoke request again
    $WebResponse = Invoke-WebRequest -Uri "$APIURL" -Body $blob -Method Get | select-object -expand content | ConvertFrom-Json

    # Define the additional meetings that are in the new page
    $Additionalmeetings = $WebResponse | select-object -ExpandProperty meetings

    # If we have to loop again, inform the user, if not, inform the user of the total number of records

    if ($WebResponse.next_page_token){
        write-host "Next Page indicated, System will Loop"}
    else{write-host "No Next Page Token returned by server"}
        write-host "Amount of Total Records are:" $WebResponse.total_records

    # Combine the additional meetings with the initialmeetings
    $InitialMeetings = $InitialMeetings + $Additionalmeetings
    write-host "So far we have gathered"$InitialMeetings.Count"meetings"
}

Write-host "End of data fetching portion - moving on to sorting"

<# Here is a list of all Users that will be filtered through - We will use this to convert all our data to only contain these people in CS.
   We could both the full name and email, as there is an 'email' column and a 'host' column, but email will suffice
   
   Additionally, we could use the Hostname, however that may conflict with other meetings                             #>
$userlist = @(
    "email1@email.com",
    "email2@email.com",
)

<#Loop through each of the entries sent back, and seperate a variable, Targettedmeetings, for requests that match the user list
  Specified above. Additionally convert each meeting to an array instead of a PSCustomObject                                    #>
Foreach ($meeting in $InitialMeetings){
    if ($meeting.email.ToString() -in $userlist){
        $targettedmeetings = $targettedmeetings + [array]$meeting
        }
    }
#

#Exports the targetted meetings into a CSV to where the script is, also exports all meetings for the date range to proof

try{$InitialMeetings | export-csv "$PSScriptRoot\Zoom Mass Export, $EnteredStartdate to $EnteredEndDate .csv" -NoTypeInformation}catch{write-error "Unrecoverable Error - Exception thrown"}
try{$targettedmeetings | export-csv "$PSScriptRoot\Zoom Targeted Export, $EnteredStartdate to $EnteredEndDate .csv" -NoTypeInformation}catch{write-error "Unrecoverable Error - Exception thrown"}

# Prompts user if they want a paste to the console of the zoom results
$PastePrompt = read-host "Would you like to see a paste of the results? (Y/N)"
if ($PastePrompt -like "Y"){
    $targettedmeetings | format-table -AutoSize
}

Read-host "Press any key to continue..." | Out-Null
