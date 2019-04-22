# Must have the following files:
# applications.tsv
# campaigns.tsv
# sentemails.tsv
# tokens.tsv
# webhits.tsv
# events.tsv (Optional with parameter below. Requires custom Cobalt Strike phishing profiler aggressor script)
# "Employee Details - PhishMe Input Format.csv" (Optional with parameter below. Includes additional employee data in the same format that PhishMe requires.)

# Set the scenario number
$scenario = 1

# Set the directory for the Cobalt Strike TSV files (Last character must be a slash "\")
$cobaltStrikeOutputTsvDirectory = "Example Input - Cobalt Strike TSV Files" + "\"

# Set the directory for the custom employee details file that is in PhishMe format (Last character must be a slash "\")
$phishMeInputCsvDirectory = "Example Input - Employee Details in PhishMe Input Format" + "\"

# Set the directory for the phishing reports (Last character must be a slash "\")
$phishingReportsOutputFolder = "Example Output - Phishing Reports" + "\"

# Limit the results to 100 for testing
$limitResults = $FALSE

# Lookup custom employee information such as the employee's location and employee's manager
$lookupAdditionalEmployeeDetailsUsingPhishMeFormat = $TRUE

# Lookup user agent from custom cna script
$lookupUserAgentFromCustomCnaScript = $TRUE

# Write the start time
Write-Host (Get-Date -f yyyy_MM_dd-HH:mm:ss)"- Start"

# Check if the phishing reports directory exists
if(!(test-path $phishingReportsOutputFolder)){

    # Create the report directory
    New-Item -ItemType Directory -Force -Path $phishingReportsOutputFolder | Out-Null
}

# Convert the Unix timestamp
Function ConvertFromUnixTimestamp($unixDate){
    [timezone]::CurrentTimeZone.ToLocalTime(([datetime]'1/1/1970').AddSeconds($unixDate))
}

# Save a report
Function SaveReport($filePath, $data){

    # Set the path for the csv file
    $csvFilePath = $filePath + ".csv"

    # Export the file to a csv
    $data | Export-Csv -noType $csvFilePath

    # Set the path for the xlsx file
    $xlsxFilePath = $filePath + ".xlsx"

    # Create an Excel object
    $Excel = New-Object -ComObject Excel.Application

    # Show the Excel window
    $Excel.Visible = $FALSE

    # Do not show the "Save/Do Not Save" window
    $Excel.DisplayAlerts = $FALSE

    # Get the current directory
    $path = iex pwd

    # Open the workbook
    $Workbook = $Excel.Workbooks.Open($path.Path + "\" + $csvFilePath)

    # Open the worksheet
    $Worksheet = $Workbook.Worksheets.Item(1)

    # Save the workbook as an xlsx file
    $Workbook.SaveAs($path.Path + "\" + $xlsxFilePath, 51)

    # Quit Excel
    $Excel.Quit()

    # Remove the temporary csv file
    Remove-Item $csvFilePath
}

# Search will take ~10 minutes for 10,000 emails instead of ~80 minutes using the Where-Object within the for loop
$Source = @"
using System;
using System.Management.Automation;
namespace FastSearch{
    public static class Search{
        public static object Find(PSObject[] collection, string column, string data){
            foreach(PSObject item in collection){
                if (item.Properties[column].Value.ToString() == data) { return item; }
            }
            return null;
        }
    }
    public static class Count{
        public static object Find(PSObject[] collection, string column, string data){
            int count;
            count = 0;
            foreach(PSObject item in collection){
                if (item.Properties[column].Value.ToString() == data) { count += 1; }
            }
            return count;
        }
    }
}
"@
Add-Type -ReferencedAssemblies $Assem -TypeDefinition $Source -Language CSharp

# Set the path for the tokens tsv file
$tokensTsvPath = $cobaltStrikeOutputTsvDirectory + "tokens.tsv"

# Check for the tokens tsv file
if((Test-Path $tokensTsvPath) -eq $false) {'"' + $tokensTsvPath + '"' + " not found."; exit}

# Set the header for the tokens.tsv file
$tokensTsvHeader = "Tokens Token", "Tokens Email", "Tokens Campaign ID"

# Import the tokens.tsv without the header row
$tokensTsvFile = Import-Csv -Delimiter "`t" -Path $tokensTsvPath -Header $tokensTsvHeader | Where-Object "Tokens Token" -ne "token"

# The tokens.tsv file should be in this format
# token    email              cid
# 123abc   name@example.com   12345abcde

# set the path for the sentemails tsv file
$sentemailsTsvPath = $cobaltStrikeOutputTsvDirectory + "sentemails.tsv"

# check for sentemails tsv file
if((Test-Path $sentemailsTsvPath) -eq $false) {'"' + $sentemailsTsvPath + '"' + " not found."; exit}

# set the header for the sentemails.tsv file
$sentemailsTsvHeader = "Sentemails Token", "Sentemails Campaign ID", "Sentemails When", "Sentemails Status", "Sentemails Status Reason"

# Import the sentemails.tsv without the header row
$sentemailsTsvFile = Import-Csv -Delimiter "`t" -Path $sentemailsTsvPath -Header $sentemailsTsvHeader | Where-Object "Sentemails Token" -ne "token"

# File should be in this format
# token    cid          when        status    data
# 123abc   12345abcde   946768081   SUCCESS   250 Ok

# Add columns for sentemails tsv file
$tokensTsvFile | Add-Member -name "Sentemails When" -value "" -MemberType NoteProperty
$tokensTsvFile | Add-Member -name "Sentemails When Converted" -value "" -MemberType NoteProperty
$tokensTsvFile | Add-Member -name "Sentemails Status" -value "" -MemberType NoteProperty
$tokensTsvFile | Add-Member -name "Sentemails Status Reason" -value "" -MemberType NoteProperty

# Set the path for the campaigns tsv file
$campaignsTsvPath = $cobaltStrikeOutputTsvDirectory + "campaigns.tsv"

# Check for campaigns tsv file
if((Test-Path $campaignsTsvPath) -eq $false) {'"' + $campaignsTsvPath + '"' + " not found."; exit}

# Set the header for the campaigns.tsv file
$campaignsTsvHeader = "Campaigns Campaign ID", "Campaigns When", "Campaigns URL", "Campaigns Attachment", "Campaigns Template", "Campaigns Subject"

# Import the campaigns.tsv without the header row
$campaignsTsvFile = Import-Csv -Delimiter "`t" -Path $campaignsTsvPath -Header $campaignsTsvHeader | Where-Object "Campaigns Campaign ID" -ne "cid"

# File should be in this format
# cid          when        url      attachment   template        subject
# 12345abcde   946768081   /login                /template.txt   Click Me

# Add columns from campaigns tsv file
$tokensTsvFile | Add-Member -name "Campaigns When" -value "" -MemberType NoteProperty
$tokensTsvFile | Add-Member -name "Campaigns When Converted" -value "" -MemberType NoteProperty
$tokensTsvFile | Add-Member -name "Campaigns URL" -value "" -MemberType NoteProperty
$tokensTsvFile | Add-Member -name "Campaigns Attachment" -value "" -MemberType NoteProperty
$tokensTsvFile | Add-Member -name "Campaigns Template" -value "" -MemberType NoteProperty
$tokensTsvFile | Add-Member -name "Campaigns Subject" -value "" -MemberType NoteProperty

# Set the path for the webhits tsv file
$webhitsTsvPath = $cobaltStrikeOutputTsvDirectory + "webhits.tsv"

# Check for webhits tsv file
if((Test-Path $webhitsTsvPath) -eq $false) {'"' + $webhitsTsvPath + '"' + "not found."; exit}

# Set the header for the webhits.tsv file
$webhitsTsvHeader = "Webhits When", "Webhits Token", "Webhits Data"

# Import the webhits.tsv without the header row
$webhitsTsvFile = Import-Csv -Delimiter "`t" -Path $webhitsTsvPath -Header $webhitsTsvHeader | Where-Object "Webhits When" -ne "when"

# File should be in this format
# when        token    data
# 946768081   123abc   visit to /login (profiler System Profiler. Redirects to /portal) by 1.2.3.4

# Add columns from webhits tsv file
$tokensTsvFile | Add-Member -name "Webhits When" -value "" -MemberType NoteProperty
$tokensTsvFile | Add-Member -name "Webhits When Converted" -value "" -MemberType NoteProperty
$tokensTsvFile | Add-Member -name "Webhits Data" -value "" -MemberType NoteProperty
$tokensTsvFile | Add-Member -name "Webhits Visit" -value "" -MemberType NoteProperty
$tokensTsvFile | Add-Member -name "Webhits IP" -value "" -MemberType NoteProperty

# Add columns in case there are multiple clicks for a token from webhits tsv file
$tokensTsvFile | Add-Member -name "Webhits When Multiple" -value "" -MemberType NoteProperty
$tokensTsvFile | Add-Member -name "Webhits When Converted Multiple" -value "" -MemberType NoteProperty
$tokensTsvFile | Add-Member -name "Webhits Data Multiple" -value "" -MemberType NoteProperty
$tokensTsvFile | Add-Member -name "Webhits Visit Multiple" -value "" -MemberType NoteProperty
$tokensTsvFile | Add-Member -name "Webhits IP Multiple" -value "" -MemberType NoteProperty

# set the path for the applications tsv file
$applicationsTsvPath = $cobaltStrikeOutputTsvDirectory + "applications.tsv"

# check for applications tsv file
if((Test-Path $applicationsTsvPath) -eq $false) {'"' + $applicationsTsvPath + '"' + " not found."; exit}

# set the header for the applications.tsv file
$applicationsTsvHeader = "Applications External", "Applications Internal", "Applications Application", "Applications Version", "Applications Date", "Applications ID"

# Import the applications.tsv without the header row
$applicationsTsvFile = Import-Csv -Delimiter "`t" -Path $applicationsTsvPath -Header $applicationsTsvHeader | Where-Object "Applications ID" -ne "id"

# File should be in this format
# external   internal      application      version     date        id
# 1.2.3.4    192.168.0.1   Windows 10 *64               946768081   123abc
# 1.2.3.4    192.168.0.1   Chrome *64       100.1.2.3   946768081   123abc

# Add columns for applications tsv file
$tokensTsvFile | Add-Member -name "Applications External" -value "" -MemberType NoteProperty
$tokensTsvFile | Add-Member -name "Applications Internal" -value "" -MemberType NoteProperty
$tokensTsvFile | Add-Member -name "Applications Application" -value "" -MemberType NoteProperty
$tokensTsvFile | Add-Member -name "Applications Version" -value "" -MemberType NoteProperty
$tokensTsvFile | Add-Member -name "Applications Date" -value "" -MemberType NoteProperty
$tokensTsvFile | Add-Member -name "Applications Date Converted" -value "" -MemberType NoteProperty
$tokensTsvFile | Add-Member -name "Applications ID" -value "" -MemberType NoteProperty

# Add columns in case there are multiple applications for a token from applications tsv file
$tokensTsvFile | Add-Member -name "Applications External Multiple" -value "" -MemberType NoteProperty
$tokensTsvFile | Add-Member -name "Applications Internal Multiple" -value "" -MemberType NoteProperty
$tokensTsvFile | Add-Member -name "Applications Application Multiple" -value "" -MemberType NoteProperty
$tokensTsvFile | Add-Member -name "Applications Version Multiple" -value "" -MemberType NoteProperty
$tokensTsvFile | Add-Member -name "Applications Date Multiple" -value "" -MemberType NoteProperty
$tokensTsvFile | Add-Member -name "Applications Date Converted Multiple" -value "" -MemberType NoteProperty
$tokensTsvFile | Add-Member -name "Applications ID Multiple" -value "" -MemberType NoteProperty

# Check if the user agent should be looked up
if($lookupUserAgentFromCustomCnaScript -eq $TRUE){

    # Set the path for the events tsv file
    $eventsTsvPath = $cobaltStrikeOutputTsvDirectory + "events.tsv"

    # Check for events tsv file
    if((Test-Path $eventsTsvPath) -eq $false) {'"' + $eventsTsvPath + '"' + " not found."; exit}

    # Set the header for the events.tsv file
    $eventsTsvHeader = "Events When", "Events Type", "Events Timestamp", "Events IP", "Events Token", "Events Email", "Events Method", "Events Link", "Events User Agent"

    # Import the events.tsv file and only get web types, not "received system profile" types
    $eventsTsvFile = Import-Csv -Delimiter "`t" -Path $eventsTsvPath -Header $eventsTsvHeader | Where-Object "Events Type" -eq "web"

    # File should be in this format
    # when        data
    # 946768081   received system profile (5 applications)
    # 946768081   web   01/01/2000 23:08:01   1.2.3.4   123abc   First.Last1@example.com   GET   /login   UserAgentString

    # Add columns from events tsv file
    $tokensTsvFile | Add-Member -name "Events When" -value "" -MemberType NoteProperty
    $tokensTsvFile | Add-Member -name "Events When Converted" -value "" -MemberType NoteProperty
    $tokensTsvFile | Add-Member -name "Events Type" -value "" -MemberType NoteProperty
    $tokensTsvFile | Add-Member -name "Events IP" -value "" -MemberType NoteProperty
    $tokensTsvFile | Add-Member -name "Events Method" -value "" -MemberType NoteProperty
    $tokensTsvFile | Add-Member -name "Events Link" -value "" -MemberType NoteProperty
    $tokensTsvFile | Add-Member -name "Events User Agent" -value "" -MemberType NoteProperty

    # Add columns in case there are multiple clicks for a token from events tsv file
    $tokensTsvFile | Add-Member -name "Events When Multiple" -value "" -MemberType NoteProperty
    $tokensTsvFile | Add-Member -name "Events When Converted Multiple" -value "" -MemberType NoteProperty
    $tokensTsvFile | Add-Member -name "Events Type Multiple" -value "" -MemberType NoteProperty
    $tokensTsvFile | Add-Member -name "Events IP Multiple" -value "" -MemberType NoteProperty
    $tokensTsvFile | Add-Member -name "Events Method Multiple" -value "" -MemberType NoteProperty
    $tokensTsvFile | Add-Member -name "Events Link Multiple" -value "" -MemberType NoteProperty
    $tokensTsvFile | Add-Member -name "Events User Agent Multiple" -value "" -MemberType NoteProperty
}

# Check if the employee's custom information should be looked up
if($lookupAdditionalEmployeeDetailsUsingPhishMeFormat -eq $TRUE){

    # Set the path for the Employee Details csv file
    $phishMeInputCsvPath = $phishMeInputCsvDirectory + "Employee Details - PhishMe Input Format.csv"

    # Check for the Employee Details csv file
    if((Test-Path $phishMeInputCsvPath) -eq $false) {'"' + $phishMeInputCsvPath + '"' + " not found."; exit}

    # Set the header for the webhits.tsv file
    $phishMeInputCsvHeader = "PhishMeInput Email", "PhishMeInput Name", "PhishMeInput Department", "PhishMeInput Location"

    # Import the phishMeInput.csv file
    $phishMeInputCsvFile = Import-Csv -Path $phishMeInputCsvPath -Header $phishMeInputCsvHeader | Where-Object "PhishMeInput Email" -ne "email"

    # File should be in this format (This is the format that PhishMe takes)
    # email              name         department           location              time_zone
    # name@example.com   First Last   Employee's Manager   Employee's Location

    # Add custom employee data columns from Employee Details csv file
    $tokensTsvFile | Add-Member -name "PhishMeInput Name" -value "" -MemberType NoteProperty
    $tokensTsvFile | Add-Member -name "PhishMeInput Department" -value "" -MemberType NoteProperty
    $tokensTsvFile | Add-Member -name "PhishMeInput Location" -value "" -MemberType NoteProperty
}

# Set the total row count
$rows = $tokensTsvFile.Count

# Set the row count to zero
$rowCount = 0

# Loop through each row
forEach($tokensTsvRow in $tokensTsvFile){

    # Lookup the sentemails tsv row
    # $sentemailsTsvRow = $sentemailsTsvFile | Where-Object {($_."Sentemails Token" -eq $tokensTsvRow."Tokens Token")} # -And ($_."Sentemails Campaign ID" -eq $tokensTsvRow."Tokens Campaign ID")}
    $sentemailsTsvRow = [FastSearch.Search]::Find($sentemailsTsvFile, "Sentemails Token", $tokensTsvRow."Tokens Token")

    # Add sentemails data
    $tokensTsvRow."Sentemails When" = $sentemailsTsvRow."Sentemails When"
    $tokensTsvRow."Sentemails When Converted" = ConvertFromUnixTimestamp ($sentemailsTsvRow."Sentemails When" / 1000)
    $tokensTsvRow."Sentemails Status" = $sentemailsTsvRow."Sentemails Status"
    $tokensTsvRow."Sentemails Status Reason" = $sentemailsTsvRow."Sentemails Status Reason"

    # Lookup the campaigns tsv row
    # $campaignsTsvRow = $campaignsTsvFile | Where-Object {($_."Campaigns Campaign ID" -eq $tokensTsvRow."Tokens Campaign ID")}
    $campaignsTsvRow = [FastSearch.Search]::Find($campaignsTsvFile, "Campaigns Campaign ID", $tokensTsvRow."Tokens Campaign ID")

    # Add campaigns data
    $tokensTsvRow."Campaigns When" = $campaignsTsvRow."Campaigns When"
    $tokensTsvRow."Campaigns When Converted" = ConvertFromUnixTimestamp ($campaignsTsvRow."Campaigns When" / 1000)
    $tokensTsvRow."Campaigns URL" = $campaignsTsvRow."Campaigns URL"
    $tokensTsvRow."Campaigns Attachment" = $campaignsTsvRow."Campaigns Attachment"
    $tokensTsvRow."Campaigns Template" = $campaignsTsvRow."Campaigns Template"
    $tokensTsvRow."Campaigns Subject" = $campaignsTsvRow."Campaigns Subject"

    # Lookup the applications tsv row
    # $applicationsTsvRow = $applicationsTsvFile | Where-Object {($_."Applications ID" -eq $tokensTsvRow."Tokens Token")}
    $applicationsTsvRow = [FastSearch.Search]::Find($applicationsTsvFile, "Applications ID", $tokensTsvRow."Tokens Token")

    # Check if this person/row has a click
    if($applicationsTsvRow -ne $null){

        # Get count of applications for person/row
        $applicationsTsvRowCount = [FastSearch.Count]::Find($applicationsTsvFile, "Applications ID", $tokensTsvRow."Tokens Token")

        # Check if there was only one applications row for the given token
        if($applicationsTsvRowCount -eq 1){

            # Add applications data
            $tokensTsvRow."Applications External" = $applicationsTsvRow."Applications External"
            $tokensTsvRow."Applications Internal" = $applicationsTsvRow."Applications Internal"
            $tokensTsvRow."Applications Application" = $applicationsTsvRow."Applications Application"
            $tokensTsvRow."Applications Version" = $applicationsTsvRow."Applications Version"
            $tokensTsvRow."Applications Date" = $applicationsTsvRow."Applications Date"
            $tokensTsvRow."Applications Date Converted" = ConvertFromUnixTimestamp ($applicationsTsvRow."Applications Date" / 1000)
            $tokensTsvRow."Applications ID" = $applicationsTsvRow."Applications ID"
        }

        # There are multiple applications rows for the given token
        else{

            # Use slower Where-Object to get multiple clicks
            $applicationsTsvRow = $applicationsTsvFile | Where-Object {($_."Applications ID" -eq $tokensTsvRow."Tokens Token")}

            # Add first applications data
            $tokensTsvRow."Applications External" = $applicationsTsvRow[0]."Applications External"
            $tokensTsvRow."Applications Internal" = $applicationsTsvRow[0]."Applications Internal"
            $tokensTsvRow."Applications Application" = $applicationsTsvRow[0]."Applications Application"
            $tokensTsvRow."Applications Version" = $applicationsTsvRow[0]."Applications Version"
            $tokensTsvRow."Applications Date" = $applicationsTsvRow[0]."Applications Date"
            $tokensTsvRow."Applications Date Converted" = ConvertFromUnixTimestamp ($applicationsTsvRow[0]."Applications Date" / 1000)
            $tokensTsvRow."Applications ID" = $applicationsTsvRow[0]."Applications ID"

            # Loop through each application row for the given token
            forEach($applicationMatch in $applicationsTsvRow){

                # Add all applications data
                $tokensTsvRow."Applications External Multiple" += $applicationMatch."Applications External" + " : "
                $tokensTsvRow."Applications Internal Multiple" += $applicationMatch."Applications Internal" + " : "
                $tokensTsvRow."Applications Application Multiple" += $applicationMatch."Applications Application" + " : "
                $tokensTsvRow."Applications Version Multiple" += $applicationMatch."Applications Version" + " : "
                $tokensTsvRow."Applications Date Multiple" += $applicationMatch."Applications Date" + " : "
                $tokensTsvRow."Applications Date Converted Multiple" += (ConvertFromUnixTimestamp ($applicationMatch."Applications Date" / 1000)).ToString("MM/dd/yyyy HH:mm:ss") + " : "
                $tokensTsvRow."Applications ID Multiple" += $applicationMatch."Applications ID" + " : "
            }

            # Remove the last " : "
            $tokensTsvRow."Applications External Multiple" = $tokensTsvRow."Applications External Multiple".Substring(0, $tokensTsvRow."Applications External Multiple".Length-3)
            $tokensTsvRow."Applications Internal Multiple" = $tokensTsvRow."Applications Internal Multiple".Substring(0, $tokensTsvRow."Applications Internal Multiple".Length-3)
            $tokensTsvRow."Applications Application Multiple" = $tokensTsvRow."Applications Application Multiple".Substring(0, $tokensTsvRow."Applications Application Multiple".Length-3)
            $tokensTsvRow."Applications Version Multiple" = $tokensTsvRow."Applications Version Multiple".Substring(0, $tokensTsvRow."Applications Version Multiple".Length-3)
            $tokensTsvRow."Applications Date Multiple" = $tokensTsvRow."Applications Date Multiple".Substring(0, $tokensTsvRow."Applications Date Multiple".Length-3)
            $tokensTsvRow."Applications Date Converted Multiple" = $tokensTsvRow."Applications Date Converted Multiple".Substring(0, $tokensTsvRow."Applications Date Converted Multiple".Length-3)
            $tokensTsvRow."Applications ID Multiple" = $tokensTsvRow."Applications ID Multiple".Substring(0, $tokensTsvRow."Applications ID Multiple".Length-3)
        }
    }

    # Lookup the webhits tsv row
    # $webhitsTsvRow = $webhitsTsvFile | Where-Object {($_."Webhits Token" -eq $tokensTsvRow."Tokens Token")}
    $webhitsTsvRow = [FastSearch.Search]::Find($webhitsTsvFile, "Webhits Token", $tokensTsvRow."Tokens Token")

    # Check if this person/row has a click
    if($webhitsTsvRow -ne $null){

        # Get count of clicks for person/row
        $webhitsTsvRowCount = [FastSearch.Count]::Find($webhitsTsvFile, "Webhits Token", $tokensTsvRow."Tokens Token")

        # Check if there was only one webhits row for the given token
        if($webhitsTsvRowCount -eq 1){

            # Add webhits data
            $tokensTsvRow."Webhits When" = $webhitsTsvRow."Webhits When"
            $tokensTsvRow."Webhits When Converted" = ConvertFromUnixTimestamp ($webhitsTsvRow."Webhits When" / 1000)
            $tokensTsvRow."Webhits Data" = $webhitsTsvRow."Webhits Data"
            $tokensTsvRow."Webhits Visit" = ($webhitsTsvRow."Webhits Data".Split(" ",4)[2]).trim()
            $tokensTsvRow."Webhits IP" = ($webhitsTsvRow."Webhits Data".Split(" ")[-1]).trim()
        }

        # There are multiple webhits rows for the given token
        else{

            # Use slower Where-Object to get multiple clicks
            $webhitsTsvRow = $webhitsTsvFile | Where-Object {($_."Webhits Token" -eq $tokensTsvRow."Tokens Token")}

            # Add first webhits data
            $tokensTsvRow."Webhits When" = $webhitsTsvRow[0]."Webhits When"
            $tokensTsvRow."Webhits When Converted" = ConvertFromUnixTimestamp ($webhitsTsvRow[0]."Webhits When" / 1000)
            $tokensTsvRow."Webhits Data" = $webhitsTsvRow[0]."Webhits Data"
            $tokensTsvRow."Webhits Visit" = ($webhitsTsvRow[0]."Webhits Data".Split(" ",4)[2]).trim()
            $tokensTsvRow."Webhits IP" = ($webhitsTsvRow[0]."Webhits Data".Split(" ")[-1]).trim()

            # Loop through each webhit row for the given token
            forEach($webhitMatch in $webhitsTsvRow){

                # Add all webhits data
                $tokensTsvRow."Webhits When Multiple" += $webhitMatch."Webhits When" + " : "
                $tokensTsvRow."Webhits When Converted Multiple" += (ConvertFromUnixTimestamp ($webhitMatch."Webhits When" / 1000)).ToString("MM/dd/yyyy HH:mm:ss") + " : "
                $tokensTsvRow."Webhits Data Multiple" += $webhitMatch."Webhits Data" + " : "
                $tokensTsvRow."Webhits Visit Multiple" += ($webhitMatch."Webhits Data".Split(" ",4)[2]).trim() + " : "
                $tokensTsvRow."Webhits IP Multiple" += ($webhitMatch."Webhits Data".Split(" ")[-1]).trim() + " : "
            }

            # Remove the last " : "
            $tokensTsvRow."Webhits When Multiple" = $tokensTsvRow."Webhits When Multiple".Substring(0, $tokensTsvRow."Webhits When Multiple".Length-3)
            $tokensTsvRow."Webhits When Converted Multiple" = $tokensTsvRow."Webhits When Converted Multiple".Substring(0, $tokensTsvRow."Webhits When Converted Multiple".Length-3)
            $tokensTsvRow."Webhits Data Multiple" = $tokensTsvRow."Webhits Data Multiple".Substring(0, $tokensTsvRow."Webhits Data Multiple".Length-3)
            $tokensTsvRow."Webhits Visit Multiple" = $tokensTsvRow."Webhits Visit Multiple".Substring(0, $tokensTsvRow."Webhits Visit Multiple".Length-3)
            $tokensTsvRow."Webhits IP Multiple" = $tokensTsvRow."Webhits IP Multiple".Substring(0, $tokensTsvRow."Webhits IP Multiple".Length-3)
        }
    }

    # Check if the user agent should be looked up
    if($lookupUserAgentFromCustomCnaScript -eq $TRUE){

        # Lookup the events tsv row
        # $eventsTsvRow = $eventsTsvFile | Where-Object {($_."Events Token" -eq $tokensTsvRow."Tokens Token")}
        $eventsTsvRow = [FastSearch.Search]::Find($eventsTsvFile, "Events Token", $tokensTsvRow."Tokens Token")

        # Check if this person/row has a click
        if($eventsTsvRow -ne $null){

            # Get count of clicks for person/row
            $eventsTsvRowCount = [FastSearch.Count]::Find($eventsTsvFile, "Events Token", $tokensTsvRow."Tokens Token")

            # Check if there was only one events row for the given token
            if($eventsTsvRowCount -eq 1){

                # add events data
                $tokensTsvRow."Events When" = $eventsTsvRow."Events When"
                $tokensTsvRow."Events When Converted" = ConvertFromUnixTimestamp ($eventsTsvRow."Events When" / 1000)
                $tokensTsvRow."Events Type" = $eventsTsvRow."Events Type"
                $tokensTsvRow."Events IP" = $eventsTsvRow."Events IP"
                $tokensTsvRow."Events Method" = $eventsTsvRow."Events Method"
                $tokensTsvRow."Events Link" = $eventsTsvRow."Events Link"
                $tokensTsvRow."Events User Agent" = $eventsTsvRow."Events User Agent"
            }

            # There are multiple events rows for the given token
            else{

                # User slower Where-Object to get multiple clicks
                $eventsTsvRow = $eventsTsvFile | Where-Object {($_."Events Token" -eq $tokensTsvRow."Tokens Token")}

                # Add first events data
                $tokensTsvRow."Events When" = $eventsTsvRow[0]."Events When"
                $tokensTsvRow."Events When Converted" = ConvertFromUnixTimestamp ($eventsTsvRow[0]."Events When" / 1000)
                $tokensTsvRow."Events Type" = $eventsTsvRow[0]."Events Type"
                $tokensTsvRow."Events IP" = $eventsTsvRow[0]."Events IP"
                $tokensTsvRow."Events Method" = $eventsTsvRow[0]."Events Method"
                $tokensTsvRow."Events Link" = $eventsTsvRow[0]."Events Link"
                $tokensTsvRow."Events User Agent" = $eventsTsvRow[0]."Events User Agent"

                # Loop through each events row for the given token
                forEach($eventMatch in $eventsTsvRow){

                    # add all events data
                    $tokensTsvRow."Events When Multiple" += $eventMatch."Events When" + " : "
                    $tokensTsvRow."Events When Converted Multiple" += (ConvertFromUnixTimestamp ($eventMatch."Events When" / 1000)).ToString("MM/dd/yyyy HH:mm:ss") + " : "
                    $tokensTsvRow."Events Type Multiple" += $eventMatch."Events Type" + " : "
                    $tokensTsvRow."Events IP Multiple" += $eventMatch."Events IP" + " : "
                    $tokensTsvRow."Events Method Multiple" += $eventMatch."Events Method" + " : "
                    $tokensTsvRow."Events Link Multiple" += $eventMatch."Events Link" + " : "
                    $tokensTsvRow."Events User Agent Multiple" += $eventMatch."Events User Agent" + " : "
                }

                # Remove the last " : "
                $tokensTsvRow."Events When Multiple" = $tokensTsvRow."Events When Multiple".Substring(0, $tokensTsvRow."Events When Multiple".Length-3)
                $tokensTsvRow."Events When Converted Multiple" = $tokensTsvRow."Events When Converted Multiple".Substring(0, $tokensTsvRow."Events When Converted Multiple".Length-3)
                $tokensTsvRow."Events Type Multiple" = $tokensTsvRow."Events Type Multiple".Substring(0, $tokensTsvRow."Events Type Multiple".Length-3)
                $tokensTsvRow."Events IP Multiple" = $tokensTsvRow."Events IP Multiple".Substring(0, $tokensTsvRow."Events IP Multiple".Length-3)
                $tokensTsvRow."Events Method Multiple" = $tokensTsvRow."Events Method Multiple".Substring(0, $tokensTsvRow."Events Method Multiple".Length-3)
                $tokensTsvRow."Events Link Multiple" = $tokensTsvRow."Events Link Multiple".Substring(0, $tokensTsvRow."Events Link Multiple".Length-3)
                $tokensTsvRow."Events User Agent Multiple" = $tokensTsvRow."Events User Agent Multiple".Substring(0, $tokensTsvRow."Events User Agent Multiple".Length-3)
            }
        }
    }

    # Check if the employee's custom information should be looked up
    if($lookupAdditionalEmployeeDetailsUsingPhishMeFormat -eq $TRUE){

        # Lookup the employee's row from the email address
        # $phishMeInputCsvRow = $phishMeInputCsvFile | Where-Object {($_."PhishMeInput Email" -eq $tokensTsvRow."Tokens Email")}
        $phishMeInputCsvRow = [FastSearch.Search]::Find($phishMeInputCsvFile, "PhishMeInput Email", $tokensTsvRow."Tokens Email")

        # Add employee's data
        $tokensTsvRow."PhishMeInput Name" = $phishMeInputCsvRow."PhishMeInput Name"
        $tokensTsvRow."PhishMeInput Department" = $phishMeInputCsvRow."PhishMeInput Department"
        $tokensTsvRow."PhishMeInput Location" = $phishMeInputCsvRow."PhishMeInput Location"
    }

    # Break for testing
    if($limitResults -And $rowCount -gt 100){break}

    # Increment row count, get percent and show status
    $rowCount = $rowCount + 1
    $percents = [math]::round((($rowCount/($rows+1)) * 100), 0)
    Write-Progress -Activity:"Combining tsv files" -Status:"Transferred $rowCount of total $rows rows ($percents%)" -PercentComplete:$percents
}

# Set the filename for the combined tsv file
$phishingCampaignResultsCombinedTsvFilesFilename = $phishingReportsOutputFolder + "Phishing Campaign Report 1 - Combined Cobalt Strike TSV Files"

# Save the report
SaveReport $phishingCampaignResultsCombinedTsvFilesFilename $tokensTsvFile

# Create a file for the end of day reporting
$endOfDayCsvFile = $tokensTsvFile | Select-Object -Property "Tokens Email","PhishMeInput Name","PhishMeInput Department","PhishMeInput Location","Webhits When Converted" | Where-Object "Webhits When Converted" -ne ""

# Set the filename for the end of day file
$phishingCampaignResultsDailyFormatFilename = $phishingReportsOutputFolder + "Phishing Campaign Report 2 - End of Day - " + (Get-Date -f yyyy_MM_dd)

# Save the report
SaveReport $phishingCampaignResultsDailyFormatFilename $endOfDayCsvFile

# Create an array of objects
$arrayOfObjects = @()

# Set total row count
$rows = $tokensTsvFile.Count

# Set row count to zero
$rowCount = 0

# Loop through each row
forEach($row in $tokensTsvFile){

    # Create a temporary object for this plugin
    $tempObject = New-Object PSObject

    # Add members to the object
    $tempObject | Add-Member -Name "Scenario" -Value $scenario -MemberType NoteProperty
    $tempObject | Add-Member -Name "Email" -Value $row."Tokens Email" -MemberType NoteProperty
    $tempObject | Add-Member -Name "Recipient Name" -Value $row."PhishMeInput Name" -MemberType NoteProperty
    $tempObject | Add-Member -Name "Recipient Group" -Value $row."Campaigns Subject" -MemberType NoteProperty
    $tempObject | Add-Member -Name "Department" -Value $row."PhishMeInput Department" -MemberType NoteProperty
    $tempObject | Add-Member -Name "Location" -Value $row."PhishMeInput Location" -MemberType NoteProperty
    $tempObject | Add-Member -Name "Clicked Link?" -Value (&{If($row."Webhits When Converted" -ne "") {"Yes"} Else {"No"}}) -MemberType NoteProperty
    $tempObject | Add-Member -Name "Clicked Link Timestamp" -Value $row."Webhits When Converted" -MemberType NoteProperty
    $tempObject | Add-Member -Name "Reported Phish?" -Value $row."Webhits When Converted Multiple" -MemberType NoteProperty
    $tempObject | Add-Member -Name "New/Repeat Reporter" -Value $row."Webhits IP Multiple" -MemberType NoteProperty
    $tempObject | Add-Member -Name "Reported Phish Timestamp" -Value $row."Webhits Visit Multiple" -MemberType NoteProperty
    $tempObject | Add-Member -Name "Time to Report (in seconds)" -Value "" -MemberType NoteProperty
    $tempObject | Add-Member -Name "Remote IP" -Value $row."Webhits IP" -MemberType NoteProperty
    $tempObject | Add-Member -Name "GeoIP Country" -Value $row."Webhits Visit" -MemberType NoteProperty
    $tempObject | Add-Member -Name "GeoIP City" -Value $row."Applications Internal" -MemberType NoteProperty
    $tempObject | Add-Member -Name "GeoIP Organization" -Value $row."Applications Internal Multiple" -MemberType NoteProperty
    $tempObject | Add-Member -Name "Last DSN" -Value ($row."Sentemails Status Reason".Split(" ")[0]).trim() -MemberType NoteProperty
    $tempObject | Add-Member -Name "Last Email Status" -Value $row."Campaigns Template" -MemberType NoteProperty
    $tempObject | Add-Member -Name "Last Email Status Timestamp" -Value $row."Sentemails When Converted" -MemberType NoteProperty
    $tempObject | Add-Member -Name "Language" -Value $row."Events IP" -MemberType NoteProperty
    $tempObject | Add-Member -Name "Browser" -Value $row."Events Method" -MemberType NoteProperty
    $tempObject | Add-Member -Name "User-Agent" -Value $row."Events User Agent" -MemberType NoteProperty
    $tempObject | Add-Member -Name "Mobile?" -Value $row."Applications Application" -MemberType NoteProperty
    $tempObject | Add-Member -Name "Seconds Spent on Education Page" -Value $row."Applications Application Multiple" -MemberType NoteProperty

    # Update the array of objects
    $arrayOfObjects += $tempObject

    # Break if testing
    if($limitResults -And $rowCount -gt 100){break}

    # Increment row count, get percent and show status
    $rowCount = $rowCount + 1
    $percents = [math]::round((($rowCount/($rows+1)) * 100), 0)
    Write-Progress -Activity:"Converting to phishing campaign results to PhishMe format" -Status:"Converted $rowCount of total $rows rows ($percents%)" -PercentComplete:$percents
}

# Set the filename for the PhishMe output file
$phishingCampaignResultsPhishMeFormatFilename = $phishingReportsOutputFolder + "Phishing Campaign Report 3 - PhishMe Output Format"

# Save the report
SaveReport $phishingCampaignResultsPhishMeFormatFilename $arrayOfObjects

# Create an array of objects
$arrayOfObjects = @()

# Set total row count
$rows = $tokensTsvFile.Count

# Set row count to zero
$rowCount = 0

# Loop through each row
forEach($row in $tokensTsvFile){

    # Create a temporary object for this plugin
    $tempObject = New-Object PSObject

    # Add members to the object
    $tempObject | Add-Member -Name "Scenario" -Value $scenario -MemberType NoteProperty
    $tempObject | Add-Member -Name "Email" -Value $row."Tokens Email" -MemberType NoteProperty
    $tempObject | Add-Member -Name "Name" -Value $row."PhishMeInput Name" -MemberType NoteProperty
    $tempObject | Add-Member -Name "Manager's Name" -Value $row."PhishMeInput Department" -MemberType NoteProperty
    $tempObject | Add-Member -Name "Location" -Value $row."PhishMeInput Location" -MemberType NoteProperty
    $tempObject | Add-Member -Name "Clicked?" -Value (&{If($row."Webhits When Converted" -ne "") {"Yes"} Else {"No"}}) -MemberType NoteProperty
    $tempObject | Add-Member -Name "Clicked Timestamp" -Value $row."Webhits When Converted" -MemberType NoteProperty
    $tempObject | Add-Member -Name "Clicked Timestamp (Multiple)" -Value $row."Webhits When Converted Multiple" -MemberType NoteProperty
    $tempObject | Add-Member -Name "IP Address" -Value $row."Webhits IP" -MemberType NoteProperty
    $tempObject | Add-Member -Name "IP Address (Multiple)" -Value $row."Webhits IP Multiple" -MemberType NoteProperty
    $tempObject | Add-Member -Name "Visit" -Value $row."Webhits Visit" -MemberType NoteProperty
    $tempObject | Add-Member -Name "Visit (Multiple)" -Value $row."Webhits Visit Multiple" -MemberType NoteProperty
    $tempObject | Add-Member -Name "Application" -Value $row."Applications Application" -MemberType NoteProperty
    $tempObject | Add-Member -Name "Application (Multiple)" -Value $row."Applications Application Multiple" -MemberType NoteProperty
    $tempObject | Add-Member -Name "Internal IP Address" -Value $row."Applications Internal" -MemberType NoteProperty
    $tempObject | Add-Member -Name "Internal IP Address (Multiple)" -Value $row."Applications Internal Multiple" -MemberType NoteProperty
    $tempObject | Add-Member -Name "Campaign Subject" -Value $row."Campaigns Subject" -MemberType NoteProperty
    $tempObject | Add-Member -Name "Campaign Template" -Value $row."Campaigns Template" -MemberType NoteProperty
    $tempObject | Add-Member -Name "Status" -Value ($row."Sentemails Status Reason".Split(" ")[0]).trim() -MemberType NoteProperty
    $tempObject | Add-Member -Name "Status Timestamp" -Value $row."Sentemails When Converted" -MemberType NoteProperty
    $tempObject | Add-Member -Name "Method [Custom]" -Value $row."Events Method" -MemberType NoteProperty
    $tempObject | Add-Member -Name "User-Agent [Custom]" -Value $row."Events User Agent" -MemberType NoteProperty
    $tempObject | Add-Member -Name "IP Address [Custom]" -Value $row."Events IP" -MemberType NoteProperty
    $tempObject | Add-Member -Name "Method (Multiple) [Custom]" -Value $row."Events Method Multiple" -MemberType NoteProperty
    $tempObject | Add-Member -Name "User-Agent (Multiple) [Custom]" -Value $row."Events User Agent Multiple" -MemberType NoteProperty
    $tempObject | Add-Member -Name "IP Address (Multiple) [Custom]" -Value $row."Events IP Multiple" -MemberType NoteProperty

    # Update the array of objects
    $arrayOfObjects += $tempObject

    # Break if testing
    if($limitResults -And $rowCount -gt 100){break}

    # Increment row count, get percent and show status
    $rowCount = $rowCount + 1
    $percents = [math]::round((($rowCount/($rows+1)) * 100), 0)
    Write-Progress -Activity:"Converting to phishing campaign results to custom format" -Status:"Converted $rowCount of total $rows rows ($percents%)" -PercentComplete:$percents
}

# Set the filename for the custom phishing report file
$phishingCampaignResultsCustomFormatFilename = $phishingReportsOutputFolder + "Phishing Campaign Report 4 - Custom Format"

# Save the report
SaveReport $phishingCampaignResultsCustomFormatFilename $arrayOfObjects

# Write the end time
Write-Host (Get-Date -f yyyy_MM_dd-HH:mm:ss)"- End"