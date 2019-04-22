# Cobalt Strike Phishing Campaign Reporting (PhishReportCS)

PhishReportCS is a penetration testing and red teaming tool that automates the phishing campaign reporting process for Cobalt Strike phishing campaigns.

## Main Features

 - Automated phishing campaign reporting
   - A phishing report that combines the Cobalt Strike TSV files
   - A phishing report that highlights all phishing clicks
   - A phishing report that is converted into a reporting format similar to PhishMe but with additional data
   - A phishing report that customized to highlight important details from the phishing campaign

## Requirements

 - Cobalt Strike
 - [Custom Cobalt Strike Aggressor Script (PhishingProfiler.cna)](https://github.com/jamesm0rr1s/Cobalt-Strike-Aggressor-Scripts)
 
## Installation

Clone the GitHub repository
```
git clone https://github.com/jamesm0rr1s/Cobalt-Strike-Phishing-Campaign-Reporting /opt/jamesm0rr1s/Cobalt-Strike-Phishing-Campaign-Reporting
```

## Usage

 - Load the [Custom Cobalt Strike Aggressor Script (PhishingProfiler.cna)](https://github.com/jamesm0rr1s/Cobalt-Strike-Aggressor-Scripts)
 - Execute a phishing campaign with Cobalt Strike
 - Export the Cobalt Strike phishing data in TSV files
 - Update the directory names in the PowerShell script (Lines 14, 17, & 20)
 - Run the following PowerShell script:
```
CreatePhishingReportsFromCobaltStrikePhishingCampaign.ps1
```

## Example Screenshots

### Input Files

[Example of applications.tsv](Example%20Input%20-%20Cobalt%20Strike%20TSV%20Files/applications.tsv)  
![ExampleInput-applications.tsv](Example%20Screenshots%20-%20Input%20and%20Output%20Files/Input%20-%20applications.tsv.png?raw=true "ExampleInput-applications.tsv")

[Example of campaigns.tsv](Example%20Input%20-%20Cobalt%20Strike%20TSV%20Files/campaigns.tsv)  
![ExampleInput-campaigns.tsv](Example%20Screenshots%20-%20Input%20and%20Output%20Files/Input%20-%20campaigns.tsv.png?raw=true "ExampleInput-campaigns.tsv")

[Example of EmployeeDetails.csv](Example%20Input%20-%20Employee%20Details%20in%20PhishMe%20Input%20Format/Employee%20Details%20-%20PhishMe%20Input%20Format.csv)  
![ExampleInput-EmployeeDetails.csv](Example%20Screenshots%20-%20Input%20and%20Output%20Files/Input%20-%20Employee%20Details.csv.png?raw=true "ExampleInput-EmployeeDetails.csv")

[Example of events.tsv](Example%20Input%20-%20Cobalt%20Strike%20TSV%20Files/events.tsv)  
![ExampleInput-events.tsv](Example%20Screenshots%20-%20Input%20and%20Output%20Files/Input%20-%20events.tsv.png?raw=true "ExampleInput-events.tsv")

[Example of sentemails.tsv](Example%20Input%20-%20Cobalt%20Strike%20TSV%20Files/sentemails.tsv)  
![ExampleInput-sentemails.tsv](Example%20Screenshots%20-%20Input%20and%20Output%20Files/Input%20-%20sentemails.tsv.png?raw=true "ExampleInput-sentemails.tsv")

[Example of tokens.tsv](Example%20Input%20-%20Cobalt%20Strike%20TSV%20Files/tokens.tsv)  
![ExampleInput-tokens.tsv](Example%20Screenshots%20-%20Input%20and%20Output%20Files/Input%20-%20tokens.tsv.png?raw=true "ExampleInput-tokens.tsv")

[Example of webhits.tsv](Example%20Input%20-%20Cobalt%20Strike%20TSV%20Files/webhits.tsv)  
![ExampleInput-webhits.tsv](Example%20Screenshots%20-%20Input%20and%20Output%20Files/Input%20-%20webhits.tsv.png?raw=true "ExampleInput-webhits.tsv")

### Output Files

[Example of Phishing Campaign Report 1](Example%20Output%20-%20Phishing%20Reports/Phishing%20Campaign%20Report%201%20-%20Combined%20Cobalt%20Strike%20TSV%20Files.csv)

[Example of Phishing Campaign Report 2](Example%20Output%20-%20Phishing%20Reports/Phishing%20Campaign%20Report%202%20-%20End%20of%20Day%20-%20YYYY_MM_DD.csv)  
![ExampleOutput-PhishingReport2.xlsx](Example%20Screenshots%20-%20Input%20and%20Output%20Files/Output%20-%20Phishing%20Report%202.png?raw=true "ExampleOutput-PhishingReport2.xlsx")

[Example of Phishing Campaign Report 3](Example%20Output%20-%20Phishing%20Reports/Phishing%20Campaign%20Report%203%20-%20PhishMe%20Output%20Format.csv)  
![ExampleOutput-PhishingReport3.xlsx](Example%20Screenshots%20-%20Input%20and%20Output%20Files/Output%20-%20Phishing%20Report%203.png?raw=true "ExampleOutput-PhishingReport3.xlsx")

[Example of Phishing Campaign Report 4](Example%20Output%20-%20Phishing%20Reports/Phishing%20Campaign%20Report%204%20-%20Custom%20Format.csv)  
![ExampleOutput-PhishingReport4.xlsx](Example%20Screenshots%20-%20Input%20and%20Output%20Files/Output%20-%20Phishing%20Report%204.png?raw=true "ExampleOutput-PhishingReport4.xlsx")
