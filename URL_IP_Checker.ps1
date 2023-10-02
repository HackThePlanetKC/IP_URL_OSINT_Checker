#Set some globals
$vtAPI = "VirusTotalAPI"
$abuseAPI = "abuseIPDBAPI"
$shodanAPI = "shodanAPI"

$outputXLSX = "C:\Path\To\output $(Get-Date -Format yyyy-MM-dd_HH-mm-ss).xlsx"
$outputCSV = "C:\Path\To\output $(Get-Date -Format yyyy-MM-dd_HH-mm-ss).csv"
$otherSearches = "C:\Path\To\Searches $(Get-Date -Format yyyy-MM-dd_HH-mm-ss).csv"

#Asks the user for the file path to the input txt file, relative path can be used if running from the same directory
$source_file = Read-Host "What's the path to the input file?"
$sources = Get-Content $source_file

# Define a regex pattern for IP address
$ipPattern = '^(25[0-5]|2[0-4][0-9]|[01]?[0-9][0-9]?)\.(25[0-5]|2[0-4][0-9]|[01]?[0-9][0-9]?)\.(25[0-5]|2[0-4][0-9]|[01]?[0-9][0-9]?)\.(25[0-5]|2[0-4][0-9]|[01]?[0-9][0-9]?)$'

#Header information for the GET requests to the API
$headers=@{}
$headers.Add("accept", "application/json")
$headers.Add("x-apikey", "$($vtAPI)")

#Create the headers for the csv
$header_information = "SOURCE,IP,Harmless,Malicious,Suspicious,Undetected,Timeout,Engine Name,Category,Result,Method,Engine Name,Category,Result,Method,Engine Name,Category,Result,Method,Engine Name,Category,Result,Method,Engine Name,Category,Result,Method,Engine Name,Category,Result,Method,Engine Name,Category,Result,Method,Engine Name,Category,Result,Method,Engine Name,Category,Result,Method,Engine Name,Category,Result,Method,Engine Name,Category,Result,Method,Engine Name,Category,Result,Method,Engine Name,Category,Result,Method,Engine Name,Category,Result,Method,Engine Name,Category,Result,Method,Engine Name,Category,Result,Method,Engine Name,Category,Result,Method,Engine Name,Category,Result,Method,Engine Name,Category,Result,Method,Engine Name,Category,Result,Method,Engine Name,Category,Result,Method,Engine Name,Category,Result,Method,Engine Name,Category,Result,Method,Engine Name,Category,Result,Method,Engine Name,Category,Result,Method,Engine Name,Category,Result,Method,Engine Name,Category,Result,Method,Engine Name,Category,Result,Method,Engine Name,Category,Result,Method,Engine Name,Category,Result,Method,Engine Name,Category,Result,Method,Engine Name,Category,Result,Method,Engine Name,Category,Result,Method,Engine Name,Category,Result,Method,Engine Name,Category,Result,Method,Engine Name,Category,Result,Method,Engine Name,Category,Result,Method,Engine Name,Category,Result,Method,Engine Name,Category,Result,Method,Engine Name,Category,Result,Method,Engine Name,Category,Result,Method,Engine Name,Category,Result,Method,Engine Name,Category,Result,Method,Engine Name,Category,Result,Method,Engine Name,Category,Result,Method,Engine Name,Category,Result,Method,Engine Name,Category,Result,Method,Engine Name,Category,Result,Method,Engine Name,Category,Result,Method,Engine Name,Category,Result,Method,Engine Name,Category,Result,Method,Engine Name,Category,Result,Method,Engine Name,Category,Result,Method,Engine Name,Category,Result,Method,Engine Name,Category,Result,Method,Engine Name,Category,Result,Method,Engine Name,Category,Result,Method,Engine Name,Category,Result,Method,Engine Name,Category,Result,Method,Engine Name,Category,Result,Method,Engine Name,Category,Result,Method,Engine Name,Category,Result,Method,Engine Name,Category,Result,Method,Engine Name,Category,Result,Method,Engine Name,Category,Result,Method,Engine Name,Category,Result,Method,Engine Name,Category,Result,Method,Engine Name,Category,Result,Method,Engine Name,Category,Result,Method,Engine Name,Category,Result,Method,Engine Name,Category,Result,Method,Engine Name,Category,Result,Method,Engine Name,Category,Result,Method,Engine Name,Category,Result,Method,Engine Name,Category,Result,Method,Engine Name,Category,Result,Method,Engine Name,Category,Result,Method,Engine Name,Category,Result,Method,Engine Name,Category,Result,Method,Engine Name,Category,Result,Method,Engine Name,Category,Result,Method,Engine Name,Category,Result,Method,Engine Name,Category,Result,Method,Engine Name,Category,Result,Method,Engine Name,Category,Result,Method,Engine Name,Category,Result,Method,Engine Name,Category,Result,Method,Engine Name,Category,Result,Method"

#Create the xlsx document and add headers
#Create new object
$xlsx_workbook = New-Object -ComObject Excel.Application

#Set visiblity to False
$xlsx_workbook.Visible = $false

#Create new workbook
$workbook = $xlsx_workbook.Workbooks.add()

#Get first worksheet
$worksheet = $workbook.Worksheets.Item(1)
$worksheet.name = "VirusTotal"

#Set up headers from the csv in a way that will work for the xlsx
$header_data = $header_information -split ","

#Sets the headers in the xlsx document
$column = 1
foreach ($head in $header_data){
    $worksheet.Cells.Item(1, $column) = $head
    $column++
}

#Starts running a loop for each line in source .txt file, $sources
foreach($source in $sources){

    #Check if $source is an IP or URL
    if ($source -match $ipPattern){
        $response = Invoke-WebRequest -Uri "https://www.virustotal.com/api/v3/ip_addresses/$($source)" -UseBasicParsing -Method GET -Headers $headers
    } else{
        $response = Invoke-WebRequest -Uri "https://www.virustotal.com/api/v3/urls/$($source)" -UseBasicParsing -Method GET -Headers $headers
    }
    
    #Start adding things to the xlsx document
    #Find the last row used in the worksheet
    $lastRow = $worksheet.Cells.Range("A1048576").End(-4162).Row

    #Increment the row number by one to get the next empty row
    $nextRow = $lastRow +1

    $test = $response | ConvertFrom-Json
    $check = $test.error.code

    #Check for a quota exceeded error
    if ($check -eq 429){
        $worksheet.Cells.Item($nextRow, 1) = "QUOTA EXCEEDED"
        #Highlight this row magenta to stand out
        $worksheet.Cells.Item($nextRow, 1).EntireRow.Interior.ColorIndex = 7 #magenta
    } else{
        $parsed = $response | ConvertFrom-Json
        $individual_IP = $parsed.data.id

        #Get the last_analysis_stats results
        $stats = $parsed.data.attributes.last_analysis_stats
        $last_analysis_stats = @()
    
        #Run through and get the stats fields first
        foreach($stat in $stats.PSObject.Properties){
            $last_analysis_stats += $stat.Value
        }
    
        #Get the last_analysis_results fields
        $results = $parsed.data.attributes.last_analysis_results

        $last_analysis_results = @()

        foreach($result in $results.PSObject.Properties){
            #Loop through the results and append the results to the array
            $last_analysis_results += $result.Value.engine_name
            $last_analysis_results += $result.Value.category
            $last_analysis_results += $result.Value.result
            $last_analysis_results += $result.Value.method
        }
    
        #Create the csv line for the results
        $engine_results = $last_analysis_results -join ","
        $csv_stats = $last_analysis_stats -join ","
    
        #Join everything together
        $csv_data = $source + "," + $individual_IP + "," + $csv_stats + "," + $engine_results

        #Add the results onto a new line in the csv that we built earlier
        Add-Content -Path $outputCSV -Value $csv_data
    
        #Split csv_data into an array
        $csvArray = $csv_data -split ","

        #loop through the array and assign each value to a cell in the worksheet
        $column = 1
        foreach ($value in $csvArray){
            $worksheet.Cells.Item($nextRow, $column) = $value
            $column++
       }
    }
}

#auto-size and add filters to the columns
$range = $worksheet.UsedRange
$range.AutoFilter()
$range.Columns.AutoFit()

#Set up list for sending to other scans
$red = @()

#Start coloring the rows based on malicious indicators
$rows = $worksheet.UsedRange.Rows.Count
for ($row = 2; $row -le $rows; $row ++){
    $rowDValue = $worksheet.Cells.Item($row, 4).Value2
    if ($rowDValue -gt 30){
        $worksheet.Cells.Item($row, 1).EntireRow.Interior.ColorIndex = 3 #red
        Add-Content -Path $otherSearches -Value $worksheet.Cells.Item($row, 2)
    } elseif ($rowDValue -gt 15){
        $worksheet.Cells.Item($row, 1).EntireRow.Interior.ColorIndex = 6 #yellow
    } elseif ($rowDValue -lt 5){
        $worksheet.Cells.Item($row, 1).EntireRow.Interior.ColorIndex = 4 #green
    }
}

#Sort xlsx to bring red and yellow to the top
for ($row = 2; $row -le $rows; $row++) { 
    #Get the color index of column A in the current row 
    $color = $worksheet.Cells.Item($row, 1).Interior.ColorIndex

    switch ($color) {
        3 {$sortOrder = 1} #red
        6 {$sortOrder = 2} #yellow
        default {$sortOrder = 3} #Other colors or none
        4 {$sortOrder = 4} #green
        7 {$sortOrder = 5} #magenta
    }
    #Set the value of column MV in the current row to the sort order
    $worksheet.Cells.Item($row, 360) = $sortOrder
}

#Read how many rows there are total, sorts them based on the value in column MV, then clears MV
$range = $worksheet.Range('A2','MV' + $rows)
$range.Sort($worksheet.Range('MV2'),1)
$worksheet.Range(‘MV2’, ‘MV’ + $rows).Clear()

#Save the workboox as XLSX
$workbook.SaveAs($outputXLSX, 51)

#Transistion period
Write-Host("`n`nVirusTotal document finished, beginning other scans")

#Run the red results through AbuseIPDB
Write-Host("`nScanning through AbuseIPDB")

$abuseHeaders =- @{
    "Key" = $vtAPI
    "Accept" = "application/json"
}

#Get second worksheet
$abuseWorksheet = $workbook.Worksheets.Item(2)
$abuseWorksheet.name = "abuseIPDB"

$abuseHeader_information = "IP Address"+","+"Country Name" +","+"Country Code"+","+"Usage Type"+","+"ISP"+","+"Domain"
#Set up headers from the csv in a way that will work for the xlsx
$abuseHeader_data = $abuseHeader_information -split ","

#Sets the headers in the abuseIPDB worksheet
$column = 1
foreach ($head in $abuseHeader_data){
    $abuseWorksheet.Cells.Item(1, $column) = $head
    $column++
}

foreach ($ip in $otherSearches){
    $query = @{
        "ipAddress" = $ip
        "maxAgeInDays" = 90
        "verbose" = $true
    }
    $response = Invoke-WebRequest -Uri "https://api.abuseipdb.com/api/v2/check" -Method Get -Headers $abuseHeaders -Body $query
    
    #Start adding things to the xlsx document
    #Find the last row used in the worksheet
    $lastRow = $abuseWorksheet.Cells.Range("A1048576").End(-4162).Row

    #Increment the row number by one to get the next empty row
    $nextRow = $lastRow +1

    $abuseResults = $response | ConvertFrom-Json

    $abuseData = $abuseResults.data.ipAddress +","+ $abuseResults.data.countryName + "," + $abuseResults.data.countryCode + "," + $abuseResults.data.usageType + "," + $abuseResults.data.isp + "," + $abuseResults.data.domain
    
    #Split abuseData into an array
    $abuseArray = $abuseData -split ","
    
    #loop through the array and assign each value to a cell in the worksheet
    $column = 1
    foreach ($value in $abuseArray){
            $abuseWorksheet.Cells.Item($nextRow, $column) = $value
            $column++
   }
Write-Host ($response)
}

#auto-size and add filters to the columns
$abuseRange = $abuseWorksheet.UsedRange
$abuseRange.AutoFilter()
$abuseRange.Columns.AutoFit()

#CrowdStrike stuff
$client_id = "Your CS Clinet ID"
$client_secret = "Your CS Client Secoret"

# Use the US-2 region endpoint
$endpoint = "https://api.us-2.crowdstrike.com/intel/queries/indicators/v1"

# Use the US-2 region token endpoint
$token_endpoint = "https://api.us-2.crowdstrike.com/oauth2/token"
$grant_type = "client_credentials"

# Invoke the token request and get the response for CS
$token_response = Invoke-RestMethod -Uri $token_endpoint -Method POST -Body @{
    "client_id" = $client_id
    "client_secret" = $client_secret
    "grant_type" = $grant_type
}

# Extract the access token from the response
$access_token = $token_response.access_token

# Define the authorization header with the access token
$header = @{
    "Authorization" = "Bearer $access_token"
}

#Get CrowdStrike worksheet
$crowdStrikeWorksheet = $workbook.Worksheets.Item(3)
$crowdStrikeWorksheet.name = "CrowdStrike"
$crowdStrikeColumn = 1
$lastCrowdStrikeRow = $crowdStrikeWorksheet.Cells.Range("A1048576").End(-4162).Row

#Increment the row number by one to get the next empty row
$nextCrowdStrikeRow = $lastCrowdStrikeRow +1

#Run the red results through CS now
foreach ($ip in $otherSearches){

    $params = @{
        type= "ip_address.$($ip)"
        "limit" = 10 # You can change this to adjust the number of results
    }

    # Invoke the web request and get the response
    $response = Invoke-WebRequest -Uri $endpoint -Method Get -Body $params -Headers $header

    # Convert the response to JSON object
    $json = $response.Content | ConvertFrom-Json

    # Check if there are any results
    if ($json.meta.pagination.total -gt 0) {
        # Loop through the results and print the actor names and descriptions
        foreach ($result in $json.resources) {
            Write-Host "Actor name: $($result.actor)"
            Write-Host "Actor description: $($result.actor_description)"
            Write-Host ""
            $crowdStrikeWorksheet.Cells.Item($nextRow, 1) = $result
            $nextCrowdStrikeRow = $lastCrowdStrikeRow +1
        }
    }
    else {
        # No results found
        Write-Host "No actors associated with the IP address $ip"
    }
}

#Finishing Touches
#auto-size and add filters to the columns
$crowdStrikeRange = $crowdStrikeWorksheet.UsedRange
$crowdStrikeRange.AutoFilter()
$crowdStrikeRange.Columns.AutoFit()

#Close and quit excel
$workbook.Close()
$xlsx_workbook.Quit()
