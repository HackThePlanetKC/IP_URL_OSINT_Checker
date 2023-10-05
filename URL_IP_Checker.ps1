#Set some globals
$vtAPI = "vtAPI"
$abuseAPI = "abuseAPI"
$shodanAPI = "shodanAPI"

#CrowdStrike stuff
$client_id = "client_id"
$client_secret = "client_secret"

#Actual report
$outputXLSX = "C:\Path\to\Report $(Get-Date -Format yyyy-MM-dd_HH-mm-ss).xlsx"

#placeholders, just used to store the data before building the actual report
$outputCSV = "C:\Path\to\output $(Get-Date -Format yyyy-MM-dd_HH-mm-ss).csv"
$otherSearches = "C:\Path\to\Searches $(Get-Date -Format yyyy-MM-dd_HH-mm-ss).csv"

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
    if ($check -eq "QuotaExceededError"){
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
    if ($rowDValue -gt 9){
        $worksheet.Cells.Item($row, 1).EntireRow.Interior.ColorIndex = 3 #red
        Add-Content -Path $otherSearches -Value $worksheet.Cells.Item($row, 2).Value2
    } elseif ($rowDValue -gt 5){
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


#Transistion period
Write-Host("`n`nVirusTotal document finished, beginning other scans")

#Run the red results through AbuseIPDB
Write-Host("`nScanning through AbuseIPDB")

$abuseHeaders =- @{
    "Key" = $vtAPI
    "Accept" = "application/json"
}

#Get second worksheet

$lastSheet = $workbook.WorkSheets($workbook.WorkSheets.Count)
$abuseWorksheet = $workbook.Worksheets.Add($lastSheet)
$abuseWorksheet.name = "abuseIPDB"
$lastSheet.Move($abuseWorksheet)

$abuseHeader_information = "IP Address"+","+"Country Name" +","+"Country Code"+","+"Usage Type"+","+"ISP"+","+"Domain"
#Set up headers from the csv in a way that will work for the xlsx
$abuseHeader_data = $abuseHeader_information -split ","

#Sets the headers in the abuseIPDB worksheet
$column = 1
foreach ($head in $abuseHeader_data){
    $abuseWorksheet.Cells.Item(1, $column) = $head
    $column++
}

$redSearches = Get-Content $otherSearches

foreach ($ip in $redSearches){
    $query = @{
        "ipAddress" = $ip
        "maxAgeInDays" = 90
        "verbose" = $true
    }
    $response = Invoke-WebRequest -Uri "https://api.abuseipdb.com/api/v2/check" -Method Get -Headers $abuseHeaders -Body $query
    
    #Start adding things to the xlsx document
    #Find the last row used in the worksheet
    $lastAbuseRow = $abuseWorksheet.Cells.Range("A1048576").End(-4162).Row

    #Increment the row number by one to get the next empty row
    $nextAbuseRow = $lastAbuseRow +1

    $abuseResults = $response | ConvertFrom-Json

    $abuseData = $abuseResults.data.ipAddress +","+ $abuseResults.data.countryName + "," + $abuseResults.data.countryCode + "," + $abuseResults.data.usageType + "," + $abuseResults.data.isp + "," + $abuseResults.data.domain
    
    #Split abuseData into an array
    $abuseArray = $abuseData -split ","
    
    #loop through the array and assign each value to a cell in the worksheet
    $column = 1
    foreach ($value in $abuseArray){
            $abuseWorksheet.Cells.Item($nextAbuseRow, $column) = $value
            $column++
   }
#Write-Host ($response)
}

#auto-size and add filters to the columns
$abuseRange = $abuseWorksheet.UsedRange
$abuseRange.AutoFilter()
$abuseRange.Columns.AutoFit()

Write-Host("`nScanning through CrowdStrike")

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

$lastSheet = $workbook.WorkSheets($workbook.WorkSheets.Count)
$crowdStrikeWorksheet = $workbook.Worksheets.Add($lastSheet)
$crowdStrikeWorksheet.name = "CrowdStrike"
$lastSheet.Move($crowdStrikeWorksheet)

#Sets up Crowdstrike Sheet Headers
$crowdStrikeHeader_information = "IP"+","+"Actor Name"+","+"Actor Description"
#Set up headers from the csv in a way that will work for the xlsx
$crowdStrikeHeader_data = $crowdStrikeHeader_information -split ","

#Sets the headers in the CrowdStrike worksheet
$column = 1
foreach ($head in $crowdStrikeHeader_data){
    $crowdStrikeWorksheet.Cells.Item(1, $column) = $head
    $column++
}

#Run the red results through CS now
foreach ($ip in $redSearches){

    #Find the last row used in the worksheet
    $lastCrowdStrikeRow = $crowdStrikeWorksheet.Cells.Range("A1048576").End(-4162).Row

    #Increment the row number by one to get the next empty row
    $nextCrowdStrikeRow = $lastCrowdStrikeRow +1

    $params = @{
        type= "ip_address.$($ip)"
        #"limit" = 10 # You can change this to adjust the number of results
    }

    # Invoke the web request and get the response
    $response = Invoke-WebRequest -Uri $endpoint -Method Get -Body $params -Headers $header

    # Convert the response to JSON object
    $json = $response.Content | ConvertFrom-Json

    # Check if there are any results
    if ($json.meta.pagination.total -gt 0) {
        # Loop through the results and print the actor names and descriptions
        foreach ($result in $json.resources) {

            #Find the last row used in the worksheet
            $lastCrowdStrikeRow = $crowdStrikeWorksheet.Cells.Range("A1048576").End(-4162).Row

            #Increment the row number by one to get the next empty row
            $nextCrowdStrikeRow = $lastCrowdStrikeRow +1

            Write-Host "Actor name: $($result.actor)"
            Write-Host "Actor description: $($result.actor_description)"
            Write-Host ""
            $crowdStrikeData = $result.actor +","+ $result.actor_description

            #Split abuseData into an array
            $crowdStrikeArray = $crowdStrikeData -split ","

            #Write the crowdStrike data into each cell on the row
            $crowdStrikeWorksheet.Cells.Item($nextCrowdStrikeRow, 1) = $ip
            $column = 2
            foreach ($value in $abuseArray){
                $crowdStrikeWorksheet.Cells.Item($nextCrowdStrikeRow, $column) = $value
                $column++
            }
        }
    }
    else {
        # No results found
        Write-Host "No actors associated with the IP address $ip"
        $crowdStrikeWorksheet.Cells.Item($nextCrowdStrikeRow, 1) = $ip
        $crowdStrikeWorksheet.Cells.Item($nextCrowdStrikeRow, 2) = "No actors associated with this IP address"
        $nextCrowdStrikeRow = $lastCrowdStrikeRow +1
    }
}

#Finishing Touches
#auto-size and add filters to the columns
$crowdStrikeRange = $crowdStrikeWorksheet.UsedRange
$crowdStrikeRange.AutoFilter()
$crowdStrikeRange.Columns.AutoFit()

Write-Host("`nScanning through Shodan")

#Start Shodan Stuff
$shodanHeaders=@{}
$shodanHeaders.Add("key", "$($shodanAPI)")

$shodanURL = "https://api.shodan.io/shodan/host/"

#Get CrowdStrike worksheet

$lastSheet = $workbook.WorkSheets($workbook.WorkSheets.Count)
$shodanWorksheet = $workbook.Worksheets.Add($lastSheet)
$shodanWorksheet.name = "Shodan"
$lastSheet.Move($shodanWorksheet)

#Sets up Crowdstrike Sheet Headers
$shodanHeader_information = "IP"+","+"Open Ports"
#Set up headers from the csv in a way that will work for the xlsx
$shodanHeader_data = $crowdStrikeHeader_information -split ","

#Sets the headers in the Shodan worksheet
$column = 1
foreach ($head in $shodanHeader_data){
    $shodanWorksheet.Cells.Item(1, $column) = $head
    $column++
}

#Start Shodan Scanning
foreach ($ip in $redSearches){
    #Find the last row used in the worksheet
    $lastShodanRow = $shodanWorksheet.Cells.Range("A1048576").End(-4162).Row

    #Increment the row number by one to get the next empty row
    $nextShodanRow = $lastShodanRow +1

    $response = Invoke-WebRequest -Uri "$($shodanURL)$($ip)" -UseBasicParsing -Method GET -Headers $shodanHeaders
    $shodanParsed = $response | ConvertFrom-Json
    $shodanWorksheet.Cells.Item($nextShodanRow, 1) = $shodanParsed.ip_str
    $shodanWorksheet.Cells.Item($nextShodanRow, 2) = $shodanParsed.ports
}

#auto-size and add filters to the columns
$shodanRange = $shodanWorksheet.UsedRange
$shodanRange.AutoFilter()
$shodanRange.Columns.AutoFit()

#Close and quit excel
#Save the workboox as XLSX
$workbook.SaveAs($outputXLSX, 51)
$workbook.Close()
$xlsx_workbook.Quit()
