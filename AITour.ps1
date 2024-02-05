# Remove-Module PowerShellAI
Import-Module PowerShellAI
set-location '.\chatGPT\AI Tour'

<#
Demo 1
#>

$AzureOpenAICred = Import-Clixml .\AzureOpenAIchatGPTAPIKey.xml
$AzureOpenAICred

Set-OpenAIKey -Key $AzureOpenAICred.Password
$AzureOpenAIAPIKey = [System.Runtime.InteropServices.marshal]::PtrToStringAuto([System.Runtime.InteropServices.marshal]::SecureStringToBSTR($AzureOpenAICred.Password))
Set-AzureOpenAI -Endpoint $AzureOpenAICred.UserName -ApiKey $AzureOpenAIAPIKey -DeploymentName "GPT35T" -ApiVersion "2023-07-01-preview" 

Get-ChatAPIProvider
Set-ChatAPIProvider AzureOpenAI

<#
Demo 2
#>

$response = (new-chat "what are the states and terriortories in Australia as json")
$response 
# $AUStatesTerriortories = ($response.split('```json')[1]).split('```')[0] | Convertfrom-json
$AUStatesTerriortories = $response | Convertfrom-json
$AUStatesTerriortories | Format-List
$AUStatesTerriortories.states
$AUStatesTerriortories.territories

foreach ($state in $AUStatesTerriortories.states) { chat "what is the population of $($state.name)" }
foreach ($territory in $AUStatesTerriortories.territories) { chat "what is the population of $($territory.name)" }

<#
Demo 3
#>

copilot "write a PowerShell function that takes a date object as input and calculates the number of days until that date from today"

function CalculateDaysUntilDate {
    param (
        [Parameter(Mandatory = $true)]
        [DateTime]$TargetDate
    )
    $today = Get-Date
    $daysUntil = ($TargetDate - $today).Days
    return $daysUntil
}
$targetDate = Get-Date -Year 2024 -Month 02 -Day 29
CalculateDaysUntilDate -TargetDate $targetDate

<# DEMO 4#>

$geo = Invoke-RestMethod http://ipinfo.io/json 
$weatherAPIKey = Import-Clixml .\weatherAPIKey.xml

# $weatherResult = Invoke-RestMethod -uri "https://api.openweathermap.org/data/2.5/weather?lat=$($geo.latitude)&lon=$($geo.longitude)&appid=$($weatherAPIKey)&units=metric" 
$weatherResult = Invoke-RestMethod -uri "https://api.openweathermap.org/data/2.5/weather?q=$($geo.city)&appid=$($weatherAPIKey)&units=metric" 
$landmark = new-chat "what is the most iconic landmark with a brief description in $($geo.city)"  
$landmarkSummary = chat "summarize $($landmark) into 30 words detailing only its key architectual features and not its history"
function Get-AzureOpenAIDalleArtistImage {
    [CmdletBinding()]
    Param
    (
        [Parameter(Mandatory)]
        [string]$prompt,
        [ValidateSet('256', '512', '1024')]
        $Size = 1024,
        $Images = 1,
        [Parameter(Mandatory)]
        [string]$artist,
        [Parameter(Mandatory)]
        [string]$city,
        [string]$apiVersion = "2023-06-01-preview"
    )

    # Azure OpenAI metadata variables
    $openai = @{
        api_key  = [System.Runtime.InteropServices.marshal]::PtrToStringAuto([System.Runtime.InteropServices.marshal]::SecureStringToBSTR($AzureOpenAICred.Password))
        api_base = $AzureOpenAICred.UserName
    }

    # Header for authentication
    $headers = [ordered]@{
        'api-key' = $openai.api_key
    }

    $targetSize = switch ($Size) {
        256 { '256x256' }
        512 { '512x512' }
        1024 { '1024x1024' }     
    }

    # Adjust these values to fine-tune completions
    $body = [ordered]@{
        prompt = $prompt
        size   = $targetSize
        n      = $Images
    } | ConvertTo-Json

    # Call the API to generate the image and retrieve the response
    $url = "$($openai.api_base)/openai/images/generations:submit?api-version=$($apiVersion)"

    $submission = Invoke-RestMethod -Uri $url -Headers $headers -Body $body -Method Post -ContentType 'application/json' -ResponseHeadersVariable submissionHeaders
    $operation_location = $submissionHeaders['operation-location'][0]

    $status = ''
    while ($status -ne 'succeeded') {
        Start-Sleep -Seconds 1
        $response = Invoke-RestMethod -Uri $operation_location -Headers $headers
        if ($response.status -eq 'failed') {
            Write-Error "Image Generation Failed with: $($response.error.code) and message: $($response.error.message)"
            exit 
        }
        $status = $response.status
    }

    # Retrieve the generated image
    $generatedImages = @()

    # Set the directory for the stored image
    $image_dir = Join-Path -Path $pwd -ChildPath 'images'

    # If the directory doesn't exist, create it
    if (-not(Resolve-Path $image_dir -ErrorAction Ignore)) {
        New-Item -Path $image_dir -ItemType Directory
    }
    
    $i = 1
    foreach ($generatedImage in $response.result.data.url) {
        $image_url = $generatedImage
        # Initialize the image path (note the filetype should be png)
        $ts = (get-date -Uformat %T).ToString().Replace(":", "-")
        $image_path = Join-Path -Path $image_dir -ChildPath "$($city)-$($artist)-$($ts)-$($i).png"
        Invoke-WebRequest -Uri $image_url -OutFile $image_path  # download the image
        $generatedImages += $image_path
        $i = $i + 1
    }
    return $generatedImages
}

Get-AzureOpenAIDalleArtistImage -Prompt "a painting of $($landmarkSummary) on a $($weatherResult.weather.description) day in the style of Rembrant" -artist "Rembrant" -city $geo.city -Verbose

$city = "New York"
$landmark = new-chat "what is the most iconic landmark in $($City)"  
$landmarkSummary = chat "summarize $($landmark) into 20 words detailing only its key architectual features and not its history"  
$weatherResult = Invoke-RestMethod -uri "https://api.openweathermap.org/data/2.5/weather?q=$($City)&appid=$($weatherAPIKey)&units=metric" 

Get-AzureOpenAIDalleArtistImage -Prompt "a painting of $($landmarkSummary) on a $($weatherResult.weather.description) day in the style of Rembrant" -artist "Van Gogh" -city $city  -Verbose

$city = "Tokyo"
$landmark = new-chat "what is the most iconic landmark in $($City)"  
$landmarkSummary = chat "summarize $($landmark) into 20 words detailing only its key architectual features and not its history"  
$weatherResult = Invoke-RestMethod -uri "https://api.openweathermap.org/data/2.5/weather?q=$($City)&appid=$($weatherAPIKey)&units=metric" 

Get-AzureOpenAIDalleArtistImage -Prompt "a painting of $($landmarkSummary) on a $($weatherResult.weather.description) day in the style of Rembrant" -artist "Monet" -city $city  -Verbose

$city = "Berlin"
$landmark = new-chat "what is the most iconic landmark in $($City)" -Verbose  
$landmarkSummary = chat "summarize $($landmark) into 20 words detailing only its key architectual features and not its history" -Verbose -max_tokens 4000
$weatherResult = Invoke-RestMethod -uri "https://api.openweathermap.org/data/2.5/weather?q=$($City)&appid=$($weatherAPIKey)&units=metric" 

Get-AzureOpenAIDalleArtistImage -Prompt "a painting of $($landmarkSummary) on a $($weatherResult.weather.description) day in the style of Rembrant" -artist "Picasso" -city $city  -Verbose


<# DEMO 5 #>

$artists = @("Rembrant", "Van Gogh", "Picasso", "Monet", "Gio Xi") 
$artist = $artists | Get-Random
$artist

$cities = @("Sydney", "Paris", "London", "New York", "Tokyo")
$city = $cities | Get-Random
$city

$landmark = new-chat "what is the most iconic landmark in $($City)"  
$landmarkSummary = chat "summarize $($landmark) into 20 words detailing only its key architectual features and not its history"  
$weatherResult = Invoke-RestMethod -uri "https://api.openweathermap.org/data/2.5/weather?q=$($City)&appid=$($weatherAPIKey)&units=metric" 

Get-AzureOpenAIDalleArtistImage -prompt "a painting of $($landmarkSummary) on a $($weatherResult.weather.description) day in the style of $($artist)" -city $city -Images 2 -Size 1024 -artist $artist


<# DEMO 5b #>

1..5 | ForEach-Object {
    $city = $cities | Get-Random
    $artist = $artists | Get-Random
    $landmark = new-chat "what is the most iconic landmark in $($city)"  
    $landmarkSummary = chat "summarize $($landmark) into 20 words detailing only its key architectual features and not its history"  
    $weatherResult = Invoke-RestMethod -uri "https://api.openweathermap.org/data/2.5/weather?q=$($city)&appid=$($weatherAPIKey)&units=metric" 

    Get-AzureOpenAIDalleArtistImage -prompt "a painting of $($landmarkSummary) on a $($weatherResult.weather.description) day in the style of $($artist)" -Size 1024 -artist $artist -city $city
}


<# DEMO 5c #>

Get-AzureOpenAIDalleImage -prompt "A Microsoft conference event held in a large convention center with 2000 attendees where a presenter is showing how to programmatically use AI with PowerShell to generate city landscapes in the styles of famous artists. " -Images 2


<# DEMO 6#>

Import-Module ImportExcel 

$australia = new-chat "a list of australian states and terrortories, population as json" | ConvertFrom-Json
$australia.states | Export-Excel 

$poly = new-chat "a list of countries in polynesia, population as json" | ConvertFrom-Json
$poly | Export-Excel 

$beers = new-chat "a list of the top 5 beers in Australia, style as json" | ConvertFrom-Json
$beers.beers | Export-Excel


function Get-AzureOpenAIDalleImage {
    [CmdletBinding()]
    Param
    (
        [Parameter(Mandatory)]
        [string]$prompt,
        [ValidateSet('256', '512', '1024')]
        $Size = 1024,
        $Images = 1,
        [string]$apiVersion = "2023-06-01-preview"
    )

    # Azure OpenAI metadata variables
    $openai = @{
        api_key  = [System.Runtime.InteropServices.marshal]::PtrToStringAuto([System.Runtime.InteropServices.marshal]::SecureStringToBSTR($AzureOpenAICred.Password))
        api_base = $AzureOpenAICred.UserName
    }

    # Header for authentication
    $headers = [ordered]@{
        'api-key' = $openai.api_key
    }

    $targetSize = switch ($Size) {
        256 { '256x256' }
        512 { '512x512' }
        1024 { '1024x1024' }     
    }

    # Adjust these values to fine-tune completions
    $body = [ordered]@{
        prompt = $prompt
        size   = $targetSize
        n      = $Images
    } | ConvertTo-Json

    # Call the API to generate the image and retrieve the response
    $url = "$($openai.api_base)/openai/images/generations:submit?api-version=$($apiVersion)"

    $submission = Invoke-RestMethod -Uri $url -Headers $headers -Body $body -Method Post -ContentType 'application/json' -ResponseHeadersVariable submissionHeaders
    $operation_location = $submissionHeaders['operation-location'][0]

    $status = ''
    while ($status -ne 'succeeded') {
        Start-Sleep -Seconds 1
        $response = Invoke-RestMethod -Uri $operation_location -Headers $headers
        if ($response.status -eq 'failed') {
            Write-Error "Image Generation Failed with: $($response.error.code) and message: $($response.error.message)"
            exit 
        }
        $status = $response.status
    }

    # Retrieve the generated image
    $generatedImages = @()

    # Set the directory for the stored image
    $image_dir = Join-Path -Path $pwd -ChildPath 'images'

    # If the directory doesn't exist, create it
    if (-not(Resolve-Path $image_dir -ErrorAction Ignore)) {
        New-Item -Path $image_dir -ItemType Directory
    }
    
    $i = 1
    foreach ($generatedImage in $response.result.data.url) {
        $image_url = $generatedImage
        # Initialize the image path (note the filetype should be png)
        $ts = (get-date -Uformat %T).ToString().Replace(":", "-")
        $image_path = Join-Path -Path $image_dir -ChildPath "$($ts)-$($i).png"
        Invoke-WebRequest -Uri $image_url -OutFile $image_path  # download the image
        $generatedImages += $image_path
        $i = $i + 1
    }
    return $generatedImages
}