[CmdletBinding()]
param (
  [string]
  $excelFile = '', #"$(Split-Path -Path $MyInvocation.MyCommand.Path)\..\PhoneUserConfig.xlsx",
  [string]
  $templateDirectory = '',
  [string]
  $outputDirectory = ''#"$(Split-Path -Path $MyInvocation.MyCommand.Path)\..\gsconfig"
)
$excelFile = "$(Split-Path -Path $MyInvocation.MyCommand.Path)\..\PhoneUserConfig.xlsx"
$templateDirectory = "$(Split-Path -Path $MyInvocation.MyCommand.Path)\Templates"
$outputDirectory = "$(Split-Path -Path $MyInvocation.MyCommand.Path)\..\gsconfig"

. "$(Split-Path -Path $MyInvocation.MyCommand.Path)\PhoneConfigHelpers.ps1"

if (-not (Import-RequiredModules)) {
  throw "Unable to load required modules."
}

# Verify that the workbook contains the required worksheets.
$worksheets = Get-WorksheetNames -Path $excelFile
$missingSheets = @("Users", "Accounts") | Where-Object { !$worksheets.Contains($_) }
if ($missingSheets) {
  throw "The workbook is missing required worksheets: $($missingSheets -join ', ')"
}

# Import the Snipe-IT credentials
. "$PSScriptRoot\SnipeConfig.ps1"


# Retrieve info from Snipe-IT
Set-Info -url $SnipeITURL -apiKey $SnipeITAPIKey

$IPPhoneCategory = Get-Category -search $SnipeITSearchCategory

$phones = Get-Asset -category_id $IPPhoneCategory.id | Where-Object { $null -ne $_.assigned_to.username }
$locations = Get-SnipeitLocation
$modelNumbers = $phones | ForEach-Object { $_.model_number } | Select-Object -Unique


# Import basic data from excel
$users = Import-XLSX -Path $excelFile -Sheet Users
$accounts = Import-XLSX -Path $excelFile -Sheet Accounts


$models = @{ }

# Phone Model initialization
foreach ($model in $modelNumbers) {
  $modelInfo = @{
    ModelSettingsPath = Join-Path -Path $templateDirectory "$model.xlsx"
    TemplatePath      = Join-Path -Path $templateDirectory "$model.txt"
    
    PCodes            = @{ }
  }

  if (-not (Test-Path $modelInfo.ModelSettingsPath)) {
    Write-Error "Unable to find settings for $model. This file must exist: $($modelInfo.ModelSettingsPath)"
  }
  if (-not (Test-Path $modelInfo.TemplatePath)) {
    Write-Error "Unable to find template for $model. This file must exist: $($modelInfo.TemplatePath)"
  }

  $modelInfo.SettingCodes = Get-SettingCodes -Path $modelInfo.ModelSettingsPath
  $modelInfo.AccountCodes = Get-AccountCodes -Path $modelInfo.ModelSettingsPath
  $modelInfo.PCodes = Get-PCodes -Path $modelInfo.TemplatePath
  $modelInfo.PCodes = Merge-HashTables $modelInfo.PCodes (Get-MPKCodes -Path $modelInfo.ModelSettingsPath)
  $models[$model] = $modelInfo
}

#######################################
# Generate each phone's configuration #
#######################################
foreach ($phone in $phones) {
  # Start by verifying some stuff
  $mac = $phone.custom_fields."$SnipeITMACField".value
  if ($mac -isnot [string]) {
    Write-Warning "$($phone.asset_tag) does not have a LAN MAC address assigned."
    continue
  }

  $mac = $mac.Replace(':', '')

  if ($mac.Length -ne 12) {
    Write-Warning "$($phone.asset_tag) has an invalid MAC property ($mac)."
    continue
  }

  $phoneModel = $models[$phone.model_number]
  $pCodes = $phoneModel.PCodes + @{ }

  # Find the User info

  $assignedTo = $phone.assigned_to.username.ToLower()

  $user = $users | Where-Object { $_.Email.ToLower() -eq $assignedTo }

  if (!$user) {
    Write-Warning "No user info found for $($phone.assigned_to.username) ($($phone.asset_tag))."
    continue
  }
  $phoneSettingValues = ConvertTo-Hashtable $user

  # Assign the Phone's host name
  $phoneSettingValues["HostName"] = $phone.asset_tag

  # Find the accounts assigned to the user
  $phoneAccounts = @{ }
  foreach ($accountNumber in $user | Get-Member -Name "Account*" | ForEach-Object { $_.Name }) {
    $accountID = $user.$accountNumber
    $phoneAccounts[$accountNumber] = ConvertTo-Hashtable ($accounts | Where-Object { $_.ID -eq $accountID })
  }

  # Calculate the weather code from the Snipe-IT Location
  $location = $phone.rtd_location.id, $phone.location.id | Select-Object -first 1
  $locationCode = ""

  if ($null -ne $location) {
    $location = $locations | where-object { $_.id -eq $location }
    $locationCode = @($location.city, $location.state, $location.country) -join ","
    $phoneSettingValues["WeatherCityCode"] = $locationCode
    $phoneSettingValues["WeatherCityCodeAutomatic"] = "0"
  }



  $phoneSettingsPCodes = Merge-PCodeValues $phoneModel.SettingCodes $phoneSettingValues
  $accountPCodes = Merge-AccountPCodes $phoneModel.AccountCodes $phoneAccounts

  $combinedPCodes = Merge-HashTables $phoneSettingsPCodes $accountPCodes
  $pCodes = Merge-HashTables $pCodes $combinedPCodes

  $pCodeString = Convert-PCodesToString -PCodes $pCodes
  $encodedString = [com.grandstream.provision.TextEncoder]::Encode($pCodeString, $mac, $false)

  $filename = Join-Path $outputDirectory "cfg$($mac.ToLower())"

  $encodedString | Set-Content -Path $filename -Encoding Byte
}