Function Import-RequiredModules {
  # Check to ensure that the PSExcel module is available...
  if ($null -eq (Get-Module PSExcel -ListAvailable)) {
    Install-Module PSExcel -Scope CurrentUser
  }

  # ...and loaded
  if ($null -eq (Get-Module PSExcel)) {
    if (-not (Import-Module PSExcel -PassThru)) {
      throw "Unable to find/install PSExcel module. Please install the PSExcel module and try again."    
    }
  }

  # Check to ensure that the SnipeITPS module is available...
  if ($null -eq (Get-Module SnipeitPS -ListAvailable)) {
    Install-Module SnipeitPS -Scope CurrentUser
  }

  # ...and loaded
  if ($null -eq (Get-Module SnipeitPS)) {
    if (-not (Import-Module SnipeitPS -PassThru)) {
      throw "Unable to find/install SnipeitPS module. Please install the SnipeitPS module and try again."    
    }
  }

  if ($null -eq ('com.grandstream.provision.TextEncoder' -as [type])) {
    if ($null -eq (Add-Type -Path "$PSScriptRoot\gs_config.dll" -PassThru)) {
      throw "Unable to find gs_config.dll in the same directory as the script."
    }
  }

  return $true
}

# https://stackoverflow.com/q/8800375
# Modified, but some of the basic stuff is here.
function Merge-HashTables($htold, $htnew) {  
  $newTable = $htold + @{ }
  $htold.Keys | foreach-object {
    if ($htnew.ContainsKey($_)) {
      $newTable.Remove($_)
    }
  }
  return $newTable + $htnew
}

Function Get-WorksheetNames {
  [CmdletBinding()]
  param (
    [Parameter()]
    [string]
    $Path
  )

  $wb = New-Excel -Path $path
  $worksheets = $wb.Workbook.Worksheets | ForEach-Object { $_.Name }
  Close-Excel $wb
  return $worksheets
}


Function Get-SettingCodes {
  [CmdletBinding()]
  param (
    # Specifies the Excel file to load.
    [Parameter()]
    [string]
    $Path
  )
  $rawCodes = Import-XLSX -Path $Path -Sheet "SettingCodes"
  $settings = @{ }

  foreach ($raw in $rawCodes) {
    $setting = @{
      Code         = -1
      DefaultValue = ""
    }
    if (($raw.Code -is [string]) -and ($raw.Code.Contains("="))) {
      $equalsIndex = $raw.Code.IndexOf("=")
      $setting.Code = [int]$raw.Code.SubString(0, $equalsIndex)
      $setting.DefaultValue = $raw.Code.SubString($equalsIndex + 1)
    }
    else {
      $setting.Code = [int]$raw.Code
    }
    $settings[$raw.Name] = $setting
  }
  return $settings
}

Function Merge-AccountPCodes {
  [CmdletBinding()]
  param (
      [Parameter()]
      $AccountCodes,
      [Parameter()]
      $AccountValues
  )

  $result = @{}

  foreach ($accountCode in $AccountCodes.Keys) {
    $account = $AccountValues[$accountCode]
    if($null -eq $account) {
      $account = @{}
    }

    $accountPCodes = Merge-PCodeValues -Codes $accountCodes[$accountCode] -Values $account
    $result = Merge-HashTables $result $accountPCodes
  }

  return $result
}

Function Get-AccountCodes {
  [CmdletBinding()]
  param (
    # Specifies the Excel file to load.
    [Parameter()]
    [string]
    $Path
  )
  $rawCodes = Import-XLSX -Path $Path -Sheet "AccountCodes"
  $codes = @{ }
  $accountKeys = $rawCodes[0] | Get-Member -Name "Account*" -MemberType Properties | ForEach-Object { $_.Name }

  # For each account, generate an object containing the PCodes and their default values if specified
  # $codes should be something like this: 
  # @{
  #   Account1 = @{
  #     Active = @{
  #       Code = 271
  #       DefaultValue = "0"
  #     }
  #     DisplayName = @{
  #       Code = 270
  #       DefaultValue = ""
  #     }
  #   }
  # }
  foreach ($accountID in $accountKeys) {
    $account = @{ }
    foreach ($rawCode in $rawCodes) {
      $value = $rawCode.$accountID
      $code = -1
      $defaultValue = ""
      if (($value -is [string]) -and ($value.Contains("="))) {
        $equalsIndex = $value.IndexOf("=")
        $code = [int]$value.SubString(0, $equalsIndex)
        $defaultValue = $value.SubString($equalsIndex + 1)
      }
      else {
        $code = [int]$value
      }
      $account[$rawCode.Name] = @{
        Code         = $code
        DefaultValue = $defaultValue
      }
    }
    $codes[$accountID] = $account
  }
  return $codes
}

# $Codes should be something like this: 
# @{
#   Active = @{
#     Code = 271
#     DefaultValue = "0"
#   }
#   DisplayName = @{
#     Code = 270
#     DefaultValue = ""
#   }
#   SipServer = @{
#     Code = 270
#     DefaultValue = "onsip.cbnova.com"
#   }
# }
# 
# $values should be something like this:
# 
# @{
#   Active = 1
#   DisplayName = "Hello, World"
# }
# 
# Given the above inputs, the output is this
# @{
#   Active = 1 
#   DisplayName = "Hello, World"
#   SipServer = "onsip.cbnova.com"
# }
Function Merge-PCodeValues {
  [CmdletBinding()]
  param (
    [Parameter()]
    $Codes,
    [Parameter()]
    $Values
  )

  $output = @{ }
  $Values = ConvertTo-Hashtable $Values

  foreach ($key in $Codes.Keys) {
    $code = $codes[$key]
    if ($Values.ContainsKey($key)) {
      $output[$code.Code] = $values[$key]
    }
    else {
      $output[$code.Code] = $code.DefaultValue
    }
  }

  return $output

}

Function ConvertTo-Hashtable {
  [CmdletBinding()]
  param (
    [Parameter()]
    $value
  )
  if ($value -is [hashtable]) {
    return $value
  }
  elseif ($value -is [PSCustomObject]) {
    $ht2 = @{ }
    $value.psobject.properties | ForEach-Object { $ht2[$_.Name] = $_.Value }

    return $ht2 
  }
}

Function Get-PCodes {
  [CmdletBinding()]
  param (
    # Specifies the txt file to load.
    [Parameter()]
    [string]
    $Path
  )

  $lines = Get-Content -Path $Path | Where-Object { $_.StartsWith("P") -and $_.Contains("=") } # Only retrieve lines that have values
  $codes = @{ }

  foreach ($line in $lines) {
    $equalsIndex = $line.IndexOf("=")
    $code = [int]$line.SubString(1, $equalsIndex - 1)
    $value = $line.SubString($equalsIndex + 1).Trim()
    if ($codes.ContainsKey($code)) {
      Write-Warning "$Path contains code $code more than once."
    }
    $codes[$code] = $value
  }
  
  return $codes
}



Function Get-MPKCodes {
  [CmdletBinding()]
  param (
    # Specifies the Excel file to load.
    [Parameter()]
    [string]
    $Path
  )

  $valueSheetName = "MPKValues"
  $codeSheetName = "MPKCodes"

  $worksheets = Get-WorksheetNames -Path $Path

  $missingSheets = @($valueSheetName, $codeSheetName) | Where-Object { !$worksheets.Contains($_) }
  if ($missingSheets) {
    throw "The workbook $path is missing required worksheets: $($missingSheets -join ', ')"
  }

  $codes = Import-XLSX -Path $Path -Sheet $codeSheetName
  $values = Import-XLSX -Path $Path -Sheet $valueSheetName
  
  $pCodeProperties = $codes | Select-Object -First 1 | Get-Member -MemberType Properties | Where-Object { !@("MPKID", "PhysicalButtonLocation").Contains($_.Name) } | ForEach-Object { $_.Name }

  $valueProperties = $values | Select-Object -First 1 | Get-Member -MemberType Properties | ForEach-Object { $_.Name }

  $missingValueColumns = $valueProperties | Where-Object { !$pCodeProperties.Contains($_) }
  if ($missingValueColumns) {
    throw "The worksheet $valueSheetName is missing required columns: $($missingValueColumns -join ', ')"
  }

  $codesTable = @{ }
  foreach ($code in $codes) {
    $codesTable[[int]$code.PhysicalButtonLocation] = $code
  }

  $pcodes = @{ }
  for ($i = 0; $i -lt $values.Count; $i++) {
    $value = $values[$i]
    $codeRow = $codesTable[$i + 1]
    foreach ($code in $pCodeProperties) {
      $pcodes[$codeRow.$code] = [string]$value.$code
    }
  }

  return $pcodes
  
}

Function Convert-PCodesToString {
  [CmdletBinding()]
  param (
      [Parameter()]
      [HashTable]
      $PCodes
  )
  $sb = [System.Text.StringBuilder]::new()
  foreach ($item in $PCodes.Keys) {
    [void]$sb.AppendFormat("P{0}={1}`r`n", $item, $pcodes[$item])
  }
  return $sb.ToString()
}
