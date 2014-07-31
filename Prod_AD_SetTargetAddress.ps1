Clear-Host

#I hate feeling stupid. And I'm hungry.

$ScriptName = $MyInvocation.MyCommand.Name -replace '.ps1'
$ScriptPath = $MyInvocation.MyCommand.Definition -replace $MyInvocation.MyCommand.Name
$ScriptDate = Get-Date -Format yyMMddHHmm
$LogFilePath = $ScriptPath + 'logs\'
$LogFileName = "$ScriptName-Log-$ScriptDate.log"
$LogFile = $LogFilePath + $LogFileName

Add-LogEntry -LogType 'INFO' -LogText "Script $ScriptName launched from $ScriptPath"
Add-LogEntry -LogType 'INFO' -LogText "Ouput file $OutFileName will be created in $OutFilePath"
Write-Host "Errors will be logged to $LogFile" -ForegroundColor Cyan

Function Add-LogEntry 
{
  param
  (
    [System.String]
    $LogType,
    
    [System.String]
    $LogText
  )
  
  Switch ($LogType)
  {
    'info' {$TextColor = 'white'; Break}
    'success' {$TextColor = 'green'; Break}
    'warning' {$TextColor = 'yellow'; Break}
    'error' {$TextColor = 'red'; Break}
    default {$TextColor = 'white'}
  }
  
  $LogDate = Get-Date -Format yyMMdd-HH:mm:ss
  $LogEntry = "$LogDate - $LogType - $LogText"
  $LogEntry | Out-File $LogFile -Append
  
  Write-Host $LogText -ForegroundColor $TextColor
}

$WorkingFile = 'c:\externalset.csv'

If (!(Test-Path $WorkingFile))
{
  Add-LogEntry -LogType 'ERROR' -LogText "File $WorkingFile could not be found"
  Exit
}


Write-Host "Opening file $WorkingFile and validating fields ..."
Add-LogEntry -LogType 'SUCCESS' -LogText "Opening file $WorkingFile and validating fields"

$MbxList = @()  
If (!($MbxList = Import-Csv -Path $WorkingFile))
{
  Add-LogEntry -LogType 'ERROR' -LogText "File $WorkingFile could not be opened"
  Exit
}

$MbxProcessed = 0
$MbxCount = $MbxList.Count
If ($MbxCount -lt 1)
{
  AddLog -LogType 'ERROR' -LogText "The input file $WorkingFile doesn't contain any mailboxes"
  Exit
}
Else
{
  AddLog -LogType 'INFO' -LogText "The input file $WorkingFile contains $MbxCount mailboxes"
}

$CsvHeaders = @()
$CsvHeaders = $MbxList | Get-Member
$ColumnNames = $CsvHeaders.Name

$Validation = @()
$ValidationError = $null

Switch ($ColumnNames)
{
  'name'
  {
    Write-Host 'The UnivADID column was detected' -ForegroundColor Green
    Add-LogEntry -LogType 'SUCCESS' -LogText 'The UnivADID column was detected'
    $Validation += 'name'
  }
  'proxyAddresses'
  {
    Write-Host 'The ExchangeSystem column was detected' -ForegroundColor Green
    Add-LogEntry -LogType 'SUCCESS' -LogText 'The ExchangeSystem column was detected'
    $Validation += 'proxyAddresses'
  }
}

Switch ($Validation)
{
  (!('name')) {
    Add-LogEntry -LogType 'ERROR' -LogText 'The name column could not be detected'
    $ValidationError += 1
  }
  (!('proxyAddresses')) {
    Add-LogEntry -LogType 'ERROR' -LogText 'The proxyAddresses column could not be detected'
    $ValidationError += 1
  }
}

If ($ValidationError -ne $null)
{
  Add-LogEntry -LogType 'ERROR' -LogText 'The required fields could not be validated in the input file. The script has terminated.'
  Exit
}
Else
{
  Add-LogEntry -LogType 'SUCCESS' -LogText 'The required fields were detected in the input file. Continuing to process.'
}

Write-Host 'Processing file ...'
$MbxCount = $MbxList.Count
Add-LogEntry -LogType 'INFO' -LogText "The input file contains $MbxCount mailboxes"

ForEach ($Mbx in $MbxList)
{

  $Name = $Mbx.Name
  $Proxy = $Mbx.ProxyAddresses
  $MbxProcessed ++
  
  Write-Host "Processing mailbox $MbxProcessed of $MbxCount"
  Add-LogEntry -LogType 'INFO' -LogText "Adding $proxy for $Name" 

  $MbxSuccess = Set-RemoteMailbox -Identity $Name -RemoteRoutingAddress @{Add=$Proxy}

  If(!($MbxSuccess))
  {
    Add-LogEntry -LogType 'ERROR' -LogText "Unable to set $Proxy as routing address for $Name"
  }
  ElseIf ($MbxSuccess)
  {
    Add-LogEntry -LogType 'SUCCESS' -LogText "$Proxy set as routing address for $Name"
  }
}

Add-LogEntry -LogType 'INFO' -LogText "Processing complete. Please review $LogFile for errors."
Exit
