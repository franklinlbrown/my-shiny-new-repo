#Common functions; do not edit directly--use parameters to change behavior where appropriate for your script

#Create on-screen and log file pocessing information for monitoring
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
  
  $LogDate = Get-Date -Format yyMMdd:HH:mm:ss
  $LogEntry = "$LogDate - $LogType - $LogText"
  $LogEntry | Out-File $LogFile -Append
  
  Write-Host $LogText -ForegroundColor $TextColor
}

#Browse to and select an input file to work with
Function New-DialogOpenFile
{
  param
  (
    [System.String]
    $WindowTitle,
    
    [System.String]
    $InitialDirectory,
    
    [System.String]
    $Filter = 'All files (*.*)|*.*',
    
    [System.Management.Automation.SwitchParameter]
    $AllowMultiSelect
  )
  
  Add-Type -AssemblyName System.Windows.Forms
  $OpenFileDialog = New-Object System.Windows.Forms.OpenFileDialog
  $OpenFileDialog.AutoUpgradeEnabled = $true
  $OpenFileDialog.Title = $WindowTitle
  If (!([string]::IsNullOrWhiteSpace($InitialDirectory)))
  {
    $OpenFileDialog.InitialDirectory = $InitialDirectory
  }
  $OpenFileDialog.Filter = $Filter
  If ($AllowMultiSelect)
  {
    $openFileDialog.MultiSelect = $true
  }
  if ($Host.name -eq 'ConsoleHost')
  {
    $ShowHelp = $true
  }
  $OpenFileDialog.ShowHelp = $ShowHelp
  $OpenFileDialog.ShowDialog() > $null
  If ($AllowMultiSelect)
  {
    return $OpenFileDialog.Filenames
  }
  else
  {
    return $OpenFileDialog.Filename
  }
}

#Convert Excel spreadsheets to CSV files for more efficient handling
Function ConvertTo-CsvFromXls 
{
  param
  (
    [System.String]
    $XLwksh,
    
    [System.String]
    $XLsource,
    
    [System.String]
    $CsvTarget
  )
  
  $XLapp = New-Object -ComObject Excel.Application
  $XLapp.Visible = $False
  $XLapp.DisplayAlerts = $False
  
  If ($WkBk = $XLapp.Workbooks.Open($XLsource))
  {
    $WkBkName = $WkBk.Name
    Add-LogEntry -LogType 'SUCCESS' -LogText "Workbook $WkBkName imported. Processing ..."
  }
  Else
  {
    Add-LogEntry -LogType 'ERROR' -LogText "Workbook $WkBkName could not be imported. Exiting."
    Exit
  }
  ForEach ($WkSh in $WkBk.Worksheets)
  {
    $WkshName = $WkSh.Name
    $CsvName = $MigInfo+ '_' +$WkshName + '.csv'
    If ($XLwksh -eq 'all')
    {
      Add-LogEntry -LogType 'INFO' -LogText "Converting worksheet $WkshName. Creating temporary file $CsvName."
      
      $WkSh.SaveAs($CsvTarget + $CsvName, 6)
    }
    ElseIf ($WkSh.name -eq $XLwksh)
    {
      Add-LogEntry -LogType 'SUCCESS' -LogText "The requested $XLwksh worksheet was found in $WkBkName"
      Add-LogEntry -LogType 'INFO' -LogText "Converting worksheet $WkshName. Creating temporary file $CsvName in $TempDir"
      $CSVTempFile = $CsvTarget + $CsvName
      $OldName = $CSVTempFile + '.' +$ScriptDate
      If (Test-Path $CSVTempFile)
      {
        $NewName = $CsvName + '.' +$ScriptDate
        Rename-Item -Path $CSVTempFile -NewName $NewName -Force
      }
      $WkSh.SaveAs($CSVTempFile, 6)
    }
  }
  $WkBk.Close()
  $XLapp.Quit()
  stop-process -processname EXCEL
  [System.Runtime.Interopservices.Marshal]::ReleaseComObject($XLapp)
  
  If($CSVTempFile)
  {
    Return $CSVTempFile
  }
  Else
  {
    Add-LogEntry -LogType 'ERROR' -LogText "The requested $XLwksh worksheet was not found in $WkBkName"
    Exit
  }
}


Clear-Host

#Browse to your copy of the Migration SharePoint library and select a department migration list
$InputFile = New-DialogOpenFile `
    -WindowTitle 'Select Migration File' `
    -InitialDirectory "$env:UserProfile\SharePoint\Office 365 Exchange - Project Docum\Full Migration Department Lists\IN PROCESS\CURRENT EXO MIGRATION LIST" `
    -Filter 'Excel Workbooks (*.xls*)|*.xls*'

#Parse file name to get dept and migration date info for filenames
$InputFileName = Split-Path $InputFile -Leaf -Resolve
$MigInput = $InputfileName.split('.')
$MigInfo = $MigInput[0]

#Set variables for log and other output file creation
$ScriptName = $MyInvocation.MyCommand.Name -replace '.ps1'
$ScriptPath = $MyInvocation.MyCommand.Definition -replace $MyInvocation.MyCommand.Name
$ScriptDate = Get-Date -Format yyMMddHHmm

$OutFilePath = $ScriptPath + 'output\'
$LogFilePath = $ScriptPath + 'logs\'
$OutFileName = "$MigInfo-$ScriptName-$ScriptDate.csv"
$LogFileName = "$MigInfo-$ScriptName-$ScriptDate.log"
$OutputFile = $OutFilePath + $OutFileName
$LogFile = $LogFilePath + $LogFileName
$RemFile = $OutFilePath + $RemFileName
$MigListFile = $OutFilePath + $MigListFileName

Add-LogEntry -LogType 'INFO' -LogText "Script $ScriptName launched from $ScriptPath"
Add-LogEntry -LogType 'INFO' -LogText "Ouput file $OutFileName will be created in $OutFilePath"
Add-LogEntry -LogType 'INFO' -LogText "Errors will be logged to $LogFile"

#Open the Excel workbook, save the mailboxes worksheet to CSV
If (!($WorkingFile = ConvertTo-CsvFromXls -XLsource $InputFile -CsvTarget $OutFilePath -XLwksh 'all_mailboxes'))
{
  Add-LogEntry -LogType 'ERROR' -LogText "File $WorkingFile could not be found"
  Exit
}

Write-Host "Opening file $WorkingFile and validating fields ..."
Add-LogEntry -LogType 'SUCCESS' -LogText "Opening file $WorkingFile and validating fields"

#Make sure we can open the CSV we created and that has data
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

#Validate that the required fields exist by examining the headers
$CsvHeaders = @()
$CsvHeaders = $MbxList | Get-Member
$ColumnNames = $CsvHeaders.Name

$Validation = @()
$ValidationError = $null

If (!($ColumnNames -eq 'UnivADID'))
{
    Add-LogEntry -LogType 'ERROR' -LogText 'The UnivADID column could not be detected'
    Exit
}
Write-Host 'The UnivADID column was detected' -ForegroundColor Green
Add-LogEntry -LogType 'SUCCESS' -LogText 'The UnivADID column was detected'
Add-LogEntry -LogType 'SUCCESS' -LogText 'The required fields were detected in the input file. Continuing to process.'

Write-Host 'Processing file ...'
$MbxCount = $MbxList.Count
Add-LogEntry -LogType 'INFO' -LogText "The input file contains $MbxCount mailboxes"

ForEach ($Mbx in $MbxList)
{

  #For each row of the spreadsheet, define variables for the data
  $Name = $Mbx.DisplayName
  $ADID = $Mbx.UnivADID
  $EXOdomain = '@hu.mail.onmicrosoft.com'

  #Count the number of mailboxes processed to provide progress indication
  $MbxProcessed ++
  
  Write-Host '--------------------------------------------------------------------------------------------------------------' -ForegroundColor White
  Write-Host "     Processing mailbox $ProcessedMbx of $MbxCount - $ExSystem account $ADID ($Name)" -ForegroundColor Cyan
  Write-Host '--------------------------------------------------------------------------------------------------------------' -ForegroundColor White

  Add-LogEntry -LogType 'INFO' -LogText "Validating ADID $ADID for $Name"
  
  #Look the account up in AD using the ADID from the CSV
  $ICEuser = Get-ADUser $ADID -Properties * | Select-Object displayName,legacyExchangeDN -ErrorAction SilentlyContinue
  If (!($ICEuser))
    {
    $AcctStatus = "No user account for $ADID found in University"
    Add-LogEntry -LogType 'ERROR' -LogText "User $ADID ($Name) not found in UNIVERSITY domain"
    Break
    }

  #Calculate the EXO routing address
  $Proxy = "$ADID@$EXOdomain"

  #Calculate the X500 routing address
  $LegDN = $ICEuser.legacyExchangeDN
  $X500= "X500:$LegDN"

  #Add the X500 address to the mailbox and validate
  $MbxLegSuccess = Set-Mailbox -Identity $ADID -EmailAddresses @{Add=$X500}
  If(!($MbxX500Success))
  {
    Add-LogEntry -LogType 'ERROR' -LogText "Unable to add X500 address to mailbox $Name"
  }
  ElseIf ($MbxX500Success)
  {
    Add-LogEntry -LogType 'SUCCESS' -LogText "X500 address $X500 added to mailbox $Name"
  }

  #Add the EXO routing address and validate
  $MbxProxySuccess = Set-Mailbox -Identity $ADID -EmailAddresses @{Add=$Proxy}
  If(!($MbxProxySuccess))
  {
    Add-LogEntry -LogType 'ERROR' -LogText "Unable to add $Proxy to mailbox $Name"
  }
  ElseIf ($MbxProxySuccess)
  {
    Add-LogEntry -LogType 'SUCCESS' -LogText "Routing address $Proxy added to mailbox $Name"
  }
}

Add-LogEntry -LogType 'INFO' -LogText "Processing complete. Please review $LogFile for errors."
Exit
