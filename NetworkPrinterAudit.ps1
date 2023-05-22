<#==============================================================================
         File Name : NetworkPrinterAudit.ps1
   Original Author : Kenneth C. Mazie (kcmjr AT kcmjr DOT com)
                   :
       Description : This script polls all specified systems, read the currently logged
                   : on user, then identifies the printers currently installed. If the
                   : associated print server matches a "bad server" list The results are
                   : noted on screen and/or in an Excel spreadsheet. Output can be
                   : "server/printer" (full) or just "server" (brief).
                   :
         Arguments : Named command line parameters: (all are optional)
                   :
             Notes : This script was originally used during print server migration.
                   : The intent was to identify users who still had mapped printers
                   : pointing to the older print servers, hence the output only showing
                   : "bad" servers. Because the logged on user environment is volatile and
                   : may change at logoff you cannot read remote systems to determine printers.
                   : This is a best effort attempt to do just that, read the volatile session
                   : to collect installed printers.
                   :
          Warnings : None
                   :
             Legal : Public Domain. Modify and redistribute freely. No rights reserved.
                   : SCRIPT PROVIDED "AS IS" WITHOUT WARRANTIES OR GUARANTEES OF
                   : ANY KIND. USE AT YOUR OWN RISK. NO TECHNICAL SUPPORT PROVIDED.
                   :
           Credits : Code snippets and/or ideas came from many sources around the web.
                   :
    Last Update by : Kenneth C. Mazie
   Version History : v1.00 - 06-04-14 - Original
    Change History : v1.01 - 00-00-00 -
                   :
#===============================================================================#>
<#PSScriptInfo
.VERSION 1.00
.AUTHOR Kenneth C. Mazie (kcmjr AT kcmjr DOT com)
.DESCRIPTION
This script polls all specified systems, read the currently logged on user, then identifies
the printers currently installed. If the associated print server matches a "bad server" list
The results are noted on screen and/or in an Excel spreadsheet. Output can be "server/printer" (full)
or just "server" (brief).
#>

#requires -version 5.0

clear-host 
Import-module ActiveDirectory 
$ErrorActionPreference = "silentlycontinue"
$Script:Console = $true
$Script:UseExcel = $true
$Script:Brief = $false

#--[ Service account configuration with AES key for hardcoded service account ]--
#--[ See https://www.powershellgallery.com/packages/CredentialsWithKey/1.10/DisplayScript ]--
#$Script:DN = "Domain.com" #--[ Correct this for your domain if used ]--
#$Script:UN = 'serviceaccount@'+($Script:DN.Split(".")[0]) #--[ Correct this for your domain if used ]--
#$Script:EPW = '76490a5345MgB8AHIAegB2AHYAZQAxAGIATgBaADcAYLO+Eyj267L6/wBtAHAAWQB6AH5wBkADEANAA4AGQAZgA3ADIAYQAwADYAZAA3AGUAZAZgGANAAyADUAYQA2AGQAZAA2A6743f0423413b16AEcAaAB1AFCh7HCvAWnHbgBkAGYAZAA='
#$Script:BA = [System.Convert]::FromBase64String('kdhCh7HCv67L6/AWnHbuTeJ7ILoAeQA0uTeJ7IXADQe8mE=') #--[ Correct this for your domain if used ]--
#$Script:SC = New-Object -TypeName System.Management.Automation.PSCredential -ArgumentList $Script:UN, ($Script:EPW | ConvertTo-SecureString -Key $Script:BA)
#$Script:SP = $SC.GetNetworkCredential().Password
$Script:SC = Get-Credential

[string]$DateTime = Get-Date -Format MM-dd-yyyy_HHmmss 
$Bad = @("printsrv1","printsrv2")    #--[ A list of "bad" servers to flag. If a general report is needed then list your valid servers ]--

#$Type = "server"
$Type = "pc"                #--[ Select which type of system to query ]--
#$Type = "test"

If ($Script:UseExcel){
    $Row = 1
    $Col = 1
    #--[ Create a new Excel object ]--
    $Excel = New-Object -Com Excel.Application
    If ($Script:Console){
        $Excel.visible = $True
        $Excel.DisplayAlerts = $true
        $Excel.ScreenUpdating = $true
        $Excel.UserControl = $true
        $Excel.Interactive = $true
    }Else{
        $Excel.visible = $False
        $Excel.DisplayAlerts = $false 
        $Excel.ScreenUpdating = $false 
        $Excel.UserControl = $false
        $Excel.Interactive = $false
    }
    $Workbook = $Excel.Workbooks.Add()
    $WorkSheet = $Workbook.WorkSheets.Item(1)
   
    #--[ Write Worksheet title ]--
    $WorkSheet.Cells.Item($Row,$Col) = "Network Printer Audit Report - ($DateTime)"
    $WorkSheet.Cells.Item($Row,$Col).font.bold = $true
    $WorkSheet.Cells.Item($Row,$Col).font.underline = $true
    $WorkSheet.Cells.Item($Row,$Col).font.size = 18
    $WorkSheet.Cells.Item($Row,$Col).HorizontalAlignment = -4108
    #--[ Write worksheet column headers ]--
    $Row++
    $WorkSheet.Cells.Item($Row,$Col) = "TARGET:"
    $WorkSheet.Cells.Item($Row,$Col).font.bold = $true
    $WorkSheet.Cells.Item($Row,$Col).HorizontalAlignment = 1
# $WorkSheet.Cells.Item($Row,$Col).Borders.Item(10).LineStyle = 1 #--[ optional formatting ]--
# $WorkSheet.Cells.Item($Row,$Col).Borders.Item(10).Weight = 4
    $Col++
    $WorkSheet.Cells.Item($Row,$Col) = "STATUS:"
    $WorkSheet.Cells.Item($Row,$Col).font.bold = $true
    $WorkSheet.Cells.Item($Row,$Col).HorizontalAlignment = 1
# $WorkSheet.Cells.Item($Row,$Col).Borders.Item(10).LineStyle = 1
# $WorkSheet.Cells.Item($Row,$Col).Borders.Item(10).Weight = 4
    $Col++
    $WorkSheet.Cells.Item($Row,$Col) = "USER:"
    $WorkSheet.Cells.Item($Row,$Col).font.bold = $true
    $WorkSheet.Cells.Item($Row,$Col).HorizontalAlignment = 1
# $WorkSheet.Cells.Item($Row,$Col).Borders.Item(10).LineStyle = 1
# $WorkSheet.Cells.Item($Row,$Col).Borders.Item(10).Weight = 4
    $Col++
    $WorkSheet.Cells.Item($Row,$Col) = "SID:"
    $WorkSheet.Cells.Item($Row,$Col).font.bold = $true
    $WorkSheet.Cells.Item($Row,$Col).HorizontalAlignment = 1
# $WorkSheet.Cells.Item($Row,$Col).Borders.Item(10).LineStyle = 1
# $WorkSheet.Cells.Item($Row,$Col).Borders.Item(10).Weight = 4
    $WorkSheet.application.activewindow.splitcolumn = 0
    $WorkSheet.application.activewindow.splitrow = 2
    $WorkSheet.application.activewindow.freezepanes = $true
    $MergeCells = $WorkSheet.Range("A1:F1")
    $MergeCells.Select() 
    $MergeCells.MergeCells = $true
    $Resize = $WorkSheet.UsedRange
    [void]$Resize.EntireColumn.AutoFit()
}

$Script:TargetSystemList = ""
If ($Type -eq "server"){ 
    $Script:TargetSystemList = get-adcomputer -Credential $Script:SC -Filter {(operatingSystem -like "*server*") -and (operatingsystem -Like "*Windows*") -and (Enabled -eq "True") -and (name -notlike "*dc0*") -and (name -notlike "*esx*") -and ($_.name -notlike "*vcsa*")} -properties name 
}ElseIf ($Type -eq "pc"){
    $Script:TargetSystemList = Get-ADComputer -Credential $Script:SC -Properties name, operatingsystem -Filter * | sort name | where {($_.operatingsystem -NotLike "*server*") -and ($_.operatingsystem -Like "*windows*")} 
}Else{ #--[ Test Only ]--
    $Script:TargetSystemList = get-adcomputer -Credential $Script:SC -Filter {(Enabled -eq "True") -and (name -like "*testpc*")} -properties name | Where-object {Test-Connection -computername $($_.name) -count 1 -quiet}
}

$Count = $Script:TargetSystemList.Count-1
$Row++
foreach ($Script:TargetSystem in $Script:TargetSystemList){
    $Col = 1
    Start-Sleep -Milliseconds 500
    $Script:Target = $TargetSystem.name
    $WorkSheet.Cells.Item($row,1) = $Target
    If ($Script:Console){Write-host `n"==[ "$Script:Target" ("$Count "remaining) ]===========================================================" -ForegroundColor Yellow }
    If (Test-Connection -computername $Script:Target -count 1 -quiet){
        $WorkSheet.Cells.Item($row,2).Font.ColorIndex = 10
        $WorkSheet.Cells.Item($row,2) = "Online"
        $CurrentUser = Get-WmiObject -ComputerName $Script:Target -Credential $Script:SC  -Class win32_computersystem | Select-Object -ExpandProperty Username 
        $SID = $CurrentUser | ForEach-Object { ([System.Security.Principal.NTAccount]$_).Translate([System.Security.Principal.SecurityIdentifier]).Value }

        If ([string]::IsNullOrEmpty($CurrentUser)){
            Write-Host " No Current User"
            $WorkSheet.Cells.Item($row,3).Font.ColorIndex = 1
            $WorkSheet.Cells.Item($row,3) = "No Current User"
        }Else{
            Write-Host " Current User Name = "$CurrentUser -ForegroundColor Cyan
            $WorkSheet.Cells.Item($row,3) = $CurrentUser
            Write-Host " Current User SID = "$SID -ForegroundColor Cyan
            $WorkSheet.Cells.Item($row,4) = $SID
            
            $RegPath = "REGISTRY::HKEY_USERS\"+$SID+"\Printers\Connections"
            $Col = 5
            $Result = Invoke-command -ComputerName $Script:Target -Credential $SC -ScriptBlock { 
                $Printers = Get-ChildItem $Using:RegPath # | select name
                $Found = @()
                foreach ($Item in $Printers){
                    $Srv = ($Item.name).split(",")[2]
                    If ($Using:Bad -contains $Srv) {
                        If ($Brief){
                            Write-host " "$Srv -ForegroundColor red
                            $Found += $Srv
                        }Else{    
                            $Srv = ($Item.name).split(",")[2]+"\"+($Item.name).split(",")[3]
                            Write-host " "$Srv -ForegroundColor red
                            $Found += $Srv
                        }
                    }
                }
                Return $Found
            }
            ForEach ($X in $Result){
                $WorkSheet.Cells.Item($row,$Col).Font.ColorIndex = 3
                $WorkSheet.Cells.Item($row,$Col) = $X
                $Col++
            }
        }
    } else {
        write-host " Computer Offline" 
        $WorkSheet.Cells.Item($row,2).Font.ColorIndex = 3
        $WorkSheet.Cells.Item($row,2) = "Offline" 
    }
    $Resize = $WorkSheet.UsedRange
    [void]$Resize.EntireColumn.AutoFit()
$Count--
$Row++
}

$Resize = $WorkSheet.UsedRange
[Void]$Resize.EntireColumn.AutoFit()

[string]$Script:FileName = "$PSScriptRoot\NetworkPrinterAudit_$DateTime.xlsx"
If ($UseExcel){    
    Try{
        $Workbook.SaveAs($Script:FileName)
        $Workbook.Saved = $true 
        $Workbook.Close() 
        $Excel.Quit()
        $Excel = $Null
        if ($Script:Console){Write-host "`nExcel Closed and Saved...`n" -ForegroundColor Cyan}
    }Catch{    
        if ($Script:COnsole){Write-host "`nThere was a problem closing and/or Saving the Excel spreadsheet...`n" -ForegroundColor Red}
    }    
}   

Write-host "--- Completed ---" -foregroundcolor red

