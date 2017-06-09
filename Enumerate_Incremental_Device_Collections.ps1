#############################################################################################################################################
###~ SCCM - Manage Incremental Device Collections ###########################################################################################
#
#~ Brian Tancredi
#~ Created: 2017-04-28
#~ Modified: 2017-04-28
#
#~ References:
#~ https://blogs.technet.microsoft.com/configmgrdogs/2015/05/26/enable-or-disable-incremental-collection-updates-via-powershell/
#~ https://gallery.technet.microsoft.com/scriptcenter/Find-all-Collections-with-6d1ea160
#
#############################################################################################################################################
 
#############################################################################################################################################
###~ Assigning Variables and Declaring Functions  ###########################################################################################
#############################################################################################################################################

#~ Import Modules and Connect to Site
Import-Module "$($ENV:SMS_ADMIN_UI_PATH)\..\ConfigurationManager.psd1" # Import the ConfigurationManager.psd1 module 
Set-Location "xxx:" # Set the current location to be the site code.

#~ ***Modify with Device Collection Names for Selection 4***
$target_Collections = @(
    "Collection_Name_1",
    "Collection_Name_2"
)

#~ The following refresh types exist for ConfigMgr collections 
$6 = "Incremental and Periodic Updates" 
$4 = "Incremental Updates Only" 
$2 = "Periodic Updates only"
$1 = "Manual Update only"

Function Get_Incrementals{
    Write-Host "`nQuerying Incremental Device Collections...`nPlease Wait...`n" -ForegroundColor White -BackgroundColor Black
    $refreshtypes = "4","6" 
    $global:CollectionsWithIncrement = Get-CMDeviceCollection | Where-Object {$_.RefreshType -in $refreshtypes} | Sort Name
    $total4 = ($CollectionsWithIncrement.RefreshType -eq 4).Count
    $total6 = ($CollectionsWithIncrement.RefreshType -eq 6).Count 
    $total = $CollectionsWithIncrement.Count
    Clear-Host
    Write-Host "`nIncremental Device Collections Counts...`n" -ForegroundColor White -BackgroundColor Black
    Write-Host "All Collections with Increment : " -ForegroundColor Cyan -NoNewline ; Write-Host "$total" -ForegroundColor Yellow
    Write-Host "Collections with $6 : " -ForegroundColor Cyan -NoNewline ; Write-Host "$total6" -ForegroundColor Yellow
    Write-Host "Collections with $4 : " -ForegroundColor Cyan -NoNewline ; Write-Host "$total4" -ForegroundColor Yellow
    Write-Host "`n"

    Create_Csv
}

Function Get_Others{
    Write-Host "`nQuerying Other Device Collections...`nPlease Wait...`n" -ForegroundColor White -BackgroundColor Black
    $refreshtypes = "2","1" 
    $global:CollectionsWithoutIncrement = Get-CMDeviceCollection | Where-Object {$_.RefreshType -in $refreshtypes} | Sort Name
    $total2 = ($CollectionsWithoutIncrement.RefreshType -eq 2).Count
    $total1 = ($CollectionsWithoutIncrement.RefreshType -eq 1).Count 
    $total = $CollectionsWithoutIncrement.Count
    Clear-Host
    Write-Host "`nOther Device Collections Counts...`n" -ForegroundColor White -BackgroundColor Black
    Write-Host "All Collections Non-Incremental: "-ForegroundColor Cyan -NoNewline ; Write-Host "$total" -ForegroundColor Yellow
    Write-Host "Collections with $2 : " -ForegroundColor Cyan -NoNewline ; Write-Host "$total2" -ForegroundColor Yellow
    Write-Host "Collections with $1 : " -ForegroundColor Cyan -NoNewline ; Write-Host "$total1" -ForegroundColor Yellow
    Write-Host "`n"

    Create_Csv
}

Function Create_Csv{
    Select_Report
    #~ Reporting for Selection 1 (with)
    If ($report -eq 'Y' -and $selection -eq '1'){
        Select_Save
        $collections = @()
        ForEach ($collection in $CollectionsWithIncrement){ 
            $object = New-Object -TypeName PSobject
            $object| Add-Member -Name CollectionName -value $collection.Name -MemberType NoteProperty
            $object| Add-Member -Name CollectionID -value $collection.CollectionID -MemberType NoteProperty
            $object| Add-Member -Name MemberCount -value $collection.LocalMemberCount -MemberType NoteProperty
            $object| Add-Member -Name RefreshType -value $collection.RefreshType -MemberType NoteProperty
            $collections += $object
        }
        If ([string]::IsNullOrEmpty($export_CSV)){
            $collections | Out-GridView -Title "Collections With Incremental Update Enabled" 
        }
        Else{
            Write-Host "Exporting Report to - ""$export_CSV""" -ForegroundColor Yellow
            $collections | Export-Csv -Path $ExportCSV -NoTypeInformation 
        } 
    }
    #~ Reporting for Selection 2 (without)
    ElseIf($report -eq 'Y' -and $selection -eq '2'){
        Select_Save
        $collections = @()
        ForEach ($collection in $CollectionsWithoutIncrement){ 
            $object = New-Object -TypeName PSobject 
            $object| Add-Member -Name CollectionName -value $collection.Name -MemberType NoteProperty
            $object| Add-Member -Name CollectionID -value $collection.CollectionID -MemberType NoteProperty
            $object| Add-Member -Name MemberCount -value $collection.LocalMemberCount -MemberType NoteProperty
            $object| Add-Member -Name RefreshType -value $collection.RefreshType -MemberType NoteProperty
            $collections += $object
        }
        If ([string]::IsNullOrEmpty($export_CSV)){
            $collections | Out-GridView -Title "Collections Without Incremental Update Enabled" 
        }
        Else{
            Write-Host "Exporting Report to - ""$export_CSV""`n" -ForegroundColor Yellow
            $collections | Export-Csv -Path $export_CSV -NoTypeInformation 
        } 
    }
}

Function Disable_Incrementals{
    ForEach($target in $target_Collections){
        $DC = Get-CMDeviceCollection -Name $target
        If($DC.RefreshType -eq 4) {$DC.RefreshType = 1}
        If($DC.RefreshType -eq 6) {$DC.RefreshType = 2}
        $DC.Put() | Out-Null
    } 
}

#############################################################################################################################################
###~ Building Menu Functions ################################################################################################################
#############################################################################################################################################

Function Select_Menu{
    Do
    {
        Write-Host "`nEvaluating Device Collection Refresh Types...`n" -ForegroundColor Black -BackgroundColor Yellow
        Write-Host "Make your selection:" -ForegroundColor White -BackgroundColor Black

        Write-Host "1) " -ForegroundColor Yellow -BackgroundColor Black -NoNewLine 
        Write-Host "Display Incremental Count" -ForegroundColor White -BackgroundColor Black

        Write-Host "2) " -ForegroundColor Yellow -BackgroundColor Black -NoNewLine 
        Write-Host "Display Other Count" -ForegroundColor White -BackgroundColor Black

        Write-Host "3) " -ForegroundColor Yellow -BackgroundColor Black -NoNewLine 
        Write-Host "Disable Incrementals " -ForegroundColor White -BackgroundColor Black -NoNewline ; Write-Host "(Must modify scripts ""target_Collections"" array)" -ForegroundColor Cyan -BackgroundColor Black

        Write-Host "4) " -ForegroundColor Yellow -BackgroundColor Black -NoNewLine 
        Write-Host "Exit" -ForegroundColor White -BackgroundColor Black

        Write-Host "Your selectection " -ForegroundColor White -BackgroundColor Black -NoNewLine 
        Write-Host "(1, 2, 3, 4)"-ForegroundColor Yellow -BackgroundColor Black -NoNewLine
        Write-Host ":" -ForegroundColor White -BackgroundColor Black -NoNewLine
        $global:selection = Read-Host
        Clear-Host
        If ($selection -eq '1' -or $selection -eq '2' -or $selection -eq '3' -or $selection -eq '4'){
            $choice = "Valid"
        }
        Else{
            Write-Host "Invalid Selection" -ForegroundColor Red -BackgroundColor Yellow
            $choice = "Invalid"
        }
    }
    While ($choice -eq "Invalid" -or $choice -eq $null)
    Remove-Variable choice
}#~ Selection Menu for Choice

Function Select_1{
    Get_Incrementals
}#~ Selection: "Display Incremental Count"

Function Select_2{
    Get_Others
}#~ Selection: "Display Others Count"

Function Select_3{
    Disable_Incrementals
}#~ Selection: "Disable Incrementals on Modified Script Array"

Function Select_4{
    Break
}#~ Selection: "Exit"

Function Select_Report{
    Write-Host "Would you like a report" -ForegroundColor White -BackgroundColor Black -NoNewLine
    Write-Host " (Y/N)" -ForegroundColor Yellow -BackgroundColor Black -NoNewLine
    Write-Host "?:" -ForegroundColor White -BackgroundColor Black -NoNewLine 
    $global:report = Read-Host
    Clear-Host
}#Selection Menu for Report Run

Function Select_Run{
    Write-Host "Would you like to make another selection" -ForegroundColor White -BackgroundColor Black -NoNewLine
    Write-Host " (Y/N)" -ForegroundColor Yellow -BackgroundColor Black -NoNewLine
    Write-Host "?:" -ForegroundColor White -BackgroundColor Black -NoNewLine 
    $global:runAgain = Read-Host
    Clear-Host
}#~ Selection Menu for Script Run

Function Select_Save{
    [System.Reflection.Assembly]::LoadWithPartialName("System.windows.forms") | Out-Null 
    $SaveFileDialog = New-Object System.Windows.Forms.SaveFileDialog
    $SaveFileDialog.initialDirectory = "$env:UserProfile\Desktop"
    $SaveFileDialog.filter = “CSV files (*.csv)|*.csv|All files (*.*)|*.*”
    $SaveFileDialog.ShowDialog() | Out-Null
    $global:Export_CSV = $SaveFileDialog.filename
}#~ Selection Pop-Up for CSV Save

#############################################################################################################################################
###~ Present Menus ##########################################################################################################################
#############################################################################################################################################

Clear-Host

Do{
    #~ Selection Menu Prompt
    Select_Menu

    #~ Selection: "Display Incremental Count"
    If ($selection -eq '1'){
        Select_1
    }
    #~ Selection: "Display Others Count"
    ElseIf ($selection -eq '2'){
        Select_2
    }
    #~ Selection: "Remediate All Deployment Types"
    ElseIf ($selection -eq '3'){
        Select_3
    }
    #~ Selection: "Exit"
    ElseIf ($selection -eq "4"){
        Select_4
    }
    #~ Selection: "Invalid" (Should Catch in Menu)
    Else{
        Select_Menu
    }
   #~ Prompt for additional Run
    Select_Run

}While ($runAgain -eq 'Y')

#~ Return $true
Return $true

#############################################################################################################################################
#~ END ######################################################################################################################################
#############################################################################################################################################