#############################################################################################################################################
###~ OSD++ - Required Application Compliance - Stand Alone Deskside Tool ####################################################################
#
#~ Brian Tancredi
#~ Created: 2017-06-07
#~ Modified: 2017-06-08
#
#~ References:
#~ https://david-obrien.net/2013/12/find-required-deployments-configmgr-clients/
#
#############################################################################################################################################

#############################################################################################################################################
###~ Assigning Variables and Declaring Functions  ###########################################################################################
#############################################################################################################################################

#~ Required applications to ignore if any
$ignore = "SCCM Admin Console"

#~ Gather all Required ($_.ResolvedState -eq "Installed") applications that are not installed ($_.InstallState -ne "Installed")
$required_NotInstalled = Get-WmiObject -Class ccm_application -Namespace root\ccm\clientsdk `
    | Where-Object {($_.IsMachineTarget) -and ($_.InstallState -ne "Installed") -and ($_.ResolvedState -eq "Installed") -and ($_.Name -ne $ignore)} `
    | Sort Name | Select -ExpandProperty Name

#~ Gather all Required ($_.ResolvedState -eq "Installed") applications that are installed ($_.InstallState -eq "Installed")
$required_Installed = Get-WmiObject -Class ccm_application -Namespace root\ccm\clientsdk `
    | Where-Object {($_.IsMachineTarget) -and ($_.InstallState -eq "Installed") -and ($_.ResolvedState -eq "Installed")} `
    | Sort Name | Select -ExpandProperty Name

Function Display-Installed{
    #~ Create List of Required Installed Applications
    ForEach ($application in $required_Installed){$apps += ">$application`n"}
    $app_List = $apps
    Remove-Variable apps

    #~ Display completion of installed apps
    $wshell = New-Object -ComObject Wscript.Shell
    $wshell.Popup("Required Installs Completed:`n`n$app_List",0,"OSD++",0x40)
}#~ Pop-up on Compliant

Function Display-Uninstalled{
    #~ Create List of Required Installed Applications
    ForEach ($application in $required_NotInstalled){$apps += ">$application`n"}
    $app_List = $apps
    Remove-Variable apps

    #~ Display completion of installed apps
    $wshell = New-Object -ComObject Wscript.Shell
    $wshell.Popup("Required Installs Missing:`n`n$app_List",0,"OSD++",0x30)
}
#############################################################################################################################################
###~ Checking Compliance ####################################################################################################################
#############################################################################################################################################

#~ Check if any required apps are not installed
If ($required_NotInstalled){$compliant = $false;$fgc = "Red"}
Else{$compliant = $true;$fgc = "Green"}

#~ Write Host for Detection
Clear-Host
Write-Host "Compliance = $compliant" -ForegroundColor $fgc

#~ Notify End User
If($compliant){Display-Installed}
Else{Display-Uninstalled}

#############################################################################################################################################
#~ END ######################################################################################################################################
#############################################################################################################################################