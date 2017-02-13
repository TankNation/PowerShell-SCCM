﻿#############################################################################################################################################
###~ AllUser Desktop Shortcut Cleanup - Uninstall ###########################################################################################
#
#~ Brian Tancredi
#~ Created: 2016-11-23
#~ Modified:
#
#~ References:
#~ 
#
#############################################################################################################################################

#############################################################################################################################################
###~ Public Folder Shortcuts (.lnk) Prompted Removal - Uninstall ############################################################################
#############################################################################################################################################

#~ Declare Variables
$tar_Dir = "$env:PUBLIC\Desktop"
$dm_Dir = "Delete_Me"
$dm_Path = "$tar_dir\$dm_Dir"

#~ Remove File for Detection Method
Remove-Item -Path $dm_Path -Force -Recurse

#############################################################################################################################################
#~ END ######################################################################################################################################
#############################################################################################################################################