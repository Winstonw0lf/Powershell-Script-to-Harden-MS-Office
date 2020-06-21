#Block Office applications from creating child processes
powershell.exe Add-MpPreference -AttackSurfaceReductionRules_Ids D4F940AB-401B-4EFC-AADC-AD5F3C50688A -AttackSurfaceReductionRules_Actions Enabled

#Block Office applications from injecting code into other processes
powershell.exe Add-MpPreference -AttackSurfaceReductionRules_Ids 75668C1F-73B5-4CF0-BB93-3ECF5CB7CC84 -AttackSurfaceReductionRules_Actions enable

#Block Win32 API calls from Office macro
powershell.exe Add-MpPreference -AttackSurfaceReductionRules_Ids 92E97FA1-2EDF-4476-BDD6-9DD0B4DDDC7B -AttackSurfaceReductionRules_Actions enable

#Block Office applications from creating executable content
powershell.exe Add-MpPreference -AttackSurfaceReductionRules_Ids 3B576869-A4EC-4529-8536-B80A7769E899 -AttackSurfaceReductionRules_Actions enable

#Block execution of potentially obfuscated scripts
powershell.exe Add-MpPreference -AttackSurfaceReductionRules_Ids 5BEB7EFE-FD9A-4556-801D-275E5FFC04CC -AttackSurfaceReductionRules_Actions Enabled

#Block executable content from email client and webmail
powershell.exe Add-MpPreference -AttackSurfaceReductionRules_Ids BE9BA2D9-53EA-4CDC-84E5-9B1EEEE46550 -AttackSurfaceReductionRules_Actions Enabled

#Block JavaScript or VBScript from launching downloaded executable content
powershell.exe Add-MpPreference -AttackSurfaceReductionRules_Ids D3E037E1-3EB8-44C8-A917-57927947596D -AttackSurfaceReductionRules_Actions Enabled


#Block only Office communication applications from creating child processes
powershell.exe Add-MpPreference -AttackSurfaceReductionRules_Ids 26190899-1602-49e8-8b27-eb1d0a1ce869 -AttackSurfaceReductionRules_Actions Enabled


#Blocks all ActiveX specifically for Office2016
Set-GPRegistryValue -Name "Test_gpo" -Key "HKCU\Software\Policies\Microsoft\office\common\security" -ValueName "disableallactivex" -Value 1 -Type DWord



#Block all Unmanaged Add-ins
Set-GPRegistryValue -Name "Test_gpo" -Key "HKCU\Software\Policies\Microsoft\office\16.0\word\resiliency" -ValueName "restricttolist" -Value 1 -Type DWord

Set-GPRegistryValue -Name "Test_gpo" -Key "HKCU\software\policies\microsoft\office\16.0\visio\resiliency" -ValueName "restricttolist" -Value 1 -Type DWord

Set-GPRegistryValue -Name "Test_gpo" -Key "HKCU\software\policies\microsoft\office\16.0\ms project\resiliency" -ValueName "restricttolist" -Value 1 -Type DWord

Set-GPRegistryValue -Name "Test_gpo" -Key "HKCU\software\policies\microsoft\office\16.0\excel\resiliency" -ValueName "restricttolist" -Value 1 -Type DWord

Set-GPRegistryValue -Name "Test_gpo" -Key "HKCU\software\policies\microsoft\office\16.0\powerpoint\resiliency" -ValueName "restricttolist" -Value 1 -Type DWord


#Force file extension to match file type
Set-GPRegistryValue -Name "Test_gpo" -Key "HKCU\software\policies\microsoft\office\16.0\excel\security" -ValueName "extensionhardening" -Value 2 -Type DWord


#Make hidden markup visible
Set-GPRegistryValue -Name "Test_gpo" -Key "HKCU\software\policies\microsoft\office\16.0\powerpoint\options" -ValueName "markupopensave" -Value 2 -Type DWord

Set-GPRegistryValue -Name "Test_gpo" -Key "HKCU\software\policies\microsoft\office\16.0\word\options" -ValueName "showmarkupopensave" -Value 1 -Type DWord


#Turn On File Validation

Set-GPRegistryValue -Name "Test_gpo" -Key "HKCU\software\policies\microsoft\office\16.0\excel\security\filevalidation" -ValueName "enableonload" -Value 1 -Type DWord

Set-GPRegistryValue -Name "Test_gpo" -Key "HKCU\software\policies\microsoft\office\16.0\powerpoint\security\filevalidation" -ValueName "enableonload" -Value 1 -Type DWord

Set-GPRegistryValue -Name "Test_gpo" -Key "HKCU\software\policies\microsoft\office\16.0\word\security\filevalidation" -ValueName "enableonload" -Value 1 -Type DWord


#Open files from the Internet zone in Protected View / Turn on Protected View for attachments opened from Outlook / Open files from the Internet zone in Protected View

Set-GPRegistryValue -Name "Test_gpo" -Key "HKCU\software\policies\microsoft\office\16.0\excel\security\ProtectedView" -ValueName "enabledatabasefileProtectedView" -Value 1 -Type DWord

Set-GPRegistryValue -Name "Test_gpo" -Key "HKCU\software\policies\microsoft\office\16.0\excel\security\ProtectedView" -ValueName "disableinternetfilesinpv" -Value 0 -Type DWord

Set-GPRegistryValue -Name "Test_gpo" -Key "HKCU\software\policies\microsoft\office\16.0\excel\security\ProtectedView" -ValueName "disableunsafelocationsinpv" -Value 0 -Type DWord

Set-GPRegistryValue -Name "Test_gpo" -Key "\HKCU\Software\Policies\Microsoft\office\16.0\excel\security\filevalidation" -ValueName "openinprotectedview" -Value 1 -Type DWord

Set-GPRegistryValue -Name "Test_gpo" -Key "\HKCU\Software\Policies\Microsoft\office\16.0\excel\security\filevalidation" -ValueName "disableeditfrompv" -Value 1 -Type DWord

Set-GPRegistryValue -Name "Test_gpo" -Key "HKCU\software\policies\microsoft\office\16.0\excel\security\ProtectedView" -ValueName "disableattachmentsinpv" -Value 0 -Type DWord

Set-GPRegistryValue -Name "Test_gpo" -Key "HKCU\software\policies\microsoft\office\16.0\powerpoint\security\ProtectedView" -ValueName "disableinternetfilesinpv" -Value 0 -Type DWord

Set-GPRegistryValue -Name "Test_gpo" -Key "HKCU\software\policies\microsoft\office\16.0\powerpoint\security\ProtectedView" -ValueName "disableunsafelocationsinpv" -Value 0 -Type DWord

Set-GPRegistryValue -Name "Test_gpo" -Key "HKCU\software\policies\microsoft\office\16.0\powerpoint\security\filevalidation" -ValueName "openinprotectedview" -Value 0 -Type DWord

Set-GPRegistryValue -Name "Test_gpo" -Key "HKCU\software\policies\microsoft\office\16.0\powerpoint\security\filevalidation" -ValueName "disableeditfrompv" -Value 1 -Type DWord

Set-GPRegistryValue -Name "Test_gpo" -Key "HKCU\software\policies\microsoft\office\16.0\word\security\ProtectedView" -ValueName "disableinternetfilesinpv" -Value 0 -Type DWord

Set-GPRegistryValue -Name "Test_gpo" -Key "HKCU\software\policies\microsoft\office\16.0\word\security\ProtectedView" -ValueName "disableunsafelocationsinpv" -Value 0 -Type DWord

Set-GPRegistryValue -Name "Test_gpo" -Key "HKCU\software\policies\microsoft\office\16.0\word\security\ProtectedView" -ValueName "disableattachmentsinpv" -Value 0 -Type DWord

Set-GPRegistryValue -Name "Test_gpo" -Key "HKCU\software\policies\microsoft\office\16.0\word\security\filevalidation" -ValueName "openinprotectedview" -Value 0 -Type DWord

Set-GPRegistryValue -Name "Test_gpo" -Key "HKCU\software\policies\microsoft\office\16.0\word\security\filevalidation" -ValueName "disableeditfrompv" -Value 1 -Type DWord


#Turn off Trusted Documents / on the network

Set-GPRegistryValue -Name "Test_gpo" -Key "HKCU\software\policies\microsoft\office\16.0\excel\security\trusted documents" -ValueName "disabletrusteddocuments" -Value 1 -Type DWord

Set-GPRegistryValue -Name "Test_gpo" -Key "HKCU\software\policies\microsoft\office\16.0\excel\security\trusted documents" -ValueName "disablenetworktrusteddocuments" -Value 1 -Type DWord

Set-GPRegistryValue -Name "Test_gpo" -Key "HKCU\software\policies\microsoft\office\16.0\powerpoint\security\trusted documents" -ValueName "disabletrusteddocuments" -Value 1 -Type DWord

Set-GPRegistryValue -Name "Test_gpo" -Key "HKCU\software\policies\microsoft\office\16.0\powerpoint\security\trusted documents" -ValueName "disablenetworktrusteddocuments" -Value 1 -Type DWord

Set-GPRegistryValue -Name "Test_gpo" -Key "HKCU\software\policies\microsoft\office\16.0\visio\security\trusted documents" -ValueName "disabletrusteddocuments" -Value 1 -Type DWord

Set-GPRegistryValue -Name "Test_gpo" -Key "HKCU\software\policies\microsoft\office\16.0\visio\security\trusted documents" -ValueName "disablenetworktrusteddocuments" -Value 1 -Type DWord

Set-GPRegistryValue -Name "Test_gpo" -Key "HKCU\software\policies\microsoft\office\16.0\word\security\trusted documents" -ValueName "disabletrusteddocuments" -Value 1 -Type DWord

Set-GPRegistryValue -Name "Test_gpo" -Key "HKCU\software\policies\microsoft\office\16.0\word\security\trusted documents" -ValueName "disablenetworktrusteddocuments" -Value 1 -Type DWord