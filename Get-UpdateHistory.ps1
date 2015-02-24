<#
    .Synopsis
    Lists all Windows updates that have been installed on the computer.
    
    .Description
    This script uses the Windows Update Agent API to list all of the Windows updates that have been installed on this computer, then outputs them as custom PowerShell objects for easy filtering and formatting.

    .Example
    PS N:\scripts\Get-UpdateHistory> $Updates = .\Get-UpdateHistory.ps1

    PS N:\scripts\Get-UpdateHistory> $Updates[0]

    Categories          : System.__ComObject
    ClientApplicationID : AutomaticUpdates
    Date                : 2/13/2015 6:10:11 PM
    Description         : A security issue has been identified in a Microsoft software product that could affect your system. You can help protect your system by installing this update from 
                          Microsoft. For a complete listing of the issues that are included in this update, see the associated Microsoft Knowledge Base article. After you install this update, 
                          you may have to restart your system.
    HResult             : 0
    Operation           : 1
    ResultCode          : 2
    ServerSelection     : 1
    ServiceID           : 
    SupportUrl          : http://support.microsoft.com
    Title               : Security Update for Internet Explorer 11 for Windows 7 for x64-based Systems (KB3034196)
    UninstallationNotes : This software update can be removed by selecting View installed updates in the Programs and Features Control Panel.
    UninstallationSteps : System.__ComObject
    UnmappedResultCode  : 0
    UpdateIdentity      : System.__ComObject

    This command captures all installed updates into the variable $Updates, which you can then index into to look at individual installed updates.

    .Example
    PS N:\scripts\Get-UpdateHistory> .\Get-UpdateHistory.ps1 | Where-Object {$_.Title -match 'KB982670'}

    Categories          : System.__ComObject
    ClientApplicationID : AutomaticUpdates
    Date                : 6/6/2014 9:24:26 PM
    Description         : The Microsoft .NET Framework 4 Client Profile provides a subset of features from the .NET Framework 4. The Client Profile is designed to run client applications and to 
                          enable the fastest possible deployment for Windows Presentation Foundation (WPF) and Windows Forms technology.
    HResult             : -2145124300
    Operation           : 1
    ResultCode          : 4
    ServerSelection     : 1
    ServiceID           : 
    SupportUrl          : http://support.microsoft.com
    Title               : Microsoft .NET Framework 4 Client Profile for Windows 7 x64-based Systems (KB982670)
    UninstallationNotes : This software update can be removed by selecting View installed updates in the Programs and Features Control Panel.
    UninstallationSteps : System.__ComObject
    UnmappedResultCode  : -2145099775
    UpdateIdentity      : System.__ComObject

    This command searches through installed updates to find the one with KB number 'KB982670'. If the update is installed, it returns details about it. If the update is not installed, it will return a null value.
#>
$Session = New-Object -ComObject Microsoft.Update.Session
$Searcher = $Session.CreateUpdateSearcher()
$HistoryCount = $Searcher.GetTotalHistoryCount()
$Updates = @()
for ($i = 0; $i -lt $HistoryCount; $i++) {
    $COMUpdate = $Searcher.QueryHistory($i,1)
    $Update = New-Object -TypeName PSCustomObject
    foreach($Property in ($COMUpdate | Get-Member -Type Property)) {
        Add-Member -MemberType NoteProperty -Name $Property.Name -Value ($COMUpdate | Select-Object -ExpandProperty $Property.Name) -InputObject $Update
	}
    $Updates += $Update
}
$Updates
