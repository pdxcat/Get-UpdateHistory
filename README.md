# Get-UpdateHistory
List installed Windows Updates.

The built-in PowerShell Get-Hotfix cmdlet does not return *all* installed Windows Updates on a system, nor does the Win32_QuickFixEngineering WMI class. This script uses the Windows Update Agent API to return the list of all installed Windows updates as a set of custom PowerShell objects.

Many thanks to the Hey, Scripting Guy! Blog for the instructions on how to get this started: http://blogs.technet.com/b/heyscriptingguy/archive/2009/03/09/how-can-i-list-all-updates-that-have-been-added-to-a-computer.aspx
