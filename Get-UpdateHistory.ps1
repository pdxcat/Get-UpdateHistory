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
