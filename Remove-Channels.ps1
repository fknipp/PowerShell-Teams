param(
  [Parameter(Mandatory = $true)]
  [string]$GroupId
)

Get-TeamAllChannel -GroupId $GroupId | Where-Object { $_.DisplayName -ne "General" } | ForEach-Object {
  $NewName = ("{0:X8}.{1,-35}" -f (Get-Random), $_.DisplayName).Trim()
  Write-Host ("Renaming channel {0} to {1}." -f $_.DisplayName, $NewName)
  Set-TeamChannel -GroupId $GroupId -CurrentDisplayName $_.DisplayName -NewDisplayName $NewName
  Write-Host ("Removing channel {0}." -f $NewName)
  Remove-TeamChannel -GroupId $GroupId -DisplayName $NewName
}
