param(
  [Parameter(Mandatory = $true)]
  [string]$GroupId
)

Get-TeamAllChannel -GroupId $GroupId | Where-Object { $_.DisplayName -ne "General" } | ForEach-Object {
  $NewName = "{0:X8}.{1,-35}" -f (Get-Random), $_.DisplayName 
  Set-TeamChannel -GroupId $GroupId -CurrentDisplayName $_.DisplayName -NewDisplayName $NewName
  Remove-TeamChannel -GroupId $GroupId -DisplayName $NewName
}