function Select-Team {
  param(
    [string]$GroupId
  )

  if (!(Get-Module MicrosoftTeams)) {
    Write-Host "Installing module MicrosoftTeams"
    Install-Module -Name MicrosoftTeams
  }

  try {
    $Teams = Get-AssociatedTeam
  }
  catch {
    Write-Host "Connecting to your MS Teams account"
    Connect-MicrosoftTeams
    $Teams = Get-AssociatedTeam  
  }

  if ($GroupId) {
    Write-Host ("Selecting team with GroupId {0}" -f $GroupId)
    $Team = Get-Team -GroupId $GroupId
  }

  if (! $Team) {
    Write-Host "Select group in a new window"
    $Group = $Teams | Select-Object DisplayName, GroupId | Out-GridView -PassThru
    if (! $Group) {
      Write-Host "No team selected."
      return
    }
    $Team = Get-Team -GroupId $Group.GroupId
  }

  Write-Host ("Selected team: {0}" -f $Team.DisplayName)

  if (! $GroupId) {
    Write-Host ("Use -GroupId {0} to select it on the command line." -f $Team.GroupId)
  }

  return $Team
}

function Get-MembersFromExcel {
  param(
    [string]$ExcelFile
  )

  if (!(Get-Module ImportExcel)) {
    Write-Host "Installing module ImportExcel"
    Install-Module -Name ImportExcel
  }

  $List = Import-Excel $ExcelFile

  $Entries = @()

  foreach ($Row in $List) {
    $Entries += [PSCustomObject]@{
      Name   = $Row.Vorname + " " + $Row.Nachname
      Member = $Row."E-Mail-Adresse"
    }
  }

  $Entries
}

function Set-IndividualChannels {
  param(
    [Parameter(Mandatory = $true)]
    [string]$ExcelFile,
    [string]$GroupId
  )

  $Team = Select-Team -GroupId $GroupId

  Write-Host ("Setting up channels for {0} from {1}." -f $Team.DisplayName, $ExcelFile)

  $Members = Get-MembersFromExcel -ExcelFile $ExcelFile

  Write-Host ("Found {0} members." -f $Members.Length)

  $Channels = Get-TeamChannel -GroupId $Team.GroupId
  $TeamUsers = Get-TeamUser -GroupId $Team.GroupId

  foreach ($Member in $Members) {
    $DisplayName = $Member.Name.Substring(0, [System.Math]::Min(48, $Member.Name.Length))

    if (! ($TeamUsers | Where-Object User -EQ $Member.Member)) {
      Write-Host ("Adding {0} to team." -f $Member.Member)
      Add-TeamUser -GroupId $Team.GroupId -User $Member.Member
    }

    if (! ($Channels | Where-Object DisplayName -EQ $DisplayName) ) {
      Write-Host ("Creating channel {0}." -f $DisplayName)
      New-TeamChannel -GroupId $Team.GroupId -DisplayName $DisplayName -MembershipType Private
    }

    $ChannelUsers = Get-TeamChannelUser -GroupId $Team.GroupId -DisplayName $DisplayName

    if (! ($ChannelUsers | Where-Object User -EQ $Member.Member)) {
      Write-Host ("Adding {0} to {1}." -f $Member.Member, $DisplayName)
      Add-TeamChannelUser -GroupId $Team.GroupId -DisplayName $DisplayName -User $Member.Member
    }
  }
}

Set-IndividualChannels @args