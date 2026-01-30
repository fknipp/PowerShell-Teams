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

function Get-ExcelGroupMembers {
  param(
    [string]$ExcelFile
  )

  if (!(Get-Module ImportExcel)) {
    Write-Host "Installing module ImportExcel"
    Install-Module -Name ImportExcel
  }

  $List = Import-Excel $ExcelFile

  $Groups = @()
  
  foreach ($Item in $List) {
    foreach ($Group in $Item.Gruppen -split ", ") {
      if ($Group -and $Group -notmatch '^[A-Z][A-Z][A-Z][A-Z]-[0-9]' -and ! $Groups.Where{ $_ -eq $Group }) {
        $Groups += $Group
      }
    }
  }

  $Entries = @()

  foreach ($Group in $Groups) {
    $Entries += [PSCustomObject]@{
      GroupName = $Group
      Members   = $List.Where{ $Group -in ($_.Gruppen -split ", ") } | ForEach-Object { $_."E-Mail-Adresse" }
    }
  }

  $Entries
}

function Set-GroupChannels {
  param(
    [Parameter(Mandatory = $true)]
    [string]$ExcelFile,
    [string]$GroupId,
    [switch]$Owner
  )

  $Team = Select-Team -GroupId $GroupId

  Write-Host ("Setting up channels for {0} from {1}." -f $Team.DisplayName, $ExcelFile)

  $GroupMembers = Get-ExcelGroupMembers -ExcelFile $ExcelFile

  Write-Host ("Found {0} groups." -f $GroupMembers.Length)

  $Channels = Get-TeamChannel -GroupId $Team.GroupId

  foreach ($Group in $GroupMembers) {
    $DisplayName = $Group.GroupName.Substring(0, [System.Math]::Min(48, $Group.GroupName.Length))

    if (! ($Channels | Where-Object DisplayName -EQ $DisplayName) ) {
      Write-Host ("Creating channel {0}." -f $DisplayName)
      New-TeamChannel -GroupId $Team.GroupId -DisplayName $DisplayName -MembershipType Shared
    }

    $ChannelUsers = Get-TeamChannelUser -GroupId $Team.GroupId -DisplayName $DisplayName

    foreach ($Member in $Group.Members) {
      if (! ($ChannelUsers | Where-Object User -EQ $Member)) {
        if ($Owner) {
          Write-Host ("Adding {0} to {1} as owner." -f $Member, $DisplayName)
          Add-TeamChannelUser -GroupId $Team.GroupId -DisplayName $DisplayName -User $Member
          Add-TeamChannelUser -GroupId $Team.GroupId -DisplayName $DisplayName -User $Member -Role Owner          
        }
        else {
          Write-Host ("Adding {0} to {1}." -f $Member, $DisplayName)
          Add-TeamChannelUser -GroupId $Team.GroupId -DisplayName $DisplayName -User $Member
        }
      }
      else {
        Write-Host ("{0} is already in {1}." -f $Member, $DisplayName)
      }
    }
  }
}

Set-GroupChannels @args