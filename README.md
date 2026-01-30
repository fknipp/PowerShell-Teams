# PowerShell-Teams

This repository contains script to simplify the setup of Microsoft Teams for
project groups. I use these scripts to automate the setup of my lectures at
[UAS Burgenland](https://www.hochschule-burgenland.at).

## Set-GroupChannels

```
Set-GroupChannels.ps1 -ExcelFile <Participants as XLSX from Moodle> -GroupId <MS Teams Group Id> [-Owner]
```

Sets up the groups and their members from the given Excel file. If the GroupId
is unknown, a list of availabe groups is shown to select the group.

The Excel file has to be exported from the participants page in Moodle.

Use -Owner to add the members with ownership privilege on the channel.

Known limitations:

- The language setting of Moodle must be German, as the column titles are
  expected in German language (Gruppen, E-Mail-Adresse).
- Group names starting with [A-Z][A-Z][A-Z][A-Z]-[0-9] are filtered, as these
  are automatically generated for the respective cohorts (e.g. BSWE-4).

## Set-IndividualChannels

```
Set-IndividualChannels.ps1 -ExcelFile <Participants as XLSX from Moodle> -GroupId <MS Teams Group Id>
```

Sets up a channel for every row from the given Excel file. If the GroupId is
unknown, a list of availabe groups is shown to select the group.

The Excel file has to be exported from the participants page in Moodle.

Known limitations:

- The language setting of Moodle must be German, as the column titles are
  expected in German language (Vorname, Nachname, E-Mail-Adresse).

## Remove-Channels

```
Remove-Channels.ps1 -GroupId <MS Teams Group Id>
```

Removes all channels from the given group after renaming them. The groups will
be permanently removed after 21 days. Renaming the groups allows the recreation
of groups with the former name.
