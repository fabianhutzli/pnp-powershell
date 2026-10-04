---
Module Name: PnP.PowerShell
schema: 2.0.0
applicable: SharePoint Online
online version: https://pnp.github.io/powershell/cmdlets/Get-PnPFile.html
external help file: PnP.PowerShell.dll-Help.xml
title: Get-PnPFile
---
  
# Get-PnPFile

## SYNOPSIS
Downloads a file

## SYNTAX

### Return as file object (Default)
```powershell
Get-PnPFile -Url <String> -AsFileObject [-UseGraph] [-Connection <PnPConnection>]
```

### Return as list item
```powershell
Get-PnPFile -Url <String> -AsListItem [-ThrowExceptionIfFileNotFound] [-UseGraph] [-Connection <PnPConnection>] 
```

### Save to local path
```powershell
Get-PnPFile -Url <String> -AsFile -Path <String> -Filename <String> [-Force] [-UseGraph] [-Connection <PnPConnection>] 
```

### Return as string
```powershell
Get-PnPFile -Url <String> -AsString [-UseGraph] [-Connection <PnPConnection>] 
```

### Return as memorystream
```powershell
Get-PnPFile -Url <String> -AsMemoryStream [-UseGraph] [-Connection <PnPConnection>] 
```

## DESCRIPTION
Allows downloading of a file from SharePoint Online. The file contents can either be read directly into memory as text, directly saved to local disk or stored in memory for further processing.

With `-UseGraph` the file is retrieved through Microsoft Graph instead of the SharePoint client object model. The file is addressed by its URL alone through the Graph shares endpoint, so the site, web and list it lives in are never queried. This makes it usable from an application that has only been granted access to that specific list, list item or file through the `Lists.SelectedOperations.Selected`, `ListItems.SelectedOperations.Selected` or `Files.SelectedOperations.Selected` permissions, such as one set up with `Grant-PnPEntraIDAppListPermission`, `Grant-PnPEntraIDAppListItemPermission` or `Grant-PnPEntraIDAppFilePermission`. With `-UseGraph`, `-AsFileObject` returns the Microsoft Graph driveItem and `-AsListItem` returns the Microsoft Graph listItem with its field values, rather than their SharePoint client object model counterparts.

## EXAMPLES

### EXAMPLE 1
```powershell
Get-PnPFile -Url "/sites/project/Shared Documents/Document.docx"
```

Retrieves the file and downloads it to the current folder

### EXAMPLE 2
```powershell
Get-PnPFile -Url "https://contoso.sharepoint.com/sites/project/Shared Documents/Document.docx"
```

Retrieves the file and downloads it to the current folder

### EXAMPLE 3
```powershell
Get-PnPFile -Url /sites/project/SiteAssets/image.jpg -Path c:\temp -FileName image.jpg -AsFile
```

Retrieves the file and downloads it to c:\temp\image.jpg

### EXAMPLE 4
```powershell
Get-PnPFile -Url /sites/project/_catalogs/themes/15/company.spcolor -AsString
```

Retrieves the contents of the file as text and outputs its contents to the console

### EXAMPLE 5
```powershell
Get-PnPFile -Url /sites/project/Shared Documents/Folder/Presentation.pptx -AsFileObject
```

Retrieves the file and returns it as a File object

### EXAMPLE 6
```powershell
Get-PnPFile -Url /sites/project/_catalogs/themes/15/company.spcolor -AsListItem
```

Retrieves the file and returns it as a ListItem object

### EXAMPLE 7
```powershell
Get-PnPFile -Url /personal/john_tenant_onmicrosoft_com/Documents/Sample.xlsx -Path c:\temp -FileName Project.xlsx -AsFile
```

Retrieves the file Sample.xlsx by its site relative URL from a OneDrive for Business site and downloads it to c:\temp\Project.xlsx

### EXAMPLE 8
```powershell
Get-PnPFile -Url "/sites/templates/Shared Documents/HR Site.pnp" -AsMemoryStream
```

Retrieves the file in memory for further processing

### EXAMPLE 9
```powershell
Connect-PnPOnline -Url "https://contoso.sharepoint.com/sites/project" -ClientId $appId -Tenant contoso.onmicrosoft.com -Thumbprint $thumbprint
Get-PnPFile -Url "/sites/project/Shared Documents/Contracts/Agreement.docx" -UseGraph
```

Retrieves the metadata of the file through Microsoft Graph. Use this to verify that an application which has only been granted access to this file, its list item or its library through a selected permission can read it.

### EXAMPLE 10
```powershell
Get-PnPFile -Url "/sites/project/Shared Documents/Contracts/Agreement.docx" -AsFile -Path c:\temp -Filename Agreement.docx -UseGraph
```

Downloads the file through Microsoft Graph to c:\temp\Agreement.docx

## PARAMETERS

### -AsFile

```yaml
Type: SwitchParameter
Parameter Sets: Save to local path

Required: True
Position: Named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -AsFileObject
Retrieve the file contents as a file object.

```yaml
Type: SwitchParameter
Parameter Sets: Return as file object

Required: False
Position: Named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -AsListItem
Returns the file as a listitem showing all its properties

```yaml
Type: SwitchParameter
Parameter Sets: Return as list item

Required: False
Position: Named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -AsString
Retrieve the file contents as a string

```yaml
Type: SwitchParameter
Parameter Sets: Return as string

Required: False
Position: Named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -AsMemoryStream

```yaml
Type: SwitchParameter
Parameter Sets: Download the content of the file to memory

Required: True
Position: Named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Connection
Optional connection to be used by the cmdlet. Retrieve the value for this parameter by either specifying -ReturnConnection on Connect-PnPOnline or by executing Get-PnPConnection.

```yaml
Type: PnPConnection
Parameter Sets: (All)

Required: False
Position: Named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Filename
Name for the local file

```yaml
Type: String
Parameter Sets: Save to local path

Required: False
Position: Named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Force
Overwrites the file if it exists.

```yaml
Type: SwitchParameter
Parameter Sets: Save to local path

Required: False
Position: Named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Path
Local path where the file should be saved

```yaml
Type: String
Parameter Sets: Save to local path

Required: False
Position: Named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -ThrowExceptionIfFileNotFound
If provided in combination with -AsListItem, a System.ArgumentException will be thrown if the file specified in the -Url argument does not exist. Otherwise it will return nothing instead.

```yaml
Type: SwitchParameter
Parameter Sets: Return as list item

Required: False
Position: Named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Url
The URL (server or site relative) to the file. A sequence such as `%20` in the URL is taken literally when a file of that name exists, and is decoded otherwise.

```yaml
Type: String
Parameter Sets: (All)
Aliases: ServerRelativeUrl, SiteRelativeUrl

Required: True
Position: 0
Default value: None
Accept pipeline input: True (ByValue)
Accept wildcard characters: False
```

### -UseGraph
Retrieves the file through Microsoft Graph instead of the SharePoint client object model, without querying the site, web or list the file lives in. Requires a Microsoft Graph permission granting read access to the file, such as Files.Read.All, Sites.Read.All, or one of the selected permissions Sites.Selected, Lists.SelectedOperations.Selected, ListItems.SelectedOperations.Selected or Files.SelectedOperations.Selected combined with a grant on the site, list, list item or file.

```yaml
Type: SwitchParameter
Parameter Sets: (All)

Required: False
Position: Named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

## RELATED LINKS

[Microsoft 365 Patterns and Practices](https://aka.ms/m365pnp)