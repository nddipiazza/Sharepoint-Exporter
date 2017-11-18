# Sharepoint-Exporter

This project will make an export of all of your sharepoint site. 

## How to work on the Project

Open the sln file in Visual Studio

### Add the sharepoint client DLLs

Download the Sharepoint Client Component SDK for the version of Sharepoint you are using:

https://www.microsoft.com/en-us/download/confirmation.aspx?id=35585 - SharePoint Server 2013 Client Components SDK 
 
In Solution explorer, right click References -> Add reference -> Browse -> Add the following DLL's as listed in `%PROGRAMFILES%\SharePoint Client Components\redist.txt`

```
All files in:
%ProgramFiles%\SharePoint Client Components\Assemblies
%ProgramFiles%\SharePoint Client Components\Scripts

The following files in %ProgramFiles%\Common Files\Microsoft Shared\Web Server Extensions\15\ISAPI:
Microsoft.Office.Client.Education.dll
Microsoft.Office.Client.Policy.dll
Microsoft.Office.Client.TranslationServices.dll
Microsoft.SharePoint.Client.dll
Microsoft.SharePoint.Client.DocumentManagement.dll
Microsoft.SharePoint.Client.Publishing.dll
Microsoft.SharePoint.Client.Runtime.dll
Microsoft.SharePoint.Client.Search.Applications.dll
Microsoft.SharePoint.Client.Search.dll
Microsoft.SharePoint.Client.Taxonomy.dll
Microsoft.SharePoint.Client.UserProfiles.dll
Microsoft.SharePoint.Client.WorkflowServices.dll
```
## How to Run the program 

`USAGE: SpPrefetchIndexBuilder.exe [siteUrl] [outputDir] [domain] [username] [password]`

* If you don't specify a siteURL, it will use localhost.

* If you don't specify an outputDir, it will use CWD.

* If you don't specify domain and username, it will use your current user.

* If you specify your own domain and username, you will be prompted for a password. But you can also specify environment variable `SP_PWD` with the password to avoid this.

* If you do not want to download files 

## What does it output?

Creates an output directory in the `outputDir` directory you specified in the cmd line arguments.

 * Metadata (including RoleAssignments and RoleDefinitions) for Webs, Lists and List Items
 * Files
 
```
 files
    |-> GUID.extension of each file
 lists
    |-> GUID.json representing each list and their list items
 web.json
```

*web.json* example format
```
{
	"Title": "Sharepoint",
	"Id": "634a49d6-40b5-4ac2-8e86-4ff4c9cb7833",
	"Description": "",
	"Url": "http://sphost",
	"LastItemModifiedDate": "11/14/2017 6:33:43 PM",
	"listsJsonPath": "c:\\outdir\\cfb9f363\\lists\\024a7730-d85e-42bc-889e-a4dc58290adb.json",
	"RoleDefinitions": {
		"1073741829": {
			"Id": 1073741829,
			"Name": "Full Control",
			"RoleTypeKind": "Administrator"
		},
		...
	},
	"RoleAssignments": {
		"Excel Services Viewers": {
			"LoginName": "Excel Services Viewers",
			"PrincipalType": "SharePointGroup",
			"RoleDefinitionIds": ["1073741924"]
		},
		...
	},
	"SubWebs": {
		"http://sphost/subsite": {
			"Title": "bradhays-team",
			"Id": "d48d1146-6de8-49a7-adb3-8ce24ed96778",
			"Description": "",
			"Url": "http://sphost/subsite",
			"LastItemModifiedDate": "11/14/2017 6:33:43 PM",
			"listsJsonPath": ...
			"RoleDefinitions": {
				
				...
			},
			"RoleAssignments": {
				
				...
```

# How to build on Linux

* Install Mono. http://www.mono-project.com/download/#download-lin

* Install MonoDevelop.

* Open the Solution file in Mono Develop.

* Copy the Sharepoint 2013 SDK DLLs into the `/path/to/Sharepoint-Exporter/SpCrawler` directory.

* In solution view, Right click on Reference -> Edit References -> .NET Assembly -> Browse -> Import the DLLs you put in `/path/to/Sharepoint-Exporter/SpCrawler`

* Ctrl+F8 Select the "Release" drop down from the build configuration and Ctrl+F8 to build the project.

* Now run this to build the Linux .bin file:

```
cd /path/to/Sharepoint-Exporter/SpCrawler/bin/Release
mkbundle -o SpPrefetchIndexBuilder.bin /path/to/Sharepoint-Exporter/SpCrawler/bin/Release/SpPrefetchIndexBuilder.exe -L /usr/lib/mono/4.6.1-api -L /home/ndipiazza/lucidworks/Sharepoint-Exporter/SpCrawler
```

At this point, `/path/to/Sharepoint-Exporter/SpCrawler/bin/Release/SpPrefetchIndexBuilder.bin http://yoursphost /your/output/path domain username` will work just like it had worked for Windows.
