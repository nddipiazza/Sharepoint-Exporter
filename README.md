# Sharepoint-Exporter

This project will make an export of all of your sharepoint site in a hierarchical metadata file folder format.

## How to work on the Project

Open the sln file in Visual Studio

### Add the sharepoint client DLLs

Download the Sharepoint Client Component SDK for the version of Sharepoint you are using:

https://www.microsoft.com/en-us/download/confirmation.aspx?id=35585 - SharePoint Server 2013 Client Components SDK 

https://www.microsoft.com/en-us/download/details.aspx?id=51679 - SharePoint Server 2016 Client Components SDK 
 
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

### Add the System.Web.Extensions reference

Again In Solution explorer, right click References -> Add reference -> Assemblies -> Framework -> Click the check next to `System.Web.Extensions`

## How to Run the program 

`USAGE: SpPrefetchIndexBuilder.exe [siteUrl] [outputDir] [domain] [username]`

* If you don't specify a siteURL, it will use localhost.

* If you don't specify an outputDir, it will use CWD.

* If you don't specify domain and username, it will use your current user.

* If you specify your own domain and username, you will be prompted for a password. But you can also specify environment variable SP_PWD with the password to avoid this.

## Output

It will create a folder with a GUID filename, and then fill it with each site by name.

SiteName
	SiteName
		SiteName
		...
