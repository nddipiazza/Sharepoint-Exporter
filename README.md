# Sharepoint-Exporter

Open the sln file in Visual Studio

## Add the sharepoint client DLLs

Copy 

	* Microsoft.SharePoint.Client.dll
	* Microsoft.SharePoint.Client.Runtime.dll
 
In Solution explorer, right click References -> Add reference -> Browser -> Add these two dll's

## Add the System.Web.Extensions reference

Again In Solution explorer, right click References -> Add reference -> Assemblies -> Framework -> Click the check next to `System.Web.Extensions`

# Running 

`USAGE: SpPrefetchIndexBuilder.exe [siteUrl] [outputDir] [domain] [username]`

* If you don't specify a siteURL, it will use localhost.

* If you don't specify an outputDir, it will use CWD.

* If you don't specify domain and username, it will use your current user.

* If you specify your own domain and username, you will be prompted for a password. But you can also specify environment variable SP_PWD with the password to avoid this.

# Output

It will create a folder with a GUID filename, and then fill it with each site by name.

SiteName
	SiteName
		SiteName
		...
