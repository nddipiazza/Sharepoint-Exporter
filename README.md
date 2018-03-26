# Sharepoint-Exporter

This project will make an export of all of your sharepoint site. 

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

* Install packages `monodevelop mono-devel mono-complete monodevelop-versioncontrol` following instructions for your particular distro here: http://www.mono-project.com/download/#download-lin

Do not use the default packages from your OS' repository using `apt-get`. Make sure to get it using the mono-project.com repository as specified in the instructions.

Does not work for mono version 4.x. Tested only on Mono JIT compiler version 5.10.1.4.

* To make code changes in the monodevelop IDE, install it with: `flatpak install --user --from https://download.mono-project.com/repo/monodevelop.flatpakref` and run it with `/usr/bin/flatpak run --branch=stable --arch=x86_64 --command=monodevelop com.xamarin.MonoDevelop %F`

* Build using: `build.sh`

* The bundled binary is at `SpCrawler/bin/Release/SpFetcherBundled`
