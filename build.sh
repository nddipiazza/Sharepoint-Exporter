#!/bin/bash
nuget restore
xbuild /p:TargetFrameworkVersion="v4.5" /p:Configuration=Release ./SpPrefetchIndexBuilder.sln
mkbundle --simple --static --deps \
	-o SpCrawler/bin/Release/SpFetcherBundled \
	--config $MONO_INSTALLATION/etc/mono/config \
	--machine-config $MONO_INSTALLATION/etc/mono/4.5/machine.config \
	-L SpCrawler/bin/Release \
	SpCrawler/bin/Release/SpPrefetchIndexBuilder.exe \
	./SpCrawler/bin/Release/Microsoft.Office.Client.Policy.dll \
	./SpCrawler/bin/Release/Microsoft.Office.Client.TranslationServices.dll \
	./SpCrawler/bin/Release/Microsoft.Online.SharePoint.Client.Tenant.dll \
	./SpCrawler/bin/Release/Microsoft.ProjectServer.Client.dll \
	./SpCrawler/bin/Release/Microsoft.SharePoint.Client.dll \
	./SpCrawler/bin/Release/Microsoft.SharePoint.Client.DocumentManagement.dll \
	./SpCrawler/bin/Release/Microsoft.SharePoint.Client.Publishing.dll \
	./SpCrawler/bin/Release/Microsoft.SharePoint.Client.Runtime.dll \
	./SpCrawler/bin/Release/Microsoft.SharePoint.Client.Search.Applications.dll \
	./SpCrawler/bin/Release/Microsoft.SharePoint.Client.Search.dll \
	./SpCrawler/bin/Release/Microsoft.SharePoint.Client.Taxonomy.dll \
	./SpCrawler/bin/Release/Microsoft.SharePoint.Client.UserProfiles.dll \
	./SpCrawler/bin/Release/Microsoft.SharePoint.Client.WorkflowServices.dll \
	./SpCrawler/bin/Release/log4net.dll 
