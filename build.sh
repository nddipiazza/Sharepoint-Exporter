nuget restore
xbuild /p:TargetFrameworkVersion="v4.5" /p:Configuration=Release ./SpPrefetchIndexBuilder.sln
mkbundle --simple --static --deps -o SpCrawler/bin/Release/SpFetcherBundled --config /etc/mono/config --machine-config /etc/mono/4.5/machine.config -L SpCrawler/bin/Release SpCrawler/bin/Release/SpPrefetchIndexBuilder.exe SpCrawler/bin/Release/*.dll
