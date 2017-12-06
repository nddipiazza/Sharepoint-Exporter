using System;
using System.Net;
using System.Diagnostics;
using System.Security;
using System.Collections.Generic;
using System.Web.Script.Serialization;
using Microsoft.SharePoint.Client;
using System.Collections.Concurrent;
using System.Threading.Tasks;
using System.Threading;
using System.Net.Http;
using System.Xml;
using System.IO;


namespace SpPrefetchIndexBuilder
{
    class WebToFetch
    {
        public String url;
        public Dictionary<string, object> webDict;
        public bool isTopLevel;
    }

    class ListToFetch
    {
        public String site;
        public Guid listId;
        public Dictionary<string, object> listsDict = new Dictionary<string, object>();
    }

    class ListsOutput
    {
        public String jsonPath;
        public Dictionary<string, object> listsDict;
    }

    class SpPrefetchIndexBuilder
    {
        public static string baseDir = null;
        public static void CheckAbort()
        {
            if (System.IO.File.Exists(baseDir + Path.DirectorySeparatorChar + ".." + Path.DirectorySeparatorChar + ".doabort"))
            {
                Console.WriteLine("The .doabort file was found. Stopping program");
                Environment.Exit(0);
            }
        }
        public static HttpClient client;
        public string defaultSite = "http://localhost/";
        public CredentialCache cc = null;
        public string site;
        public static string topParentSite;
        public JavaScriptSerializer serializer = new JavaScriptSerializer();
        public int maxFileSizeBytes = -1;
        public static int numThreads = 50;
        public static bool onlyWebs = false;
        public static bool excludeRoleDefinitions = false;
        public static bool excludeRoleAssignments = false;
        public static bool deleteExistingOutputDir = false;

        public BlockingCollection<ListToFetch> listFetchBlockingCollection = new BlockingCollection<ListToFetch>();
        public BlockingCollection<WebToFetch> webFetchBlockingCollection = new BlockingCollection<WebToFetch>();
        public BlockingCollection<FileToDownload> fileDownloadBlockingCollection = new BlockingCollection<FileToDownload>();
        public Dictionary<string, object> rootWebDict;
        public List<ListsOutput> listsOutput = new List<ListsOutput>();

        public List<string> ignoreSiteNames = new List<string>();

        public SpPrefetchIndexBuilder(String[] args)
        {
            ignoreSiteNames.Add("Cache Profiles");
            ignoreSiteNames.Add("Content and Structure Reports");
            ignoreSiteNames.Add("Content Organizer Rules");
            ignoreSiteNames.Add("Content type publishing error log");
            ignoreSiteNames.Add("Converted Forms");
            ignoreSiteNames.Add("Device Channels");
            ignoreSiteNames.Add("Drop Off Library");
            ignoreSiteNames.Add("Form Templates");
            ignoreSiteNames.Add("Hold Reports");
            ignoreSiteNames.Add("Holds");
            ignoreSiteNames.Add("Long Running Operation Status");
            ignoreSiteNames.Add("MicroFeed");
            ignoreSiteNames.Add("Notification List");
            ignoreSiteNames.Add("Project Policy Item List");
            ignoreSiteNames.Add("Quick Deploy Items");
            ignoreSiteNames.Add("Relationships List");
            ignoreSiteNames.Add("Reusable Content");
            ignoreSiteNames.Add("Site Collection Documents");
            ignoreSiteNames.Add("Site Collection Images");
            ignoreSiteNames.Add("Solution Gallery");
            ignoreSiteNames.Add("Style Library");
            ignoreSiteNames.Add("Submitted E-mail Records");
            ignoreSiteNames.Add("Suggested Content Browser Locations");
            ignoreSiteNames.Add("TaxonomyHiddenList");
            ignoreSiteNames.Add("Theme Gallery");
            ignoreSiteNames.Add("Translation Packages");
            ignoreSiteNames.Add("Translation Status");
            ignoreSiteNames.Add("User Information List");
            ignoreSiteNames.Add("Variation Labels");
            ignoreSiteNames.Add("Web Part Gallery");
            ignoreSiteNames.Add("wfpub");
            ignoreSiteNames.Add("Composed Looks");
            ignoreSiteNames.Add("Master Page Gallery");
            ignoreSiteNames.Add("Site Assets");
            ignoreSiteNames.Add("Site Pages");

            string spMaxFileSizeBytes = Environment.GetEnvironmentVariable("SP_MAX_FILE_SIZE_BYTES");
            if (spMaxFileSizeBytes != null)
            {
                maxFileSizeBytes = int.Parse(spMaxFileSizeBytes);
            }
            string spNumThreads = Environment.GetEnvironmentVariable("SP_NUM_THREADS");
            if (spNumThreads != null)
            {
                numThreads = int.Parse(spNumThreads);
            }
            serializer.MaxJsonLength = 1677721600;

            site = defaultSite;

            bool help = false;

            string spDomain = null;
            string spUsername = null;
            string spPassword = Environment.GetEnvironmentVariable("SP_PWD");
            baseDir = Directory.GetCurrentDirectory();
            bool customBaseDir = false;

            foreach (string arg in args)
            {
                if (arg.Equals("--help") || arg.Equals("-help") || arg.Equals("/help"))
                {
                    help = true;
                    break;
                }
                else if (arg.StartsWith("--siteUrl="))
                {
                    site = arg.Split(new Char[] { '=' })[1];
                }
                else if (arg.StartsWith("--outputDir="))
                {
                    baseDir = arg.Split(new Char[] { '=' })[1];
                    customBaseDir = true;
                }
                else if (arg.StartsWith("--domain="))
                {
                    spDomain = arg.Split(new Char[] { '=' })[1];
                }
                else if (arg.StartsWith("--username="))
                {
                    spUsername = arg.Split(new Char[] { '=' })[1];
                }
                else if (arg.StartsWith("--password="))
                {
                    spPassword = arg.Split(new Char[] { '=' })[1];
                }
                else if (arg.StartsWith("--numThreads="))
                {
                    numThreads = int.Parse(arg.Split(new Char[] { '=' })[1]);
                }
                else if (arg.StartsWith("--maxFileSizeBytes="))
                {
                    maxFileSizeBytes = int.Parse(arg.Split(new Char[] { '=' })[1]);
                }
                else if (arg.StartsWith("--onlyWebs="))
                {
                    onlyWebs = Boolean.Parse(arg.Split(new Char[] { '=' })[1]);
                }
                else if (arg.StartsWith("--excludeRoleAssignments="))
                {
                    excludeRoleAssignments = Boolean.Parse(arg.Split(new Char[] { '=' })[1]);
                }
                else if (arg.StartsWith("--excludeRoleDefinitions="))
                {
                    excludeRoleDefinitions = Boolean.Parse(arg.Split(new Char[] { '=' })[1]);
                }
                else if (arg.StartsWith("--deleteExistingOutputDir="))
                {
                    deleteExistingOutputDir = Boolean.Parse(arg.Split(new Char[] { '=' })[1]);
                }
                else
                {
                    help = true;
                }
            }

            if (help)
            {
                Console.WriteLine("USAGE: SpPrefetchIndexBuilder.exe --siteUrl=siteUrl --outputDir=[outputDir] --domain=[domain] --username=[username] --password=[password (not recommended, do not specify to be prompted or use SP_PWD environment variable)] --numThreads=[optional number of threads to use while fetching] --maxFileSizeBytes=[optional maximum file size]");
                Environment.Exit(0);
            }

            if (customBaseDir && deleteExistingOutputDir && Directory.Exists(baseDir))
            {
                deleteDirectory(baseDir);
            }
            Directory.CreateDirectory(baseDir);
            if (!onlyWebs)
            {
                Directory.CreateDirectory(baseDir + Path.DirectorySeparatorChar + "lists");
                Directory.CreateDirectory(baseDir + Path.DirectorySeparatorChar + "files");
            }
            if (site.EndsWith("/"))
            {
                site = site.Substring(0, site.Length - 1);
            }
            cc = new CredentialCache();
            NetworkCredential nc;
            if (spPassword == null)
            {
                Console.WriteLine("Please enter password for {0}", spUsername);
                nc = new NetworkCredential(spUsername, GetPassword(), spDomain);
            }
            else
            {
                nc = new NetworkCredential(spUsername, spPassword, spDomain);
            }
            cc.Add(new Uri(site), "NTLM", nc);
            HttpClientHandler handler = new HttpClientHandler();
            handler.Credentials = cc;
            client = new HttpClient(handler);
            client.Timeout = TimeSpan.FromMinutes(4);
        }


        static void Main(string[] args)
        {
            try
            {
                Stopwatch sw = Stopwatch.StartNew();
                SpPrefetchIndexBuilder spib = new SpPrefetchIndexBuilder(args);

                List<string> topParentSites = spib.GetAllTopLevelSites();

                Console.WriteLine("Starting export of top level sites {0}", string.Join(", ", topParentSites));

                topParentSite = spib.site;

                foreach (String nextTopParentSite in topParentSites)
                {
                    Stopwatch swWeb = Stopwatch.StartNew();
                    spib.getSubWebs(topParentSite, null);
                    Parallel.For(0, numThreads, x => spib.FetchWeb());
                    spib.writeWebJson();
                    Console.WriteLine("Web fetch of {0} complete. Took {1} milliseconds.", topParentSite, swWeb.ElapsedMilliseconds);

                    if (!onlyWebs)
                    {
                        Stopwatch swLists = Stopwatch.StartNew();
                        Parallel.For(0, numThreads, x => spib.FetchList());
                        spib.writeAllListsToJson();
                        Console.WriteLine("Lists metadata dump of {0} complete. Took {1} milliseconds.", topParentSite, swLists.ElapsedMilliseconds);
                        Console.WriteLine("Downloading the files recieved during the index building");
                        Parallel.For(0, numThreads, x => spib.DownloadFilesFromQueue());
                    }
                }
                Console.WriteLine("Export complete. Took {0} milliseconds.", sw.ElapsedMilliseconds);
            }
            catch (Exception anyException)
            {
                Console.WriteLine("Prefetch index building failed for {0}: {1}", string.Join(" ", args), anyException.Message);
                Console.WriteLine(anyException.StackTrace);
                Environment.Exit(1);
            }
        }


        public void DownloadFilesFromQueue()
        {
            //Console.WriteLine("Starting Thread {0}", Thread.CurrentThread.ManagedThreadId);
            FileDownloader.DownloadFiles(fileDownloadBlockingCollection, 240000, client);
        }

        public void FetchWeb()
        {
            //Console.WriteLine("Starting Thread {0}", Thread.CurrentThread.ManagedThreadId);
            WebToFetch webToFetch;
            while (webFetchBlockingCollection.TryTake(out webToFetch))
            {
                CheckAbort();
                string url = webToFetch.url;
                Console.WriteLine("Thread {0} exporting site {1}", Thread.CurrentThread.ManagedThreadId, url);
                ClientContext clientContext = getClientContext(url);
                Web web = clientContext.Web;
                if (excludeRoleDefinitions && excludeRoleDefinitions)
                {
                    clientContext.Load(web, website => website.Webs, website => website.Title, website => website.Url, website => website.Description, website => website.Id, website => website.LastItemModifiedDate);
                }
                else
                {
                    clientContext.Load(web, website => website.Webs, website => website.Title, website => website.Url, website => website.RoleDefinitions, website => website.RoleAssignments, website => website.HasUniqueRoleAssignments, website => website.Description, website => website.Id, website => website.LastItemModifiedDate);
                }
                try
                {
                    clientContext.ExecuteQuery();
                }
                catch (Exception ex)
                {
                    Console.WriteLine("Could not load site {0} because of Error {1}", url, ex.Message);
                    continue;
                }
                string listsFileName = Guid.NewGuid().ToString() + ".json";
                string listsJsonPath = baseDir + Path.DirectorySeparatorChar + "lists" + Path.DirectorySeparatorChar + listsFileName;
                Dictionary<string, object> webDict = webToFetch.webDict;
                webDict.Add("Title", web.Title);
                webDict.Add("Id", web.Id);
                webDict.Add("Description", web.Description);
                webDict.Add("Url", url);
                webDict.Add("LastItemModifiedDate", web.LastItemModifiedDate.ToString());
                if (!onlyWebs)
                {
                    webDict.Add("ListsFileName", listsFileName);
                }
                if (!excludeRoleAssignments && web.HasUniqueRoleAssignments)
                {
                    Dictionary<string, Dictionary<string, object>> roleDefsDict = new Dictionary<string, Dictionary<string, object>>();
                    foreach (RoleDefinition roleDefition in web.RoleDefinitions)
                    {
                        Dictionary<string, object> roleDefDict = new Dictionary<string, object>();
                        roleDefDict.Add("Id", roleDefition.Id);
                        roleDefDict.Add("Name", roleDefition.Name);
                        roleDefDict.Add("RoleTypeKind", roleDefition.RoleTypeKind.ToString());
                        roleDefsDict.Add(roleDefition.Id.ToString(), roleDefDict);
                    }
                    webDict.Add("RoleDefinitions", roleDefsDict);
                    clientContext.Load(web.RoleAssignments,
                        roleAssignment => roleAssignment.Include(
                                item => item.PrincipalId,
                                item => item.Member.LoginName,
                                item => item.Member.Title,
                                item => item.Member.PrincipalType,
                                item => item.RoleDefinitionBindings
                            ));
                    clientContext.ExecuteQuery();
                    SetRoleAssignments(web.RoleAssignments, webDict);
                }

                ListCollection lists = web.Lists;
                GroupCollection groups = web.SiteGroups;
                UserCollection users = web.SiteUsers;
                clientContext.Load(lists);
                clientContext.Load(groups,
                    grp => grp.Include(
                        item => item.Users,
                        item => item.Id,
                        item => item.LoginName,
                        item => item.PrincipalType,
                        item => item.Title
                    ));
                clientContext.Load(users);
                clientContext.ExecuteQuery();

                Dictionary<string, object> usersAndGroupsDict = new Dictionary<string, object>();
                if (webToFetch.isTopLevel)
                {
                    foreach (Group group in groups)
                    {
                        Dictionary<string, object> groupDict = new Dictionary<string, object>();
                        groupDict.Add("Id", "" + group.Id);
                        groupDict.Add("LoginName", group.LoginName);
                        groupDict.Add("PrincipalType", group.PrincipalType.ToString());
                        groupDict.Add("Title", group.Title);
                        Dictionary<string, object> innerUsersDict = new Dictionary<string, object>();
                        foreach (User user in group.Users)
                        {
                            Dictionary<string, object> innerUserDict = new Dictionary<string, object>();
                            innerUserDict.Add("LoginName", user.LoginName);
                            innerUserDict.Add("Id", "" + user.Id);
                            innerUserDict.Add("PrincipalType", user.PrincipalType.ToString());
                            innerUserDict.Add("IsSiteAdmin", "" + user.IsSiteAdmin);
                            innerUserDict.Add("Title", user.Title);
                            innerUsersDict.Add(user.LoginName, innerUserDict);
                        }
                        groupDict.Add("Users", innerUsersDict);
                        usersAndGroupsDict.Add(group.LoginName, groupDict);
                    }
                    foreach (User user in users)
                    {
                        Dictionary<string, object> userDict = new Dictionary<string, object>();
                        userDict.Add("LoginName", user.LoginName);
                        userDict.Add("Id", "" + user.Id);
                        userDict.Add("PrincipalType", user.PrincipalType.ToString());
                        userDict.Add("IsSiteAdmin", "" + user.IsSiteAdmin);
                        userDict.Add("Title", user.Title);
                        usersAndGroupsDict.Add(user.LoginName, userDict);
                    }
                }
                webDict.Add("UsersAndGroups", usersAndGroupsDict);
                Dictionary<string, object> listsDict = new Dictionary<string, object>();
                foreach (List list in lists)
                {
                    // All sites have a few lists that we don't care about exporting. Exclude these.
                    if (ignoreSiteNames.Contains(list.Title))
                    {
                        //Console.WriteLine("Skipping built-in sharepoint list " + list.Title);
                        continue;
                    }
                    ListToFetch listToFetch = new ListToFetch();
                    listToFetch.listId = list.Id;
                    listToFetch.listsDict = listsDict;
                    listToFetch.site = url;
                    listFetchBlockingCollection.Add(listToFetch);
                }
                ListsOutput nextListOutput = new ListsOutput();
                nextListOutput.jsonPath = listsJsonPath;
                nextListOutput.listsDict = listsDict;
                listsOutput.Add(nextListOutput);
            }
        }

        public void FetchList()
        {
            //Console.WriteLine("Starting Thread {0}", Thread.CurrentThread.ManagedThreadId);
            ListToFetch listToFetch;
            while (listFetchBlockingCollection.TryTake(out listToFetch))
            {
                try
                {
                    CheckAbort();
                    ClientContext clientContext = getClientContext(listToFetch.site);
                    List list = clientContext.Web.Lists.GetById(listToFetch.listId);
                    clientContext.Load(list, lslist => lslist.HasUniqueRoleAssignments, lslist => lslist.Id, lslist => lslist.Title, lslist => lslist.BaseType,
                        lslist => lslist.Description, lslist => lslist.LastItemModifiedDate, lslist => lslist.RootFolder, lslist => lslist.DefaultDisplayFormUrl);
                    clientContext.ExecuteQuery();
                    Console.WriteLine("Thread {0} - Parsing list site={1}, listID={2}, listTitle={3}", Thread.CurrentThread.ManagedThreadId, listToFetch.site, list.Id, list.Title);
                    CamlQuery camlQuery = new CamlQuery();
                    camlQuery.ViewXml = "<View Scope=\"RecursiveAll\"></View>";
                    ListItemCollection collListItem = list.GetItems(camlQuery);
                    clientContext.Load(collListItem);
                    clientContext.Load(collListItem,
                        items => items.Include(
                            item => item.Id,
                            item => item.DisplayName,
                            item => item.HasUniqueRoleAssignments,
                            item => item.Folder,
                            item => item.File,
                            item => item.ContentType
                            ));
                    clientContext.Load(list.RootFolder.Files);
                    clientContext.Load(list.RootFolder.Folders);
                    clientContext.Load(list.RootFolder);
                    try
                    {
                        clientContext.ExecuteQuery();
                    }
                    catch (Exception e)
                    {
                        Console.WriteLine("Could not fetch listID=" + list.Id + ", listTitle=" + list.Title + " because of error " + e.Message);
                        return;
                    }
                    Dictionary<string, object> listDict = new Dictionary<string, object>();
                    listDict.Add("Id", list.Id);
                    listDict.Add("Title", list.Title);
                    listDict.Add("BaseType", list.BaseType.ToString());
                    listDict.Add("Description", list.Description);
                    listDict.Add("LastItemModifiedDate", list.LastItemModifiedDate.ToString());
                    List<Dictionary<string, object>> itemsList = new List<Dictionary<string, object>>();
                    foreach (ListItem listItem in collListItem)
                    {
                        Dictionary<string, object> itemDict = new Dictionary<string, object>();
                        itemDict.Add("DisplayName", listItem.DisplayName);
                        itemDict.Add("Id", listItem.Id);
                        string contentTypeName = "";
                        try
                        {
                            contentTypeName = listItem.ContentType.Name;
                        }
                        catch (Exception excep)
                        {
                            Console.WriteLine("Couldn't get listItem.ContentType.Name for list item {0} due to {1}", listItem.Id, excep.Message);
                        }
                        itemDict.Add("ContentTypeName", contentTypeName);
                        if (contentTypeName.Equals("Document") && listItem.FieldValues.ContainsKey("FileRef"))
                        {
                            itemDict.Add("Url", topParentSite + listItem["FileRef"]);
                        }
                        else
                        {
                            itemDict.Add("Url", topParentSite + list.DefaultDisplayFormUrl + string.Format("?ID={0}", listItem.Id));
                        }
                        if (listItem.File.ServerObjectIsNull == false)
                        {
                            itemDict.Add("TimeLastModified", listItem.File.TimeLastModified.ToString());
                            itemDict.Add("ListItemType", "List_Item");
                            if (maxFileSizeBytes < 0 || listItem.FieldValues.ContainsKey("File_x0020_Size") == false || int.Parse((string)listItem.FieldValues["File_x0020_Size"]) < maxFileSizeBytes)
                            {
                                string filePath = baseDir + Path.DirectorySeparatorChar + "files" + Path.DirectorySeparatorChar + Guid.NewGuid().ToString() + Path.GetExtension(listItem.File.Name);
                                FileToDownload toDownload = new FileToDownload();
                                toDownload.saveToPath = filePath;
                                toDownload.serverRelativeUrl = listItem.File.ServerRelativeUrl;
                                toDownload.site = site;
                                fileDownloadBlockingCollection.Add(toDownload);
                                itemDict.Add("ExportPath", filePath);
                            }
                        }
                        else if (listItem.Folder.ServerObjectIsNull == false)
                        {
                            itemDict.Add("ListItemType", "Folder");
                        }
                        else
                        {
                            itemDict.Add("ListItemType", "List_Item");
                        }
                        if (listItem.HasUniqueRoleAssignments)
                        {
                            clientContext.Load(listItem.RoleAssignments,
                                ras => ras.Include(
                                        item => item.PrincipalId,
                                        item => item.Member.LoginName,
                                        item => item.Member.Title,
                                        item => item.Member.PrincipalType,
                                        item => item.RoleDefinitionBindings));
                            clientContext.ExecuteQuery();
                            SetRoleAssignments(listItem.RoleAssignments, itemDict);
                        }
                        itemDict.Add("FieldValues", listItem.FieldValues);
                        if (listItem.FieldValues.ContainsKey("Attachments") && (bool)listItem.FieldValues["Attachments"])
                        {
                            clientContext.Load(listItem.AttachmentFiles);
                            clientContext.ExecuteQuery();
                            List<Dictionary<string, object>> attachmentFileList = new List<Dictionary<string, object>>();
                            foreach (Attachment attachmentFile in listItem.AttachmentFiles)
                            {
                                Dictionary<string, object> attachmentFileDict = new Dictionary<string, object>();
                                attachmentFileDict.Add("Url", topParentSite + attachmentFile.ServerRelativeUrl);
                                string filePath = baseDir + Path.DirectorySeparatorChar + "files" + Path.DirectorySeparatorChar + Guid.NewGuid().ToString() + Path.GetExtension(attachmentFile.FileName);
                                FileToDownload toDownload = new FileToDownload();
                                toDownload.saveToPath = filePath;
                                toDownload.serverRelativeUrl = attachmentFile.ServerRelativeUrl;
                                toDownload.site = site;
                                fileDownloadBlockingCollection.Add(toDownload);
                                attachmentFileDict.Add("ExportPath", filePath);
                                attachmentFileDict.Add("FileName", attachmentFile.FileName);
                                attachmentFileList.Add(attachmentFileDict);
                            }
                            itemDict.Add("AttachmentFiles", attachmentFileList);
                        }
                        itemsList.Add(itemDict);
                    }
                    listDict.Add("Items", itemsList);
                    listDict.Add("Url", topParentSite + list.RootFolder.ServerRelativeUrl);
                    //listDict.Add("Files", IndexFolder(clientContext, list.RootFolder));
                    if (list.HasUniqueRoleAssignments)
                    {
                        clientContext.Load(list.RoleAssignments,
                        roleAssignments => roleAssignments.Include(
                                item => item.PrincipalId,
                                item => item.Member.LoginName,
                                item => item.Member.Title,
                                item => item.Member.PrincipalType,
                                item => item.RoleDefinitionBindings
                        ));
                        clientContext.ExecuteQuery();
                        SetRoleAssignments(list.RoleAssignments, listDict);
                    }
                    if (listToFetch.listsDict.ContainsKey(list.Id.ToString()))
                    {
                        Console.WriteLine("Duplicate key " + list.Id);
                    }
                    else
                    {
                        listToFetch.listsDict.Add(list.Id.ToString(), listDict);
                    }
                }
                catch (Exception e)
                {
                    Console.WriteLine("Got error trying to fetch list {0}: {1}", listToFetch.listId, e.Message);
                    Console.WriteLine(e.StackTrace);
                }
            }
        }

        public static void deleteDirectory(string targetDir)
        {
            string[] files = Directory.GetFiles(targetDir);
            string[] dirs = Directory.GetDirectories(targetDir);

            foreach (string file in files)
            {
                System.IO.File.SetAttributes(file, FileAttributes.Normal);
                System.IO.File.Delete(file);
            }

            foreach (string dir in dirs)
            {
                deleteDirectory(dir);
            }

            Directory.Delete(targetDir, false);
        }

        void writeWebJson()
        {
            string webJsonPath = baseDir + Path.DirectorySeparatorChar + "web-" + Guid.NewGuid() + ".json";
            System.IO.File.WriteAllText(webJsonPath, serializer.Serialize(rootWebDict));
        }

        public ClientContext getClientContext(string site)
        {
            ClientContext clientContext = new ClientContext(site);
            clientContext.RequestTimeout = -1;
            if (cc != null)
            {
                clientContext.Credentials = cc;
            }
            return clientContext;
        }

        public void writeAllListsToJson()
        {
            foreach (ListsOutput nextListOutput in listsOutput)
            {
                System.IO.File.WriteAllText(nextListOutput.jsonPath, serializer.Serialize(nextListOutput.listsDict));
                Console.WriteLine("Exported list to {0}", nextListOutput.jsonPath);
            }
        }

        public void getSubWebs(string url, Dictionary<string, object> parentWebDict)
        {
            CheckAbort();
            ClientContext clientContext = getClientContext(url);
            Web oWebsite = clientContext.Web;
            clientContext.Load(oWebsite, website => website.Webs);
            try
            {
                clientContext.ExecuteQuery();
            }
            catch (Exception ex)
            {
                Console.WriteLine("Could not load site {0} because of Error {1}", url, ex.Message);
                return;
            }
            WebToFetch webToFetch = new WebToFetch();
            webToFetch.url = url;
            webToFetch.isTopLevel = parentWebDict == null;
            webToFetch.webDict = new Dictionary<string, object>();

            foreach (Web orWebsite in oWebsite.Webs)
            {
                getSubWebs(orWebsite.Url, webToFetch.webDict);
            }
            if (parentWebDict != null)
            {
                Dictionary<string, object> subWebsDict = null;
                if (!parentWebDict.ContainsKey("SubWebs"))
                {
                    subWebsDict = new Dictionary<string, object>();
                    parentWebDict.Add("SubWebs", subWebsDict);
                }
                else
                {
                    subWebsDict = (Dictionary<string, object>)parentWebDict["SubWebs"];
                }
                subWebsDict.Add(url, webToFetch.webDict);
            }
            else
            {
                rootWebDict = webToFetch.webDict;
            }
            webFetchBlockingCollection.Add(webToFetch);
        }


        static void SetRoleAssignments(RoleAssignmentCollection roleAssignments, Dictionary<string, object> itemDict)
        {
            Dictionary<string, object> roleAssignmentsDict = new Dictionary<string, object>();
            foreach (RoleAssignment roleAssignment in roleAssignments)
            {
                Dictionary<string, object> roleAssignmentDict = new Dictionary<string, object>();
                List<string> defs = new List<string>();
                foreach (RoleDefinition roleDefinition in roleAssignment.RoleDefinitionBindings)
                {
                    defs.Add(roleDefinition.Id.ToString());
                }
                roleAssignmentDict.Add("LoginName", roleAssignment.Member.LoginName);
                roleAssignmentDict.Add("Title", roleAssignment.Member.Title);
                roleAssignmentDict.Add("PrincipalType", roleAssignment.Member.PrincipalType.ToString());
                roleAssignmentDict.Add("RoleDefinitionIds", defs);
                roleAssignmentsDict.Add(roleAssignment.Member.LoginName, roleAssignmentDict);
            }
            itemDict.Add("RoleAssignments", roleAssignmentsDict);
        }

        static public List<Dictionary<string, object>> IndexFolder(ClientContext clientContext, Folder folder)
        {
            List<Dictionary<string, object>> files = new List<Dictionary<string, object>>();
            foreach (Microsoft.SharePoint.Client.File file in folder.Files)
            {
                Dictionary<string, object> fileDict = new Dictionary<string, object>();
                fileDict.Add("Title", file.Title);
                fileDict.Add("FileType", "file");
                fileDict.Add("Name", file.Name);
                fileDict.Add("TimeCreated", file.TimeCreated);
                fileDict.Add("TimeLastModified", file.TimeLastModified);
                fileDict.Add("ServerRelativeUrl", file.ServerRelativeUrl);
                files.Add(fileDict);
            }
            Dictionary<string, object> foldersDict = new Dictionary<string, object>();
            foreach (Folder innerFolder in folder.Folders)
            {
                clientContext.Load(innerFolder);
                clientContext.Load(innerFolder.Files);
                clientContext.Load(innerFolder.Folders);
                clientContext.ExecuteQuery();
                Dictionary<string, object> innerFolderDict = new Dictionary<string, object>();
                innerFolderDict.Add("Name", innerFolder.Name);
                innerFolderDict.Add("FileType", "folder");
                //innerFolderDict.Add("Properties", innerFolder.Properties);
                innerFolderDict.Add("WelcomePage", innerFolder.WelcomePage);
                innerFolderDict.Add("ServerRelativeUrl", innerFolder.ServerRelativeUrl);
                innerFolderDict.Add("ParentServerRelativeUrl", folder.ServerRelativeUrl);
                innerFolderDict.Add("Files", IndexFolder(clientContext, innerFolder));
                files.Add(innerFolderDict);
            }
            return files;
        }

        static public SecureString GetPassword()
        {
            var pwd = new SecureString();
            while (true)
            {
                ConsoleKeyInfo i = Console.ReadKey(true);
                if (i.Key == ConsoleKey.Enter)
                {
                    break;
                }
                else if (i.Key == ConsoleKey.Backspace)
                {
                    if (pwd.Length > 0)
                    {
                        pwd.RemoveAt(pwd.Length - 1);
                        Console.Write("\b \b");
                    }
                }
                else
                {
                    pwd.AppendChar(i.KeyChar);
                    Console.Write("*");
                }
            }
            return pwd;
        }

        public XmlDocument getContent(string siteUrl, String contentType, string contentId)
        {
            List<string> allSites = new List<string>();
            var _url = string.Format("{0}/_vti_bin/SiteData.asmx", siteUrl);
            var _action = "http://schemas.microsoft.com/sharepoint/soap/GetContent";

            XmlDocument soapEnvelopeXml = CreateSoapEnvelope(contentType, contentId);
            HttpWebRequest webRequest = CreateWebRequest(_url, _action);
            InsertSoapEnvelopeIntoWebRequest(soapEnvelopeXml, webRequest);

            // begin async call to web request.
            IAsyncResult asyncResult = webRequest.BeginGetResponse(null, null);

            // suspend this thread until call is complete. You might want to
            // do something usefull here like update your UI.
            asyncResult.AsyncWaitHandle.WaitOne();

            // get the response from the completed web request.
            XmlDocument contentDatabaseResult = new XmlDocument();
            using (WebResponse webResponse = webRequest.EndGetResponse(asyncResult))
            {
                using (StreamReader rd = new StreamReader(webResponse.GetResponseStream()))
                {
                    contentDatabaseResult.Load(rd);
                }
            }
            return contentDatabaseResult;
        }

        public List<string> GetAllTopLevelSites()
        {
            List<String> allSites = new List<string>();
            XmlDocument virtualServerGetContentResult = getContent(site, "VirtualServer", null);
            XmlNode contentResultNode = virtualServerGetContentResult.SelectSingleNode("//*[local-name() = 'GetContentResult']");
            if (contentResultNode == null || contentResultNode.InnerText == null)
            {
                throw new Exception(string.Format("Cannot list top level sites from {0}", site));
            }
            XmlDocument innerXmlDoc = new XmlDocument();
            innerXmlDoc.LoadXml(contentResultNode.InnerText);
            string contentDatabaseId = innerXmlDoc.SelectSingleNode("//*[local-name() = 'ContentDatabase']").Attributes["ID"].Value;
            if (contentDatabaseId == null)
            {
                throw new Exception(string.Format("Cannot list top level sites from {0}", site));
            }
            XmlDocument contentDatabaseGetContentResult = getContent(site, "ContentDatabase", contentDatabaseId);
            XmlNode contentDatabaseResultNode = contentDatabaseGetContentResult.SelectSingleNode("//*[local-name() = 'GetContentResult']");
            if (contentDatabaseResultNode == null || contentDatabaseResultNode.InnerText == null)
            {
                throw new Exception(string.Format("Cannot list top level sites from {0}", site));
            }
            innerXmlDoc = new XmlDocument();
            innerXmlDoc.LoadXml(contentDatabaseResultNode.InnerText);
            XmlNodeList sites = innerXmlDoc.SelectNodes("//*[local-name() = 'Site']");
            foreach (XmlNode siteNode in sites)
            {
                allSites.Add(siteNode.Attributes["URL"].Value);
            }
            return allSites;
        }

        HttpWebRequest CreateWebRequest(string url, string action)
        {
            HttpWebRequest webRequest = (HttpWebRequest)WebRequest.Create(url);
            webRequest.Headers.Add("SOAPAction", action);
            webRequest.ContentType = "text/xml;charset=\"utf-8\"";
            webRequest.Accept = "text/xml";
            webRequest.Method = "POST";
            webRequest.Credentials = cc;
            return webRequest;
        }

        XmlDocument CreateSoapEnvelope(string objectType, string objectId)
        {
            XmlDocument soapEnvelopeDocument = new XmlDocument();
            string soapEnv = @"<soapenv:Envelope xmlns:soapenv=""http://schemas.xmlsoap.org/soap/envelope/"" xmlns:soap=""http://schemas.microsoft.com/sharepoint/soap/"">
   <soapenv:Header/>
   <soapenv:Body>
      <soap:GetContent>
         <soap:objectType>{0}</soap:objectType>
{1}
         <soap:retrieveChildItems>true</soap:retrieveChildItems>
         <soap:securityOnly>false</soap:securityOnly>
      </soap:GetContent>
   </soapenv:Body>
</soapenv:Envelope>";
            soapEnvelopeDocument.LoadXml(string.Format(soapEnv, objectType, objectId == null ? "" : "<soap:objectId>" + objectId + "</soap:objectId>"));
            return soapEnvelopeDocument;
        }

        void InsertSoapEnvelopeIntoWebRequest(XmlDocument soapEnvelopeXml, HttpWebRequest webRequest)
        {
            using (Stream stream = webRequest.GetRequestStream())
            {
                soapEnvelopeXml.Save(stream);
            }
        }
    }
}