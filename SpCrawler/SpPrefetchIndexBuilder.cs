using System;
using System.Net;
using System.Diagnostics;
using System.Collections.Generic;
using Microsoft.SharePoint.Client;
using System.Threading.Tasks;
using System.Threading;
using System.Net.Http;
using System.IO;

namespace SpPrefetchIndexBuilder {
  class SpPrefetchIndexBuilder {
    public static SharepointExporterConfig config;
    public static void CheckAbort() {
      if (System.IO.File.Exists(config.baseDir + Path.DirectorySeparatorChar + ".." + Path.DirectorySeparatorChar + ".doabort")) {
        Console.WriteLine("The .doabort file was found. Stopping program");
        Environment.Exit(0);
      }
    }
    public static int fileCount = 0;
    public string rootSite;
    public static HttpClient client;
    public CredentialCache cc = null;
    public List<ListToFetch> listFetchList = new List<ListToFetch>();
    public List<WebToFetch> webFetchList = new List<WebToFetch>();
    public List<FileToDownload> fileDownloadList = new List<FileToDownload>();
    public Dictionary<string, object> rootWebDict;
    public List<ListsOutput> listsOutput = new List<ListsOutput>();

    public List<string> ignoreListNames = new List<string>();

    static void Main(string[] args) {
      
      config = new SharepointExporterConfig(args);
      if (config.customBaseDir && config.deleteExistingOutputDir && Directory.Exists(config.baseDir)) {
        deleteDirectory(config.baseDir);
      }
      Directory.CreateDirectory(config.baseDir);
      if (!config.excludeLists) {
        Directory.CreateDirectory(config.baseDir + Path.DirectorySeparatorChar + "lists");
      }
      if (!config.excludeLists && !config.excludeFiles) {
        Directory.CreateDirectory(config.baseDir + Path.DirectorySeparatorChar + "files");
      }

      ServicePointManager.DefaultConnectionLimit = config.numThreads;

      foreach (string site in config.sites) {
        SpPrefetchIndexBuilder spib = new SpPrefetchIndexBuilder(site);
        spib.buildFullIndex();
        // todo check for incremental
      }

    }

    public void buildFullIndex() {
      try {
        Stopwatch swAll = Stopwatch.StartNew();
        Console.WriteLine("Building full index for site {0}", rootSite);

        Stopwatch swWeb = Stopwatch.StartNew();
        getWebs(rootSite, rootSite, null);
        Parallel.ForEach(
          webFetchList,
          new ParallelOptions { MaxDegreeOfParallelism = config.numThreads },
          toFetchWeb => { FetchWeb(toFetchWeb); }
        );
        writeWebJson();
        Console.WriteLine("Web fetch of {0} complete. Took {1} milliseconds.", rootSite, swWeb.ElapsedMilliseconds);

        if (!config.excludeLists) {
          Stopwatch swLists = Stopwatch.StartNew();
          Parallel.ForEach(
            listFetchList,
            new ParallelOptions { MaxDegreeOfParallelism = config.numThreads },
            toFetchList => { FetchList(toFetchList); }
          );
          writeAllListsToJson();
          Console.WriteLine("Lists metadata dump of {0} complete. Took {1} milliseconds.", 
                            rootSite, swLists.ElapsedMilliseconds);
          if (config.excludeFiles) {
            Console.WriteLine("Downloading the files recieved during the index building");
            Parallel.ForEach(
              fileDownloadList,
              new ParallelOptions { MaxDegreeOfParallelism = config.numThreads },
              toDownload => { DownloadFile(toDownload); }
            );
          }
        }
        Console.WriteLine("Export complete! Took {0} milliseconds.", swAll.ElapsedMilliseconds);
      } catch (Exception anyException) {
        Console.WriteLine("Prefetch index building failed for site {0} due to {1}", rootSite, anyException);
        Environment.Exit(1);
      }
    }

    public SpPrefetchIndexBuilder(string rootSite) {
      this.rootSite = rootSite;

      if (rootSite.EndsWith("/", StringComparison.CurrentCulture)) {
        rootSite = rootSite.Substring(0, rootSite.Length - 1);
      }

      cc = new CredentialCache();

      cc.Add(new Uri(rootSite), "NTLM", config.networkCredentials);
      HttpClientHandler handler = new HttpClientHandler();
      handler.Credentials = cc;
      client = new HttpClient(handler);
      client.Timeout = TimeSpan.FromSeconds(30);
      client.DefaultRequestHeaders.ConnectionClose = true;
    }

    public void DownloadFile(FileToDownload toDownload) {
      if (config.maxFiles > 0 && fileCount++ >= config.maxFiles) {
        Console.WriteLine("Not downloading file {0} because maxFiles limit of {1} has been reached.", 
                          toDownload.serverRelativeUrl, config.maxFiles);
        return;
      }
      try {
        var responseResult = client.GetAsync(rootSite + toDownload.serverRelativeUrl);
        if (responseResult.Result != null && responseResult.Result.StatusCode == System.Net.HttpStatusCode.OK) {
          using (var memStream = responseResult.Result.Content.ReadAsStreamAsync().GetAwaiter().GetResult()) {
            using (var fileStream = System.IO.File.Create(toDownload.saveToPath)) {
              memStream.CopyTo(fileStream);
            }
          }
          Console.WriteLine("Thread {0} - Successfully downloaded {1} to {2}", Thread.CurrentThread.ManagedThreadId, 
                            toDownload.serverRelativeUrl, toDownload.saveToPath);
        } else {
          Console.WriteLine("Got non-OK status {0} when trying to download url {1}", responseResult.Result.StatusCode, 
                            rootSite + toDownload.serverRelativeUrl);
        }
      } catch (Exception e) {
        Console.WriteLine("Gave up trying to download url {0}{1} to file {2} due to error: {3}", rootSite, 
                          toDownload.serverRelativeUrl, toDownload.saveToPath, e);
      }
    }

    public void FetchWeb(WebToFetch webToFetch) {
      CheckAbort();
      DateTime now = DateTime.Now;
      string url = webToFetch.url;
      Console.WriteLine("Thread {0} exporting web {1}", Thread.CurrentThread.ManagedThreadId, url);
      ClientContext clientContext = getClientContext(url);

      Web web = clientContext.Web;

      var site = clientContext.Site;
      if (config.excludeRoleDefinitions && config.excludeRoleDefinitions) {
        clientContext.Load(web, website => website.Webs,
                           website => website.Title,
                           website => website.Url,
                           website => website.Description,
                           website => website.Id,
                           website => website.LastItemModifiedDate);
      } else {
        clientContext.Load(web, website => website.Webs,
                           website => website.Title,
                           website => website.Url,
                           website => website.RoleDefinitions,
                           website => website.RoleAssignments,
                           website => website.HasUniqueRoleAssignments,
                           website => website.Description,
                           website => website.Id,
                           website => website.LastItemModifiedDate);
      }
      try {
        clientContext.ExecuteQuery();
      } catch (Exception ex) {
        Console.WriteLine("Could not load site {0} because of Error {1}", url, ex.Message);
        return;
      }

      string listsFileName = Guid.NewGuid().ToString() + ".json";
      string listsJsonPath = config.baseDir + Path.DirectorySeparatorChar.ToString() + "lists" + Path.DirectorySeparatorChar.ToString() + listsFileName;
      Dictionary<string, object> webDict = webToFetch.webDict;
      webDict.Add("Title", web.Title);
      webDict.Add("Id", web.Id);
      webDict.Add("Description", web.Description);
      webDict.Add("Url", url);
      webDict.Add("LastItemModifiedDate", web.LastItemModifiedDate.ToString());
      webDict.Add("FetchedDate", now);
      if (!config.excludeLists) {
        webDict.Add("ListsFileName", listsFileName);
      }
      if (!config.excludeRoleAssignments && web.HasUniqueRoleAssignments) {
        Dictionary<string, Dictionary<string, object>> roleDefsDict = new Dictionary<string, Dictionary<string, object>>();
        foreach (RoleDefinition roleDefition in web.RoleDefinitions) {
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
                    item => item.PrincipalId,
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
      if (config.excludeGroupMembers) {
        clientContext.Load(groups,
          grp => grp.Include(
              item => item.Id,
              item => item.LoginName,
              item => item.PrincipalType,
              item => item.Title
          ));
      } else {
        clientContext.Load(groups,
          grp => grp.Include(
              item => item.Users,
              item => item.Id,
              item => item.LoginName,
              item => item.PrincipalType,
              item => item.Title
          ));
      }
      clientContext.Load(users);
      clientContext.ExecuteQuery();

      if (webToFetch.isRootLevelSite && !config.excludeUsersAndGroups) {
        Dictionary<string, object> usersAndGroupsDict = new Dictionary<string, object>();
        foreach (Group group in groups) {
          Dictionary<string, object> groupDict = new Dictionary<string, object>();
          groupDict.Add("Id", "" + group.Id);
          groupDict.Add("LoginName", group.LoginName);
          groupDict.Add("PrincipalType", group.PrincipalType.ToString());
          groupDict.Add("Title", group.Title);
          Dictionary<string, object> innerUsersDict = new Dictionary<string, object>();
          if (!config.excludeGroupMembers) {
            foreach (User user in group.Users) {
              Dictionary<string, object> innerUserDict = new Dictionary<string, object>();
              innerUserDict.Add("LoginName", user.LoginName);
              innerUserDict.Add("Id", "" + user.Id);
              innerUserDict.Add("PrincipalType", user.PrincipalType.ToString());
              innerUserDict.Add("IsSiteAdmin", "" + user.IsSiteAdmin);
              innerUserDict.Add("Title", user.Title);
              innerUsersDict.Add(user.LoginName, innerUserDict);
            }
            groupDict.Add("Users", innerUsersDict);
          }
          usersAndGroupsDict.Add(group.LoginName, groupDict);
        }
        foreach (User user in users) {
          Dictionary<string, object> userDict = new Dictionary<string, object>();
          userDict.Add("LoginName", user.LoginName);
          userDict.Add("Id", "" + user.Id);
          userDict.Add("PrincipalType", user.PrincipalType.ToString());
          userDict.Add("IsSiteAdmin", "" + user.IsSiteAdmin);
          userDict.Add("Title", user.Title);
          usersAndGroupsDict.Add(user.LoginName, userDict);
        }
        webDict.Add("UsersAndGroups", usersAndGroupsDict);
      }
      webDict.Add("IsRootLevelSite", webToFetch.isRootLevelSite);
      if (webToFetch.rootLevelSiteUrl != null) {
        webDict.Add("RootLevelSiteUrl", webToFetch.rootLevelSiteUrl);
      }
      Dictionary<string, object> listsDict = new Dictionary<string, object>();
      foreach (List list in lists) {
        // All sites have a few lists that we don't care about exporting. Exclude these.
        if (ignoreListNames.Contains(list.Title)) {
          //Console.WriteLine("Skipping built-in sharepoint list " + list.Title);
          continue;
        }
        ListToFetch listToFetch = new ListToFetch();
        listToFetch.listId = list.Id;
        listToFetch.listsDict = listsDict;
        listToFetch.site = url;
        listFetchList.Add(listToFetch);
      }
      ListsOutput nextListOutput = new ListsOutput();
      nextListOutput.jsonPath = listsJsonPath;
      nextListOutput.listsDict = listsDict;
      listsOutput.Add(nextListOutput);
    }

    public void FetchList(ListToFetch listToFetch) {
      try {
        CheckAbort();
        DateTime now = DateTime.Now;
        ClientContext clientContext = getClientContext(listToFetch.site);
        List list = clientContext.Web.Lists.GetById(listToFetch.listId);
        clientContext.Load(list, lslist => lslist.HasUniqueRoleAssignments, lslist => lslist.Id, 
                           lslist => lslist.Title, lslist => lslist.BaseType,
            lslist => lslist.Description, lslist => lslist.LastItemModifiedDate, lslist => lslist.RootFolder, 
                           lslist => lslist.DefaultDisplayFormUrl);
        clientContext.ExecuteQuery();
        Console.WriteLine("Thread {0} - Parsing list site={1}, listID={2}, listTitle={3}", Thread.CurrentThread.ManagedThreadId, 
                          listToFetch.site, list.Id, list.Title);
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
        try {
          clientContext.ExecuteQuery();
        } catch (Exception e) {
          Console.WriteLine("Could not fetch listID=" + list.Id + ", listTitle=" + list.Title + " because of error " + e.Message);
          return;
        }
        Dictionary<string, object> listDict = new Dictionary<string, object>();
        listDict.Add("Id", list.Id);
        listDict.Add("Title", list.Title);
        listDict.Add("BaseType", list.BaseType.ToString());
        listDict.Add("Description", list.Description);
        listDict.Add("LastItemModifiedDate", list.LastItemModifiedDate.ToString());
        listDict.Add("FetchedDate", now);
        List<Dictionary<string, object>> itemsList = new List<Dictionary<string, object>>();
        foreach (ListItem listItem in collListItem) {
          Dictionary<string, object> itemDict = new Dictionary<string, object>();
          itemDict.Add("DisplayName", listItem.DisplayName);
          itemDict.Add("Id", listItem.Id);
          string contentTypeName = "";
          try {
            contentTypeName = listItem.ContentType.Name;
          } catch (Exception excep) {
            Console.WriteLine("Couldn't get listItem.ContentType.Name for list item {0} due to {1}", listItem.Id, excep.Message);
          }
          itemDict.Add("ContentTypeName", contentTypeName);
          if (contentTypeName.Equals("Document") && listItem.FieldValues.ContainsKey("FileRef")) {
            itemDict.Add("Url", rootSite + listItem["FileRef"]);
          } else {
            itemDict.Add("Url", rootSite + list.DefaultDisplayFormUrl + string.Format("?ID={0}", listItem.Id));
          }
          if (listItem.File.ServerObjectIsNull == false) {
            itemDict.Add("TimeLastModified", listItem.File.TimeLastModified.ToString());
            itemDict.Add("ListItemType", "List_Item");
            if (config.maxFileSizeBytes < 0 || listItem.FieldValues.ContainsKey("File_x0020_Size") == false || 
                int.Parse((string)listItem.FieldValues["File_x0020_Size"]) < maxFileSizeBytes) {
              string filePath = config.baseDir + Path.DirectorySeparatorChar + "files" + Path.DirectorySeparatorChar + 
                                              Guid.NewGuid().ToString() + Path.GetExtension(listItem.File.Name);
              FileToDownload toDownload = new FileToDownload();
              toDownload.saveToPath = filePath;
              toDownload.serverRelativeUrl = listItem.File.ServerRelativeUrl;
              fileDownloadList.Add(toDownload);
              itemDict.Add("ExportPath", filePath);
            }
          } else if (listItem.Folder.ServerObjectIsNull == false) {
            itemDict.Add("ListItemType", "Folder");
          } else {
            itemDict.Add("ListItemType", "List_Item");
          }
          if (listItem.HasUniqueRoleAssignments) {
            clientContext.Load(listItem.RoleAssignments,
                ras => ras.Include(
                        item => item.PrincipalId,
                        item => item.Member.LoginName,
                        item => item.Member.Title,
                        item => item.Member.PrincipalType,
                        item => item.RoleDefinitionBindings));
            clientContext.ExecuteQuery();
            Console.WriteLine("List Item {0} has unique role assignments: {1}", itemDict["Url"], listItem.RoleAssignments);
            SetRoleAssignments(listItem.RoleAssignments, itemDict);
          }
          itemDict.Add("FieldValues", listItem.FieldValues);
          if (listItem.FieldValues.ContainsKey("Attachments") && (bool)listItem.FieldValues["Attachments"]) {
            clientContext.Load(listItem.AttachmentFiles);
            clientContext.ExecuteQuery();
            List<Dictionary<string, object>> attachmentFileList = new List<Dictionary<string, object>>();
            foreach (Attachment attachmentFile in listItem.AttachmentFiles) {
              Dictionary<string, object> attachmentFileDict = new Dictionary<string, object>();
              attachmentFileDict.Add("Url", rootSite + attachmentFile.ServerRelativeUrl);
              string filePath = config.baseDir + Path.DirectorySeparatorChar + "files" + Path.DirectorySeparatorChar + 
                                              Guid.NewGuid().ToString() + Path.GetExtension(attachmentFile.FileName);
              FileToDownload toDownload = new FileToDownload();
              toDownload.saveToPath = filePath;
              toDownload.serverRelativeUrl = attachmentFile.ServerRelativeUrl;
              fileDownloadList.Add(toDownload);
              attachmentFileDict.Add("ExportPath", filePath);
              attachmentFileDict.Add("FileName", attachmentFile.FileName);
              attachmentFileList.Add(attachmentFileDict);
            }
            itemDict.Add("AttachmentFiles", attachmentFileList);
          }
          itemsList.Add(itemDict);
        }
        listDict.Add("Items", itemsList);
        listDict.Add("Url", rootSite + list.RootFolder.ServerRelativeUrl);
        //listDict.Add("Files", IndexFolder(clientContext, list.RootFolder));
        if (list.HasUniqueRoleAssignments) {
          clientContext.Load(list.RoleAssignments,
          roleAssignments => roleAssignments.Include(
                  item => item.PrincipalId,
                  item => item.Member.LoginName,
                  item => item.Member.Title,
                  item => item.Member.PrincipalType,
                  item => item.RoleDefinitionBindings
          ));
          clientContext.ExecuteQuery();
          Console.WriteLine("List {0} has unique role assignments: {1}", listDict["Url"], list.RoleAssignments);
          SetRoleAssignments(list.RoleAssignments, listDict);
        }
        if (listToFetch.listsDict.ContainsKey(list.Id.ToString())) {
          Console.WriteLine("Duplicate key " + list.Id);
        } else {
          listToFetch.listsDict.Add(list.Id.ToString(), listDict);
        }
      } catch (Exception e) {
        Console.WriteLine("Got error trying to fetch list {0}: {1}", listToFetch.listId, e.Message);
        Console.WriteLine(e.StackTrace);
      }
    }

    public static void deleteDirectory(string targetDir) {
      string[] files = Directory.GetFiles(targetDir);
      string[] dirs = Directory.GetDirectories(targetDir);

      foreach (string file in files) {
        System.IO.File.SetAttributes(file, FileAttributes.Normal);
        System.IO.File.Delete(file);
      }

      foreach (string dir in dirs) {
        deleteDirectory(dir);
      }

      Directory.Delete(targetDir, false);
    }

    void writeWebJson() {
      string webJsonPath = config.baseDir + Path.DirectorySeparatorChar + "web-" + Guid.NewGuid() + ".json";
      System.IO.File.WriteAllText(webJsonPath, config.serializer.Serialize(rootWebDict));
    }

    public ClientContext getClientContext(string site) {
      ClientContext clientContext = new ClientContext(site);
      clientContext.RequestTimeout = -1;
      if (cc != null) {
        clientContext.Credentials = cc;
      }
      return clientContext;
    }

    public void writeAllListsToJson() {
      foreach (ListsOutput nextListOutput in listsOutput) {
        System.IO.File.WriteAllText(nextListOutput.jsonPath, config.serializer.Serialize(nextListOutput.listsDict));
        Console.WriteLine("Exported list to {0}", nextListOutput.jsonPath);
      }
    }

    public void getWebs(string url, string rootLevelSiteUrl, Dictionary<string, object> parentWebDict) {
      CheckAbort();
      ClientContext clientContext = getClientContext(url);
      Web oWebsite = clientContext.Web;
      clientContext.Load(oWebsite, website => website.Webs);
      try {
        clientContext.ExecuteQuery();
      } catch (Exception ex) {
        Console.WriteLine("Could not load site {0} because of Error {1}", url, ex.Message);
        return;
      }
      WebToFetch webToFetch = new WebToFetch();
      webToFetch.url = url;
      if (parentWebDict != null) {
        webToFetch.rootLevelSiteUrl = rootLevelSiteUrl;
      }
      webToFetch.isRootLevelSite = parentWebDict == null;
      webToFetch.webDict = new Dictionary<string, object>();

      if (!config.excludeSubSites) {
        foreach (Web orWebsite in oWebsite.Webs) {
          getWebs(orWebsite.Url, rootLevelSiteUrl, webToFetch.webDict);
        }
      }
      if (parentWebDict != null) {
        Dictionary<string, object> subWebsDict = null;
        if (!parentWebDict.ContainsKey("SubWebs")) {
          subWebsDict = new Dictionary<string, object>();
          parentWebDict.Add("SubWebs", subWebsDict);
        } else {
          subWebsDict = (Dictionary<string, object>)parentWebDict["SubWebs"];
        }
        subWebsDict.Add(url, webToFetch.webDict);
      } else {
        rootWebDict = webToFetch.webDict;
      }
      webFetchList.Add(webToFetch);
    }


    static void SetRoleAssignments(RoleAssignmentCollection roleAssignments, Dictionary<string, object> itemDict) {
      Dictionary<string, object> roleAssignmentsDict = new Dictionary<string, object>();
      foreach (RoleAssignment roleAssignment in roleAssignments) {
        Dictionary<string, object> roleAssignmentDict = new Dictionary<string, object>();
        List<string> defs = new List<string>();
        foreach (RoleDefinition roleDefinition in roleAssignment.RoleDefinitionBindings) {
          defs.Add(roleDefinition.Id.ToString());
        }
        roleAssignmentDict.Add("LoginName", roleAssignment.Member.LoginName);
        roleAssignmentDict.Add("Title", roleAssignment.Member.Title);
        roleAssignmentDict.Add("PrincipalType", roleAssignment.Member.PrincipalType.ToString());
        roleAssignmentDict.Add("RoleDefinitionIds", defs);
        roleAssignmentDict.Add("PrincipalId", roleAssignment.PrincipalId);
        if (roleAssignment.Member.PrincipalType.Equals(Microsoft.SharePoint.Client.Utilities.PrincipalType.SecurityGroup) || 
            roleAssignment.Member.PrincipalType.Equals(Microsoft.SharePoint.Client.Utilities.PrincipalType.DistributionList)) {
          roleAssignmentsDict.Add(roleAssignment.Member.Title, roleAssignmentDict);  // Store these as domain\username
        } else {
          roleAssignmentsDict.Add(roleAssignment.Member.LoginName, roleAssignmentDict); // Use the normal LoginName for these
        }
      }
      itemDict.Add("RoleAssignments", roleAssignmentsDict);
    }

    static public List<Dictionary<string, object>> IndexFolder(ClientContext clientContext, Folder folder) {
      List<Dictionary<string, object>> files = new List<Dictionary<string, object>>();
      foreach (Microsoft.SharePoint.Client.File file in folder.Files) {
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
      foreach (Folder innerFolder in folder.Folders) {
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



  }
}