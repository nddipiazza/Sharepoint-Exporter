using System;
using System.Net;
using System.Diagnostics;
using System.Security;
using System.Collections.Generic;
using System.Web.Script.Serialization;
using Microsoft.SharePoint.Client;
using System.Collections.Concurrent;
using System.Threading.Tasks;

namespace SpPrefetchIndexBuilder
{
    class ToDownload
    {
        public String site;
        public String serverRelativeUrl;
        public String saveToPath;
    }

    class SpPrefetchIndexBuilder
    {
        public string defaultSite = "http://localhost/";
        public CredentialCache cc = null;
        public string site;
        public JavaScriptSerializer serializer = new JavaScriptSerializer();
        public string baseDir;
        public int maxFileSizeBytes = -1;

        public BlockingCollection<ToDownload> downloadList = new BlockingCollection<ToDownload>();

        public List<string> ignoreSiteNames = new List<string>();

        public void DownloadFilesFromQueue()
        {
            ToDownload toDownload;
            while (downloadList.TryTake(out toDownload))
            {
                ClientContext clientContext = getClientContext(toDownload.site);
                var fileInfo = File.OpenBinaryDirect(clientContext, toDownload.serverRelativeUrl);
                using (var fileStream = System.IO.File.Create(toDownload.saveToPath))
                {
                    Console.WriteLine("Saving {0} to {1}", toDownload.serverRelativeUrl, toDownload.saveToPath);
                    fileInfo.Stream.CopyTo(fileStream);
                }
            }
        }

        public SpPrefetchIndexBuilder(String [] args)
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

            String spMaxFileSizeBytes = Environment.GetEnvironmentVariable("SP_MAX_FILE_SIZE_BYTES");
            if (spMaxFileSizeBytes != null)
            {
                maxFileSizeBytes = int.Parse(spMaxFileSizeBytes);
            }
            serializer.MaxJsonLength = 209715200;
            if (args.Length >= 2 && (args[0].Equals("--help") || args[0].Equals("-help") || args[0].Equals("/help") || args.Length > 5 || args.Length == 3))
            {
                Console.WriteLine("USAGE: SpPrefetchIndexBuilder.exe [siteUrl] [outputDir] [domain] [username] [password (not recommended, do not specify to be prompted or use SP_PWD environment variable)]");
            }
            site = args.Length > 0 ? args[0] : defaultSite;
            baseDir = args.Length > 1 ? args[1] : System.IO.Directory.GetCurrentDirectory();
            if (site.EndsWith("/"))
            {
                site = site.Substring(0, site.Length - 1);
            }
            if (args.Length > 2)
            {
                cc = new CredentialCache();
                String spPassword = Environment.GetEnvironmentVariable("SP_PWD");
                if (spPassword == null)
                {
                    spPassword = args.Length >= 5 ? args[4] : null;
                }
                NetworkCredential nc;
                if (spPassword == null)
                {
                    Console.WriteLine("Please enter password for {0}", args[3]);
                    nc = new NetworkCredential(args[3], GetPassword(), args[2]);
                }
                else
                {
                    nc = new NetworkCredential(args[3], spPassword, args[2]);
                }
                cc.Add(new Uri(site), "NTLM", nc);
            }
        }

        static void Main(string[] args)
        {
            Stopwatch sw = Stopwatch.StartNew();
            SpPrefetchIndexBuilder spib = new SpPrefetchIndexBuilder(args);
            spib.getSubWebs(spib.site, spib.baseDir + "\\site-export-" + Guid.NewGuid().ToString().Substring(0, 8));
            Console.WriteLine("Downloading the files recieved during the index building");
            Parallel.For(0, 10, x => spib.DownloadFilesFromQueue());
            Console.WriteLine("Export complete. Took {0} milliseconds.", sw.ElapsedMilliseconds);
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

        public void getSubWebs(string url, string parentPath)
        {
            ClientContext clientContext = getClientContext(url);
            Web oWebsite = clientContext.Web;
            clientContext.Load(oWebsite, website => website.Webs, website => website.Title, website => website.Url, website => website.RoleDefinitions, website => website.RoleAssignments, website => website.HasUniqueRoleAssignments, website => website.Description, website => website.Id, website => website.LastItemModifiedDate);
            clientContext.ExecuteQuery();
            string path = parentPath + "\\" + oWebsite.Id;
            DownloadWeb(clientContext, oWebsite, url, path);
            foreach (Web orWebsite in oWebsite.Webs)
            {
                getSubWebs(orWebsite.Url, path);
            }
        }

        void DownloadWeb(ClientContext clientContext, Web web, string url, string path)
        {
            Console.WriteLine("Exporting site {0}", url);
            System.IO.Directory.CreateDirectory(path);
            Dictionary<string, object> webDict = new Dictionary<string, object>();
            webDict.Add("Title", web.Title);
            webDict.Add("Id", web.Id);
            webDict.Add("Description", web.Description);
            webDict.Add("Url", url);
            webDict.Add("LastItemModifiedDate", web.LastItemModifiedDate.ToString());
            if (web.HasUniqueRoleAssignments)
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
                            item => item.Member.PrincipalType,
                            item => item.RoleDefinitionBindings
                        ));
                clientContext.ExecuteQuery();
                SetRoleAssignments(web.RoleAssignments, webDict);
            }
            string siteJsonPath = path + "\\web.json";
            System.IO.File.WriteAllText(siteJsonPath, serializer.Serialize(webDict));
            Console.WriteLine("Exported site properties for site {0} to {1}", url, siteJsonPath);

            ListCollection lists = web.Lists;

            clientContext.Load(lists);            
            clientContext.ExecuteQuery();

            Dictionary<string, object> listsDict = new Dictionary<string, object>();
            foreach (List list in lists)
            {
                // All sites have a few lists that we don't care about exporting. Exclude these.
                if (ignoreSiteNames.Contains(list.Title))
                {
                    Console.WriteLine("Skipping built-in sharepoint site " + list.Title);
                    continue;
                }
                Console.WriteLine("Parsing list site={0}, listID={1}, listTitle={2}", url, list.Id, list.Title);
                clientContext.Load(list, lslist => lslist.HasUniqueRoleAssignments, lslist => lslist.Title, lslist => lslist.BaseType, lslist => lslist.Description, lslist => lslist.LastItemModifiedDate, lslist => lslist.RootFolder, lslist => lslist.DefaultDisplayFormUrl);
                clientContext.ExecuteQuery();
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
                     item => item.File
                     //,item => item.ContentType
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
                    continue;
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
                    if (listItem.File.ServerObjectIsNull == false)
                    {
                        itemDict.Add("TimeLastModified", listItem.File.TimeLastModified.ToString());
                        itemDict.Add("ListItemType", "List_Item");
                        if (maxFileSizeBytes < 0 || listItem.FieldValues.ContainsKey("File_x0020_Size") == false || int.Parse((string)listItem.FieldValues["File_x0020_Size"]) < maxFileSizeBytes)
                        {
                            string filePath = path + "\\" + Guid.NewGuid().ToString().Substring(0, 8) + System.IO.Path.GetExtension(listItem.File.Name);
                            ToDownload toDownload = new ToDownload();
                            toDownload.saveToPath = filePath;
                            toDownload.serverRelativeUrl = listItem.File.ServerRelativeUrl;
                            toDownload.site = site;
                            downloadList.Add(toDownload);
                            itemDict.Add("FileExportPath", filePath);
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
                    itemDict.Add("Url", site + list.DefaultDisplayFormUrl + string.Format("?ID={0}", listItem.Id));
                    if (listItem.HasUniqueRoleAssignments)
                    {
                        clientContext.Load(listItem.RoleAssignments,
                            ras => ras.Include(
                                    item => item.PrincipalId,
                                    item => item.Member.LoginName,
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
                            attachmentFileDict.Add("Url", site + attachmentFile.ServerRelativeUrl);
                            string filePath = path + "\\att_" + Guid.NewGuid().ToString().Substring(0, 8) + System.IO.Path.GetExtension(attachmentFile.FileName);
                            ToDownload toDownload = new ToDownload();
                            toDownload.saveToPath = filePath;
                            toDownload.serverRelativeUrl = attachmentFile.ServerRelativeUrl;
                            toDownload.site = site;
                            downloadList.Add(toDownload);
                            attachmentFileDict.Add("ExportPath", filePath);
                            attachmentFileDict.Add("FileName", attachmentFile.FileName);
                            attachmentFileList.Add(attachmentFileDict);
                        }
                        itemDict.Add("AttachmentFiles", attachmentFileList);
                    }
                    itemsList.Add(itemDict);
                }
                listDict.Add("Items", itemsList);
                listDict.Add("Url", site + list.RootFolder.ServerRelativeUrl);
                //listDict.Add("Files", IndexFolder(clientContext, list.RootFolder));
                if (list.HasUniqueRoleAssignments)
                {
                    clientContext.Load(list.RoleAssignments,
                    roleAssignments => roleAssignments.Include(
                            item => item.PrincipalId,
                            item => item.Member.LoginName,
                            item => item.Member.PrincipalType,
                            item => item.RoleDefinitionBindings
                    ));
                    clientContext.ExecuteQuery();
                    SetRoleAssignments(list.RoleAssignments, listDict);
                }
                if (listsDict.ContainsKey(list.Id.ToString()))
                {
                    Console.WriteLine("Duplicate key " + list.Id);
                }
                else
                {
                    listsDict.Add(list.Id.ToString(), listDict);
                }
            }
            string listJsonPath = path + "\\lists.json";
            Console.WriteLine("Exported lists for site {0} to {1}", url, listJsonPath);
            System.IO.File.WriteAllText(listJsonPath, serializer.Serialize(listsDict));
        }

        private static void SetRoleAssignments(RoleAssignmentCollection roleAssignments, Dictionary<string, object> itemDict)
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
                roleAssignmentDict.Add("PrincipalType", roleAssignment.Member.PrincipalType.ToString());
                roleAssignmentDict.Add("RoleDefinitionIds", defs);
                roleAssignmentsDict.Add(roleAssignment.Member.LoginName, roleAssignmentDict);
            }
            itemDict.Add("RoleAssignments", roleAssignmentsDict);
        }

        static public List<Dictionary<string, object>> IndexFolder(ClientContext clientContext, Folder folder)
        {
            List<Dictionary<string, object>> files = new List<Dictionary<string, object>>();
            foreach (File file in folder.Files)
            {
                Dictionary<string, object> fileDict = new Dictionary<string, object>();
                fileDict.Add("Title", file.Title);
                fileDict.Add("FileType", "file");
                fileDict.Add("Name", file.Name);
                fileDict.Add("TimeCreated", file.TimeCreated);
                fileDict.Add("TimeLastModified", file.TimeLastModified);
                // TODO: how do i get the author info to return? it gives me error when I try to get it.
                // fileDict.Add("Author.LoginName", file.Author.LoginName);
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
    }
}