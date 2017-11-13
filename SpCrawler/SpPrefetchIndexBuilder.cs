using System;
using System.Net;
using System.Diagnostics;
using System.Security;
using System.Collections.Generic;
using System.Web.Script.Serialization;
using Microsoft.SharePoint.Client;

namespace SpPrefetchIndexBuilder
{
    class SpPrefetchIndexBuilder
    {
        public static string defaultSite = "http://localhost/";
        public static CredentialCache cc = null;
        public static string site = defaultSite;
        public static JavaScriptSerializer serializer = new JavaScriptSerializer();
        public static string baseDir;

        static void Main(string[] args)
        {
            Stopwatch sw = Stopwatch.StartNew();
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
                    spPassword = args[4];
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
            getSubWebs(site, baseDir + "\\" + Guid.NewGuid().ToString());
            Console.WriteLine("Export complete. Took {0} milliseconds.", sw.ElapsedMilliseconds);
        }

        public static ClientContext getClientContext(string site)
        {
            ClientContext clientContext = new ClientContext(site);
            if (cc != null)
            {
                clientContext.Credentials = cc;
            }
            return clientContext;
        }

        public static void getSubWebs(string url, string parentPath)
        {
            try
            {
                ClientContext clientContext = getClientContext(url);
                Web oWebsite = clientContext.Web;
                clientContext.Load(oWebsite, website => website.Webs, website => website.Title, website => website.Url, website => website.RoleDefinitions, website => website.RoleAssignments, website => website.HasUniqueRoleAssignments, website => website.Description, website => website.Id);
                clientContext.ExecuteQuery();
                string path = parentPath + "\\" + oWebsite.Title;
                DownloadWeb(clientContext, oWebsite, url, path);
                foreach (Web orWebsite in oWebsite.Webs)
                {
                    getSubWebs(orWebsite.Url, path);

                }
            }
            catch (Exception ex)
            {

                Console.WriteLine(ex.ToString());
            }
        }

        static void DownloadWeb(ClientContext clientContext, Web web, string url, string path)
        {
            Console.WriteLine("Exporting site {0}", url);
            System.IO.Directory.CreateDirectory(path);
            Dictionary<string, object> webDict = new Dictionary<string, object>();
            webDict.Add("Title", web.Title);
            webDict.Add("Id", web.Id);
            webDict.Add("Description", web.Description);
            webDict.Add("Url", url);
            if (web.HasUniqueRoleAssignments)
            {
                List<object[]> roleDefArray = new List<object[]>();
                foreach (RoleDefinition roleDefition in web.RoleDefinitions)
                {
                    roleDefArray.Add(new object[] { roleDefition.Id, roleDefition.Name, roleDefition.RoleTypeKind});
                }
                webDict.Add("RoleDefinitions", roleDefArray);
                clientContext.Load(web.RoleAssignments,
                    roleAssignment => roleAssignment.Include(
                            item => item.PrincipalId,
                            item => item.Member.LoginName,
                            item => item.RoleDefinitionBindings
                        ));
                clientContext.ExecuteQuery();
                Dictionary<string, object> roleAssignmentDict = new Dictionary<string, object>();                
                foreach (RoleAssignment roleAssignment in web.RoleAssignments)
                {
                    List<int> defs = new List<int>();
                    foreach (RoleDefinition roleDefinition in roleAssignment.RoleDefinitionBindings)
                    {
                        defs.Add(roleDefinition.Id);
                    }
                    roleAssignmentDict.Add(roleAssignment.Member.LoginName, defs);
                }
                webDict.Add("RoleAssignments", roleAssignmentDict);
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
                if (list.Title.Equals("Composed Looks") || list.Title.Equals("Master Page Gallery") || list.Title.Equals("Site Assets") || list.Title.Equals("Site Pages"))
                {
                    continue;
                }
                Dictionary <string, object> listDict = new Dictionary<string, object>();
                listDict.Add("Id", list.Id);
                listDict.Add("Title", list.Title);
                CamlQuery camlQuery = new CamlQuery();
                camlQuery.ViewXml = "<View Scope=\"RecursiveAll\"></View>";
                ListItemCollection collListItem = list.GetItems(camlQuery);
                clientContext.Load(collListItem);
                clientContext.Load(collListItem,
                  items => items.Include(
                     item => item.Id,
                     item => item.DisplayName,
                     item => item.HasUniqueRoleAssignments));
                
                clientContext.ExecuteQuery();
                foreach (ListItem listItem in collListItem)
                {
                    if (listItem.HasUniqueRoleAssignments)
                    {
                        clientContext.Load(listItem.RoleAssignments,
                            roleAssignments => roleAssignments.Include(
                                    item => item.PrincipalId,
                                    item => item.Member.LoginName,
                                    item => item.RoleDefinitionBindings
                            ));
                    }
                }
                clientContext.Load(list.RootFolder.Files);
                clientContext.Load(list.RootFolder.Folders);
                clientContext.ExecuteQuery();
                List<Dictionary<string, object>> itemsList = new List<Dictionary<string, object>>();
                foreach (ListItem listItem in collListItem)
                {
                    Dictionary<string, object> itemDict = new Dictionary<string, object>();
                    itemDict.Add("DisplayName", listItem.DisplayName);
                    itemDict.Add("Id", listItem.Id);
                    Dictionary<string, object> roleAssignmentDict = new Dictionary<string, object>();
                    if (listItem.HasUniqueRoleAssignments)
                    {
                        foreach (RoleAssignment roleAssignment in listItem.RoleAssignments)
                        {
                            List<object> permissions = new List<object>();
                            foreach (RoleDefinition roleDefinition in roleAssignment.RoleDefinitionBindings)
                            {
                                permissions.Add(roleDefinition.Id);
                            }
                            roleAssignmentDict.Add(roleAssignment.Member.LoginName, permissions);
                        }
                    }
                    itemDict.Add("RoleAssignments", roleAssignmentDict);
                    itemDict.Add("FieldValues", listItem.FieldValues);
                    if (listItem.FieldValues.ContainsKey("Attachments") && (bool)listItem.FieldValues["Attachments"])
                    {
                        clientContext.Load(listItem.AttachmentFiles);
                        clientContext.ExecuteQuery();
                        List<Dictionary<string, object>> attachmentFileList = new List<Dictionary<string, object>>();
                        foreach (Attachment attachmentFile in listItem.AttachmentFiles)
                        {
                            Dictionary<string, object> attachmentFileDict = new Dictionary<string, object>();
                            attachmentFileDict.Add("ServerRelativeUrl", attachmentFile.ServerRelativeUrl);
                            attachmentFileDict.Add("FileName", attachmentFile.FileName);
                            attachmentFileList.Add(attachmentFileDict);
                        }
                        itemDict.Add("AttachmentFiles", attachmentFileList);
                    }
                    itemsList.Add(itemDict);
                }
                listDict.Add("Items", itemsList);
                clientContext.Load(list.RootFolder);
                clientContext.Load(list.RootFolder.Files);
                clientContext.Load(list.RootFolder.Folders);
                clientContext.ExecuteQuery();
                listDict.Add("Files", IndexFolder(clientContext, list.RootFolder));
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
