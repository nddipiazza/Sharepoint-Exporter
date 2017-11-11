using System;
using System.Net;
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
            serializer.MaxJsonLength = 209715200;
            if (args.Length >= 2 && (args[0].Equals("--help") || args[0].Equals("-help") || args[0].Equals("/help") || args.Length > 4 || args.Length == 3))
            {
                Console.WriteLine("USAGE: SpPrefetchIndexBuilder.exe [siteUrl] [outputDir] [domain] [username]");
            }
            site = args.Length > 0 ? args[0] : defaultSite;
            baseDir = args.Length > 1 ? args[1] : System.IO.Directory.GetCurrentDirectory();
            if (site.EndsWith("/"))
            {
                site = site.Substring(0, site.Length - 1);
            }
            if (args.Length > 2)
            {
                Console.WriteLine("Please enter password for {0}", args[3]);
                cc = new CredentialCache();
                String spPassword = Environment.GetEnvironmentVariable("SP_PWD");
                NetworkCredential nc = spPassword == null ? new NetworkCredential(args[3], GetPassword(), args[2]) : new NetworkCredential(args[3], spPassword, args[2]);
                cc.Add(new Uri(site), "NTLM", nc);
            }
            getSubWebs(site, baseDir + "\\" + Guid.NewGuid().ToString());
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
                    string newUrl = site + orWebsite.ServerRelativeUrl;
                    getSubWebs(newUrl, path);

                }
            }
            catch (Exception ex)
            {

                Console.WriteLine(ex.ToString());
            }
        }

        static void DownloadWeb(ClientContext clientContext, Web web, string url, string path)
        {
            System.IO.Directory.CreateDirectory(path);
            Dictionary<string, object> siteDict = new Dictionary<string, object>();
            siteDict.Add("Title", web.Title);
            siteDict.Add("Id", web.Id);
            siteDict.Add("Description", web.Description);
            siteDict.Add("Url", url);
            if (web.HasUniqueRoleAssignments)
            {
                List<object[]> roleDefArray = new List<object[]>();
                foreach (RoleDefinition roleDefition in web.RoleDefinitions)
                {
                    roleDefArray.Add(new object[] { roleDefition.Id, roleDefition.Name, roleDefition.RoleTypeKind});
                }
                siteDict.Add("RoleDefinitions", roleDefArray);
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
                siteDict.Add("RoleAssignments", roleAssignmentDict);
            }
            System.IO.File.WriteAllText(path + "\\site.json", serializer.Serialize(siteDict));

            ListCollection lists = web.Lists;

            clientContext.Load(lists);
            clientContext.ExecuteQuery();

            Dictionary<string, object> listsDict = new Dictionary<string, object>();
            foreach (List list in lists)
            {
                Dictionary<string, object> listDict = new Dictionary<string, object>();
                listDict.Add("Id", list.Id);
                listDict.Add("Title", list.Title);
                CamlQuery camlQuery = new CamlQuery();
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
                    itemsList.Add(itemDict);
                }
                listDict.Add("Items", itemsList);
                Dictionary<string, object> filesDict = new Dictionary<string, object>();
                foreach (Microsoft.SharePoint.Client.File file in list.RootFolder.Files)
                {
                    filesDict.Add(file.Name, file.Properties.FieldValues);
                }
                listDict.Add("Files", filesDict);
                if (listsDict.ContainsKey(list.Id.ToString()))
                {
                    Console.WriteLine("Duplicate key " + list.Id);
                }
                else
                {
                    listsDict.Add(list.Id.ToString(), listDict);
                }
                
            }
            System.IO.File.WriteAllText(path + "\\lists.json", serializer.Serialize(listsDict));
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
