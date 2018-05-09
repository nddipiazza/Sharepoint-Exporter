using System;
using System.Collections.Generic;
using System.IO;
using System.Net;
using System.Security;
using System.Text;
using System.Web.Script.Serialization;

namespace SpPrefetchIndexBuilder {
  public class SharepointExporterConfig {

    public List<string> sites = new List<string>();
    public List<string> ignoreListNames = new List<string>();
    public string baseDir = null;
    public bool customBaseDir = false;
    public string rootSite = null;
    public int numThreads = 50;
    public bool excludeUsersAndGroups = false;
    public bool excludeGroupMembers = false;
    public bool excludeSubSites = false;
    public bool excludeLists = false;
    public bool excludeRoleDefinitions = false;
    public bool excludeRoleAssignments = false;
    public bool deleteExistingOutputDir = false;
    public bool excludeFiles = false;
    public int maxFiles = -1;
    public JavaScriptSerializer serializer = new JavaScriptSerializer();
    public int maxFileSizeBytes = -1;
    public int fileCount = 0;
    public NetworkCredential networkCredentials;

    public SharepointExporterConfig(string[] args) {
      ignoreListNames.Add("Cache Profiles");
      ignoreListNames.Add("Content and Structure Reports");
      ignoreListNames.Add("Content Organizer Rules");
      ignoreListNames.Add("Content type publishing error log");
      ignoreListNames.Add("Converted Forms");
      ignoreListNames.Add("Device Channels");
      ignoreListNames.Add("Drop Off Library");
      ignoreListNames.Add("Form Templates");
      ignoreListNames.Add("Hold Reports");
      ignoreListNames.Add("Holds");
      ignoreListNames.Add("Long Running Operation Status");
      ignoreListNames.Add("MicroFeed");
      ignoreListNames.Add("Notification List");
      ignoreListNames.Add("Project Policy Item List");
      ignoreListNames.Add("Quick Deploy Items");
      ignoreListNames.Add("Relationships List");
      ignoreListNames.Add("Reusable Content");
      ignoreListNames.Add("Site Collection Documents");
      ignoreListNames.Add("Site Collection Images");
      ignoreListNames.Add("Solution Gallery");
      ignoreListNames.Add("Style Library");
      ignoreListNames.Add("Submitted E-mail Records");
      ignoreListNames.Add("Suggested Content Browser Locations");
      ignoreListNames.Add("TaxonomyHiddenList");
      ignoreListNames.Add("Theme Gallery");
      ignoreListNames.Add("Translation Packages");
      ignoreListNames.Add("Translation Status");
      ignoreListNames.Add("User Information List");
      ignoreListNames.Add("Variation Labels");
      ignoreListNames.Add("Web Part Gallery");
      ignoreListNames.Add("wfpub");
      ignoreListNames.Add("Composed Looks");
      ignoreListNames.Add("Master Page Gallery");
      ignoreListNames.Add("Site Assets");
      ignoreListNames.Add("Site Pages");

      string spMaxFileSizeBytes = Environment.GetEnvironmentVariable("SP_MAX_FILE_SIZE_BYTES");
      if (spMaxFileSizeBytes != null) {
        maxFileSizeBytes = int.Parse(spMaxFileSizeBytes);
      }
      string spNumThreads = Environment.GetEnvironmentVariable("SP_NUM_THREADS");
      if (spNumThreads != null) {
        numThreads = int.Parse(spNumThreads);
      }
      serializer.MaxJsonLength = 1677721600;

      bool help = false;

      string spDomain = null;
      string spUsername = null;
      string spPassword = Environment.GetEnvironmentVariable("SP_PWD");
      baseDir = Directory.GetCurrentDirectory();
      customBaseDir = false;
      string sitesFilePath = null;

      foreach (string arg in args) {
        if (arg.Equals("--help") || arg.Equals("-help") || arg.Equals("/help")) {
          help = true;
          break;
        }
        if (arg.StartsWith("--sitesFile=", StringComparison.CurrentCulture)) {
          sitesFilePath = arg.Split(new Char[] { '=' })[1];
        } else if (arg.StartsWith("--sharepointUrl=", StringComparison.CurrentCulture)) {
          sites.Add(arg.Split(new Char[] { '=' })[1]);
        } else if (arg.StartsWith("--outputDir=", StringComparison.CurrentCulture)) {
          baseDir = arg.Split(new Char[] { '=' })[1];
          customBaseDir = true;
        } else if (arg.StartsWith("--deleteExistingOutputDir=", StringComparison.CurrentCulture)) {
          deleteExistingOutputDir = Boolean.Parse(arg.Split(new Char[] { '=' })[1]);
        } else if (arg.StartsWith("--domain=", StringComparison.CurrentCulture)) {
          spDomain = arg.Split(new Char[] { '=' })[1];
        } else if (arg.StartsWith("--username=", StringComparison.CurrentCulture)) {
          spUsername = arg.Split(new Char[] { '=' })[1];
        } else if (arg.StartsWith("--password=", StringComparison.CurrentCulture)) {
          spPassword = arg.Split(new Char[] { '=' })[1];
        } else if (arg.StartsWith("--numThreads=", StringComparison.CurrentCulture)) {
          numThreads = int.Parse(arg.Split(new Char[] { '=' })[1]);
        } else if (arg.StartsWith("--maxFileSizeBytes=", StringComparison.CurrentCulture)) {
          maxFileSizeBytes = int.Parse(arg.Split(new Char[] { '=' })[1]);
        } else if (arg.StartsWith("--maxFiles=", StringComparison.CurrentCulture)) {
          maxFiles = int.Parse(arg.Split(new Char[] { '=' })[1]);
        } else if (arg.StartsWith("--excludeUsersAndGroups=", StringComparison.CurrentCulture)) {
          excludeUsersAndGroups = Boolean.Parse(arg.Split(new Char[] { '=' })[1]);
        } else if (arg.StartsWith("--excludeGroupMembers=", StringComparison.CurrentCulture)) {
          excludeGroupMembers = Boolean.Parse(arg.Split(new Char[] { '=' })[1]);
        } else if (arg.StartsWith("--excludeSubSites=", StringComparison.CurrentCulture)) {
          excludeSubSites = Boolean.Parse(arg.Split(new Char[] { '=' })[1]);
        } else if (arg.StartsWith("--excludeLists=", StringComparison.CurrentCulture)) {
          excludeLists = Boolean.Parse(arg.Split(new Char[] { '=' })[1]);
        } else if (arg.StartsWith("--excludeRoleAssignments=", StringComparison.CurrentCulture)) {
          excludeRoleAssignments = Boolean.Parse(arg.Split(new Char[] { '=' })[1]);
        } else if (arg.StartsWith("--excludeRoleDefinitions=", StringComparison.CurrentCulture)) {
          excludeRoleDefinitions = Boolean.Parse(arg.Split(new Char[] { '=' })[1]);
        } else if (arg.StartsWith("--excludeFiles=", StringComparison.CurrentCulture)) {
          excludeFiles = Boolean.Parse(arg.Split(new Char[] { '=' })[1]);
        } else {
          Console.WriteLine("ERROR - Unrecognized argument {0}.", arg);
          help = true;
        }
      }

      if (spPassword == null && spUsername != null) {
        Console.WriteLine("Please enter password for {0}", spUsername);
        networkCredentials = new NetworkCredential(spUsername, GetPassword(), spDomain);
      } else if (spUsername != null) {
        networkCredentials = new NetworkCredential(spUsername, spPassword, spDomain);
      } else {
        networkCredentials = CredentialCache.DefaultNetworkCredentials;
      }

      if (sitesFilePath != null) {
        FileInfo sitesFile = new FileInfo(sitesFilePath);
        if (!sitesFile.Exists) {
          Console.WriteLine("Error - sites file {0} doesn't exist", sitesFilePath);
          Environment.Exit(1);
        }
        if (sitesFile != null && sitesFile.Exists) {
          foreach (string nextSite in File.ReadLines(sitesFile.FullName)) {
            sites.Add(nextSite);
          }
        }
      }

      if (sites.Count <= 0) {
        Console.WriteLine("ERROR - Must specify --sharepointUrl argument or a --sitesFile argument to specify what sharepoint sites to fetch.");
        help = true;
      }

      if (help) {
        Console.WriteLine(new StringBuilder().AppendLine("USAGE: SpPrefetchIndexBuilder.exe")
                          .AppendLine("    --sharepointUrl=[The sharepoint url. I.e. http://oursharepoint]   (*required)")
                          .AppendLine("    --incrementalFile=[optional - path to incremental file created during a previous run. if specified, will only fetch incremental changes based on this file.]")
                          .AppendLine("    --sitesFile=[optional - path to sites file. this is a list]")
                          .AppendLine("    --outputDir=[optional - where to save the output. default will use this directory.]")
                          .AppendLine("    --domain=[optional - netbios domain of the user to crawl as]")
                          .AppendLine("    --username=[optional - specify a username to crawl as. must specify domain if using this]")
                          .AppendLine("    --password=[password (not recommended, do not specify to be prompted or use SP_PWD environment variable)]")
                          .AppendLine("    --numThreads=[optional number of threads to use while fetching. Default 50]")
                          .AppendLine("    --excludeUsersAndGroups=[exclude users and groups from the top level site collections. default false]")
                          .AppendLine("    --excludeGroupMembers=[exclude group members from the UsersAndGroups section. default false]")
                          .AppendLine("    --excludeRoleDefinitions=[if true will not store obtain role definition metadata from the top level site collections. default false] ")
                          .AppendLine("    --excludeSubSites=[only output the top level sites, do not descend into sub-sites. default false]")
                          .AppendLine("    --excludeLists=[exclude lists from the results. default false]")
                          .AppendLine("    --excludeFiles=[Do not download the files from the results] ")
                          .AppendLine("    --excludeRoleAssignments=[if true will not store obtain role assignment metadata. default false] ")
                          .AppendLine("    --maxFileSizeBytes=[optional maximum file size. Must be > 0. Default unlimited]")
                          .AppendLine("    --maxFiles=[if > 0 will only download this many files before quitting. default -1]"));
        Environment.Exit(0);
      }

    }
    private SecureString GetPassword() {
      var pwd = new SecureString();
      while (true) {
        ConsoleKeyInfo i = Console.ReadKey(true);
        if (i.Key == ConsoleKey.Enter) {
          break;
        } else if (i.Key == ConsoleKey.Backspace) {
          if (pwd.Length > 0) {
            pwd.RemoveAt(pwd.Length - 1);
            Console.Write("\b \b");
          }
        } else {
          pwd.AppendChar(i.KeyChar);
          Console.Write("*");
        }
      }
      return pwd;
    }
  }
}
