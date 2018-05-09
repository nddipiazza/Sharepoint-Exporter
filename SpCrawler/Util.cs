using System;
using System.IO;
using System.Reflection;

namespace SpPrefetchIndexBuilder {
  public class Util {
    public static string addSlashToUrlIfNeeded(string siteUrl) {
      if (siteUrl.EndsWith("/", StringComparison.CurrentCulture)) {
        siteUrl = siteUrl.Substring(0, siteUrl.Length - 1);
      }
      return siteUrl;
    }
    public static string AssemblyDirectory {
      get {
        string codeBase = Assembly.GetExecutingAssembly().CodeBase;
        UriBuilder uri = new UriBuilder(codeBase);
        string path = Uri.UnescapeDataString(uri.Path);
        return Path.GetDirectoryName(path);
      }
    }
  }
}
