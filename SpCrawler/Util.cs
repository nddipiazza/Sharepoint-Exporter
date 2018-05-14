using System;
using System.IO;
using System.Reflection;

namespace SpPrefetchIndexBuilder {
  public class Util {
    public static string addSlashToUrlIfNeeded(string siteUrl) {
      string res = siteUrl;
      if (!res.EndsWith("/", StringComparison.CurrentCulture)) {
        res += "/";
      }
      return res;
    }
    public static string AssemblyDirectory {
      get {
        string codeBase = Assembly.GetExecutingAssembly().CodeBase;
        UriBuilder uri = new UriBuilder(codeBase);
        string path = Uri.UnescapeDataString(uri.Path);
        return Path.GetDirectoryName(path);
      }
    }
    public static string getBaseUrl(string siteUrl) {
      return new Uri(siteUrl).Scheme + "://" + new Uri(siteUrl).Host;
    }

    public static int getBaseUrlPort(string siteUrl) {
      return new Uri(siteUrl).Port;
    }

    public static string getBaseUrlHost(string siteUrl) {
      return new Uri(siteUrl).Host;
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
  }
}
