using System;
namespace SpPrefetchIndexBuilder {
  public class Util {
    public static string addSlashToUrlIfNeeded(string siteUrl) {
      if (siteUrl.EndsWith("/", StringComparison.CurrentCulture)) {
        siteUrl = siteUrl.Substring(0, siteUrl.Length - 1);
      }
      return siteUrl;
    }
  }
}
