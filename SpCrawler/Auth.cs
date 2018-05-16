using System;
using System.Net;
using System.Net.Http;
using System.Security;
using System.Text.RegularExpressions;
using Microsoft.SharePoint.Client;

namespace SpPrefetchIndexBuilder {
  public class Auth {
    public CredentialCache credentialsCache;
    public SharePointOnlineCredentials sharepointOnlineCredentials;
    public HttpClientHandler httpHandler;
    public Auth(string rootSite,
                bool isSharepointOnline,
                string domain,
                string username,
                string password,
                string authScheme) {
      if (!isSharepointOnline) {
        NetworkCredential networkCredential;
        if (password == null && username != null) {
          Console.WriteLine("Please enter password for {0}", username);
          networkCredential = new NetworkCredential(username, GetPassword(), domain);
        } else if (username != null) {
          networkCredential = new NetworkCredential(username, password, domain);
        } else {
          networkCredential = CredentialCache.DefaultNetworkCredentials;
        }
        credentialsCache = new CredentialCache();
        credentialsCache.Add(new Uri(rootSite), authScheme, networkCredential);
        CredentialCache credentialCache = new CredentialCache { { Util.getBaseUrlHost(rootSite), Util.getBaseUrlPort(rootSite), authScheme, networkCredential } };
        httpHandler = new HttpClientHandler() {
          CookieContainer = new CookieContainer(),
          Credentials = credentialCache.GetCredential(Util.getBaseUrlHost(rootSite), Util.getBaseUrlPort(rootSite), authScheme)
        };
      } else {
        SecureString securePassword = new SecureString();
        foreach (char c in password) {
          securePassword.AppendChar(c);
        }
        sharepointOnlineCredentials = new SharePointOnlineCredentials(username, securePassword);
        httpHandler = new HttpClientHandler();
        Uri rootSiteUri = new Uri(Util.getBaseUrl(rootSite));
        httpHandler.CookieContainer.SetCookies(rootSiteUri, sharepointOnlineCredentials.GetAuthenticationCookie(rootSiteUri));
      }

    }
    public HttpClient createHttpClient(int fileDownloadTimeoutSecs, int numRetries) {

      HttpClient httpClient = new HttpClient(new HttpRetryMessageHandler(httpHandler, numRetries));
      httpClient.Timeout = TimeSpan.FromSeconds(fileDownloadTimeoutSecs);
      return httpClient;
    }
    SecureString GetPassword() {
      var pwd = new SecureString();
      while (true) {
        ConsoleKeyInfo i = Console.ReadKey(true);
        if (i.Key == ConsoleKey.Enter) {
          break;
        } 
        if (i.Key == ConsoleKey.Backspace) {
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
