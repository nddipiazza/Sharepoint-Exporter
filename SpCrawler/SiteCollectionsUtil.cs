using System;
using System.Collections.Generic;
using System.IO;
using System.Net;
using System.Xml;

namespace SpPrefetchIndexBuilder {
  public class SiteCollectionsUtil {

    private CredentialCache credentialCache;
    private string rootSite;

    public SiteCollectionsUtil(CredentialCache credentialCache, string rootSite) {
      this.rootSite = rootSite;
      this.credentialCache = credentialCache;
    }

    public XmlDocument getContent(string siteUrl, String contentType, string contentId) {
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
      using (WebResponse webResponse = webRequest.EndGetResponse(asyncResult)) {
        using (StreamReader rd = new StreamReader(webResponse.GetResponseStream())) {
          contentDatabaseResult.Load(rd);
        }
      }
      return contentDatabaseResult;
    }

    public List<string> GetAllSiteCollections() {
      List<String> allSites = new List<string>();
      XmlDocument virtualServerGetContentResult = getContent(rootSite, "VirtualServer", null);
      XmlNode contentResultNode = virtualServerGetContentResult.SelectSingleNode("//*[local-name() = 'GetContentResult']");
      if (contentResultNode == null || contentResultNode.InnerText == null) {
        throw new Exception(string.Format("Cannot list top level sites from {0}", rootSite));
      }
      XmlDocument innerXmlDoc = new XmlDocument();
      innerXmlDoc.LoadXml(contentResultNode.InnerText);
      string contentDatabaseId = innerXmlDoc.SelectSingleNode("//*[local-name() = 'ContentDatabase']").Attributes["ID"].Value;
      if (contentDatabaseId == null) {
        throw new Exception(string.Format("Cannot list top level sites from {0}", rootSite));
      }
      XmlDocument contentDatabaseGetContentResult = getContent(rootSite, "ContentDatabase", contentDatabaseId);
      XmlNode contentDatabaseResultNode = contentDatabaseGetContentResult.SelectSingleNode("//*[local-name() = 'GetContentResult']");
      if (contentDatabaseResultNode == null || contentDatabaseResultNode.InnerText == null) {
        throw new Exception(string.Format("Cannot list top level sites from {0}", rootSite));
      }
      innerXmlDoc = new XmlDocument();
      innerXmlDoc.LoadXml(contentDatabaseResultNode.InnerText);
      XmlNodeList sites = innerXmlDoc.SelectNodes("//*[local-name() = 'Site']");
      foreach (XmlNode siteNode in sites) {
        allSites.Add(siteNode.Attributes["URL"].Value);
      }
      return allSites;
    }

    HttpWebRequest CreateWebRequest(string url, string action) {
      HttpWebRequest webRequest = (HttpWebRequest)WebRequest.Create(url);
      webRequest.Headers.Add("SOAPAction", action);
      webRequest.ContentType = "text/xml;charset=\"utf-8\"";
      webRequest.Accept = "text/xml";
      webRequest.Method = "POST";
      webRequest.Credentials = credentialCache;
      return webRequest;
    }

    XmlDocument CreateSoapEnvelope(string objectType, string objectId) {
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
      soapEnvelopeDocument.LoadXml(string.Format(soapEnv, objectType,
                                                 objectId == null ? "" : "<soap:objectId>" + objectId + "</soap:objectId>"));
      return soapEnvelopeDocument;
    }

    void InsertSoapEnvelopeIntoWebRequest(XmlDocument soapEnvelopeXml, HttpWebRequest webRequest) {
      using (Stream stream = webRequest.GetRequestStream()) {
        soapEnvelopeXml.Save(stream);
      }
    }
  }
}
