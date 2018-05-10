using System.Collections.Generic;
using Microsoft.SharePoint.Client;

namespace SpPrefetchIndexBuilder {
  public class ChangeOutput {
    public string site;
    public Change change;
    public Dictionary<string, object> changeDict;
  }
}
