using System;
using System.Collections.Generic;

namespace SpPrefetchIndexBuilder {
  class WebToFetch {
    public String url;
    public String parentSiteUrl;
    public String topLevelSiteUrl;
    public Dictionary<string, object> webDict;
    public bool isTopLevelSite;
  }
}
