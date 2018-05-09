using System;
using System.Collections.Generic;

namespace SpPrefetchIndexBuilder {
  class WebToFetch {
    public String url;
    public String rootLevelSiteUrl;
    public Dictionary<string, object> webDict;
    public bool isRootLevelSite;
  }
}
