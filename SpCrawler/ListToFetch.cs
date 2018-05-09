using System;
using System.Collections.Generic;

namespace SpPrefetchIndexBuilder {
  class ListToFetch {
    public String site;
    public Guid listId;
    public Dictionary<string, object> listsDict = new Dictionary<string, object>();
  }
}
