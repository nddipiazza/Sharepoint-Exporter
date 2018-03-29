using System;
using Microsoft.SharePoint.Client;


namespace SpPrefetchIndexBuilder {
  public class SharepointChanges {

    // The change tokens used below are described by the following:

    // A change token is a delimited string with the following parts in the following order:
    // * Version number
    // * A number indicating the change scope: 0 – Content Database, 1 – site collection, 2 – site (aka web), 3 – list.
    // * GUID representing the scope ID of the change token
    // * Time(in UTC) when the change occurred
    // * Number of the change relative to other changes

    public static ChangeCollection GetSecurityChanges(ClientContext ctx, Site site, DateTime since) {
      ChangeQuery siteCQ = new ChangeQuery(false, false);
      siteCQ.ChangeTokenStart = new ChangeToken();
      siteCQ.ChangeTokenStart.StringValue = string.Format("1;1;{0};{1};-1", site.Id, since.ToUniversalTime().Ticks.ToString());
      siteCQ.ChangeTokenEnd = new ChangeToken();
      siteCQ.ChangeTokenEnd.StringValue = string.Format("1;1;{0};{1};-1", site.Id, since.ToUniversalTime().Ticks.ToString());
      siteCQ.GroupMembershipAdd = true;
      siteCQ.GroupMembershipDelete = true;
      siteCQ.RoleAssignmentAdd = true;
      siteCQ.RoleAssignmentDelete = true;
      siteCQ.RoleDefinitionUpdate = true;
      siteCQ.RoleDefinitionAdd = true;
      siteCQ.RoleDefinitionDelete = true;
      siteCQ.User = true;
      siteCQ.SystemUpdate = true;
      siteCQ.Group = true;
      var siteChanges = site.GetChanges(siteCQ);
      ctx.Load(siteChanges);
      ctx.ExecuteQuery();
      return siteChanges;
    }

    public static ChangeCollection GetSecurityChanges(ClientContext ctx, Web web, DateTime since) {
      ChangeQuery siteCQ = new ChangeQuery(false, false);
      siteCQ.ChangeTokenStart = new ChangeToken();
      siteCQ.ChangeTokenStart.StringValue = string.Format("1;2;{0};{1};-1", web.Id, since.ToUniversalTime().Ticks.ToString());
      siteCQ.ChangeTokenEnd = new ChangeToken();
      siteCQ.ChangeTokenEnd.StringValue = string.Format("1;2;{0};{1};-1", web.Id, since.ToUniversalTime().Ticks.ToString());
      siteCQ.GroupMembershipAdd = true;
      siteCQ.GroupMembershipDelete = true;
      siteCQ.RoleAssignmentAdd = true;
      siteCQ.RoleAssignmentDelete = true;
      siteCQ.RoleDefinitionUpdate = true;
      siteCQ.RoleDefinitionAdd = true;
      siteCQ.RoleDefinitionDelete = true;
      siteCQ.User = true;
      siteCQ.SystemUpdate = true;
      siteCQ.Group = true;
      var siteChanges = web.GetChanges(siteCQ);
      ctx.Load(siteChanges);
      ctx.ExecuteQuery();
      return siteChanges;
    }

    public static ChangeCollection GetSecurityChanges(ClientContext ctx, List list, DateTime since) {
      ChangeQuery siteCQ = new ChangeQuery(false, false);
      siteCQ.ChangeTokenStart = new ChangeToken();
      siteCQ.ChangeTokenStart.StringValue = string.Format("1;3;{0};{1};-1", list.Id, since.ToUniversalTime().Ticks.ToString());
      siteCQ.ChangeTokenEnd = new ChangeToken();
      siteCQ.ChangeTokenEnd.StringValue = string.Format("1;3;{0};{1};-1", list.Id, since.ToUniversalTime().Ticks.ToString());
      siteCQ.GroupMembershipAdd = true;
      siteCQ.GroupMembershipDelete = true;
      siteCQ.RoleAssignmentAdd = true;
      siteCQ.RoleAssignmentDelete = true;
      siteCQ.RoleDefinitionUpdate = true;
      siteCQ.RoleDefinitionAdd = true;
      siteCQ.RoleDefinitionDelete = true;
      siteCQ.User = true;
      siteCQ.SystemUpdate = true;
      siteCQ.Group = true;
      var siteChanges = list.GetChanges(siteCQ);
      ctx.Load(siteChanges);
      ctx.ExecuteQuery();
      return siteChanges;
    }
  }
}
