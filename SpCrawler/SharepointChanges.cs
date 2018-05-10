using System;
using System.Collections.Generic;
using Microsoft.SharePoint.Client;

namespace SpPrefetchIndexBuilder {
  public class SharepointChanges {

    //public List<ChangeWeb> changeWebList = new List<ChangeWeb>();
    //public List<ChangeSite> changeSiteList = new List<ChangeSite>();
    //public List<ChangeList> changeListList = new List<ChangeList>();
    //public List<ChangeItem> changeItemList = new List<ChangeItem>();
    //public List<ChangeFile> changeFileList = new List<ChangeFile>();
    //public List<ChangeGroup> changeGroupList = new List<ChangeGroup>();
    //public List<ChangeUser> changeUserList = new List<ChangeUser>();
    public List<ChangeOutput> changeOutputs = new List<ChangeOutput>();

    // The change tokens used below are described by the following:

    // A change token is a delimited string with the following parts in the following order:
    // * Version number
    // * A number indicating the change scope: 0 – Content Database, 1 – site collection, 2 – site (also referred to as web), 3 – list.
    // * GUID representing the scope ID of the change token
    // * Time(in UTC) when the change occurred
    // * Number of the change relative to other changes

    public static ChangeCollection GetChanges(ClientContext ctx, Site site, DateTime since) {
      ChangeQuery cq = new ChangeQuery(true, true);
      cq.ChangeTokenStart = new ChangeToken();
      cq.ChangeTokenStart.StringValue = string.Format("1;1;{0};{1};-1", site.Id, since.ToUniversalTime().Ticks.ToString());
      cq.ChangeTokenEnd = new ChangeToken();
      cq.ChangeTokenEnd.StringValue = string.Format("1;1;{0};{1};-1", site.Id, DateTime.Now.AddDays(10).ToUniversalTime().Ticks.ToString());
      cq.GroupMembershipAdd = true;
      cq.GroupMembershipDelete = true;
      cq.RoleAssignmentAdd = true;
      cq.RoleAssignmentDelete = true;
      cq.RoleDefinitionUpdate = true;
      cq.RoleDefinitionAdd = true;
      cq.RoleDefinitionDelete = true;
      cq.User = true;
      cq.SystemUpdate = true;
      cq.Group = true;
      cq.SecurityPolicy = true;
      cq.Update = true;
      var changes = site.GetChanges(cq);
      ctx.Load(changes);
      ctx.ExecuteQuery();
      return changes;
    }

    public static ChangeCollection GetChanges(ClientContext ctx, Web web, DateTime since) {
      ChangeQuery cq = new ChangeQuery(true, true);
      cq.ChangeTokenStart = new ChangeToken();
      cq.ChangeTokenStart.StringValue = string.Format("1;2;{0};{1};-1", web.Id, since.ToUniversalTime().Ticks.ToString());
      cq.ChangeTokenEnd = new ChangeToken();
      cq.ChangeTokenEnd.StringValue = string.Format("1;2;{0};{1};-1", web.Id, DateTime.Now.AddDays(10).ToUniversalTime().Ticks.ToString());
      cq.GroupMembershipAdd = true;
      cq.GroupMembershipDelete = true;
      cq.RoleAssignmentAdd = true;
      cq.RoleAssignmentDelete = true;
      cq.RoleDefinitionUpdate = true;
      cq.RoleDefinitionAdd = true;
      cq.RoleDefinitionDelete = true;
      cq.User = true;
      cq.SystemUpdate = true;
      cq.Group = true;
      cq.SecurityPolicy = true;
      cq.Web = true;
      var changes = web.GetChanges(cq);
      ctx.Load(changes);
      ctx.ExecuteQuery();
      return changes;
    }

    public static ChangeCollection GetChanges(ClientContext ctx, List list, DateTime since) {
      ChangeQuery cq = new ChangeQuery(false, false);
      cq.ChangeTokenStart = new ChangeToken();
      cq.ChangeTokenStart.StringValue = string.Format("1;3;{0};{1};-1", list.Id, since.ToUniversalTime().Ticks.ToString());
      cq.ChangeTokenEnd = new ChangeToken();
      cq.ChangeTokenEnd.StringValue = string.Format("1;3;{0};{1};-1", list.Id, DateTime.Now.AddDays(10).ToUniversalTime().Ticks.ToString());
      cq.GroupMembershipAdd = true;
      cq.GroupMembershipDelete = true;
      cq.RoleAssignmentAdd = true;
      cq.RoleAssignmentDelete = true;
      cq.RoleDefinitionUpdate = true;
      cq.RoleDefinitionAdd = true;
      cq.RoleDefinitionDelete = true;
      cq.User = true;
      cq.SystemUpdate = true;
      cq.Group = true;
      var changes = list.GetChanges(cq);
      ctx.Load(changes);
      ctx.ExecuteQuery();
      return changes;
    }

    public void AddChangeToIncrementalDict(Dictionary<string, object> changesDict, string type, string ownerOfChangeUrl, Change change) {
      Dictionary<string, object> changeDict = new Dictionary<string, object>();
      ChangeOutput changeOutput = new ChangeOutput();
      changeOutput.change = change;
      changeOutput.site = ownerOfChangeUrl;
      changeOutput.changeDict = changeDict;
      changeOutputs.Add(changeOutput);
      if (change is ChangeGroup) {
        ChangeGroup changeGroup = (ChangeGroup)change;
        changeDict.Add("GroupId", changeGroup.GroupId);
        //changeGroupList.Add(changeGroup);
      } else if (change is ChangeUser) {
        ChangeUser changeUser = (ChangeUser)change;
        changeDict.Add("Activate", changeUser.Activate);
        changeDict.Add("UserId", changeUser.UserId);
        //changeUserList.Add(changeUser);
      } else if (change is ChangeItem) {
        ChangeItem changeItem = (ChangeItem)change;
        changeDict.Add("ItemId", changeItem.ItemId);
        changeDict.Add("ListId", changeItem.ListId);
        changeDict.Add("WebId", changeItem.WebId);
        //changeItemList.Add(changeItem);
      } else if (change is ChangeFolder) {
        ChangeFolder changeFolder = (ChangeFolder)change;
        changeDict.Add("UniqueId", changeFolder.UniqueId);
        changeDict.Add("WebId", changeFolder.WebId);
      } else if (change is ChangeList) {
        ChangeList changeList = (ChangeList)change;
        changeDict.Add("ListId", changeList.ListId);
        changeDict.Add("WebId", changeList.WebId);
        //changeListList.Add(changeList);
      } else if (change is ChangeFile) {
        ChangeFile changeFile = (ChangeFile)change;
        changeDict.Add("UniqueId", changeFile.UniqueId);
        changeDict.Add("WebId", changeFile.WebId);
        //changeFileList.Add(changeFile);
      } else if (change is ChangeWeb) {
        ChangeWeb changeWeb = (ChangeWeb)change;
        changeDict.Add("WebId", changeWeb.WebId);
        //changeWebList.Add(changeWeb);
      } else if (change is ChangeView) {
        return;
      } else {
        Console.WriteLine("Unhandled change type: {0}", change);
      }
      changeDict.Add("OwnerOfChangeType", type);
      changeDict.Add("OwnerOfChangeUrl", ownerOfChangeUrl);
      changeDict.Add("Tag", change.Tag);
      changeDict.Add("ChangeToken", change.ChangeToken);
      changeDict.Add("ChangeType", change.ChangeType.ToString());
      changeDict.Add("Time", change.Time);
      changeDict.Add("SiteId", change.SiteId);
      changeDict.Add("Type", change.GetType().Name);
      changesDict.Add(type + "|;" + ownerOfChangeUrl + "|;" + change.ChangeToken.StringValue, changeDict);
    }
  }
}
