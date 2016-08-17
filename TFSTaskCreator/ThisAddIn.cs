using Microsoft.TeamFoundation.Client;
using Microsoft.TeamFoundation.WorkItemTracking.Client;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Configuration;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Net;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Xml.Linq;
using Outlook = Microsoft.Office.Interop.Outlook;
using Office = Microsoft.Office.Core;
using Microsoft.Office.Interop.Outlook;
using System.IO;

namespace TFSTaskCreator
{
    public partial class ThisAddIn
    {
        Outlook.NameSpace outlookNameSpace;
        Outlook.MAPIFolder inbox;
        Outlook.Items items;
        static string tfsUrl = ConfigurationManager.AppSettings["tfsUrl"];
        static string userName = ConfigurationManager.AppSettings["userName"];
        static string password = ConfigurationManager.AppSettings["password"];
        static string domain = ConfigurationManager.AppSettings["domain"];
        string project = ConfigurationManager.AppSettings["project"];
        string areaPath = ConfigurationManager.AppSettings["areaPath"];
        string assignedTo = ConfigurationManager.AppSettings["assignedTo"];
        string configTitle = ConfigurationManager.AppSettings["workItemTitle"];
        string filter = ConfigurationManager.AppSettings["filter"];
        string iterationPath = ConfigurationManager.AppSettings["iterationPath"];
        string team = ConfigurationManager.AppSettings["team"];
        string backlogPriority = ConfigurationManager.AppSettings["backlogPriority"];
        string filter2 = @"""Development"" assignment";
        string filter3 = @"""IT"" assignment";

        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            try
            {
                outlookNameSpace = this.Application.GetNamespace("MAPI");
                inbox = outlookNameSpace.GetDefaultFolder(
                        Microsoft.Office.Interop.Outlook.
                        OlDefaultFolders.olFolderInbox);

                items = inbox.Items;
                items.ItemAdd +=
                    new Outlook.ItemsEvents_ItemAddEventHandler(filter_Item);
            }
            catch (System.Exception ex)
            {
                Helper.Log("\n Exception : " + ex);
            }
        }

        void filter_Item(object Item)
        {
            try
            {
                Outlook.MailItem mail = (Outlook.MailItem)Item;
                if (Item != null)
                {
                    if (mail.Subject.ToUpper().StartsWith(filter.ToUpper()) ||
                        mail.Subject.ToUpper().StartsWith(filter2.ToUpper()) ||
                        mail.Subject.ToUpper().StartsWith(filter3.ToUpper()))
                    {  
                        // Create TFS work Item
                        CreateTFSWorkItem(mail);
                    }
                }
            }
            catch (System.Exception ex)
            {
                Helper.Log(ex);
            }
        }

        private void CreateTFSWorkItem(MailItem mailItem)
        {
            try
            {
                var tfs = GetTFSProject();

                if (tfs != null)
                {
                    WorkItemStore store = (WorkItemStore)tfs.GetService(typeof(WorkItemStore));
                    WorkItemTypeCollection workItemTypes = store.Projects[project].WorkItemTypes;

                    var workItemDetails = checkPBIExists(mailItem.Subject, tfs);

                    if (workItemDetails != null && workItemDetails.Count == 0)
                    {
                        var taskType = workItemTypes["Product Backlog Item"];
                        var task = new WorkItem(taskType);
                        task.Title = configTitle + mailItem.Subject;
                        task.AreaPath = areaPath;
                        task.Fields["Assigned To"].Value = assignedTo;
                        task.IterationPath = iterationPath;
                        task.Fields["Team"].Value = team;
                        task.Fields["Backlog Priority"].Value = backlogPriority;
                        task.Description = mailItem.HTMLBody;

                        // Attachment code
                        if (mailItem != null)
                        {
                            if (mailItem.Attachments.Count > 0)
                            {
                                for (int i = 1; i <= mailItem
                                   .Attachments.Count; i++)
                                {
                                    var fileName = mailItem.Attachments[i].FileName;
                                    var path = Path.Combine(Path.GetTempPath(), fileName);

                                    if (!mailItem.HTMLBody.Contains(fileName))
                                    {
                                        mailItem.SaveAs(path, OlSaveAsType.olMSG);
                                        task.Attachments.Add(new Microsoft.TeamFoundation.WorkItemTracking.Client.Attachment(path));
                                    }
                                }
                            }
                        }

                        var taskDetails = task.Validate();
                        if (taskDetails.Count == 0)
                        {
                            task.Save();
                            Helper.Log("\n TFS WorkItem created. \n ");
                        }
                    }
                }
                else
                {
                    Helper.Log("\n Unable to connect to TFS \n ");
                }
            }
            catch (System.Exception exc)
            {
                Helper.Log(exc);
            }
        }

        private static TfsTeamProjectCollection GetTFSProject()
        {
            NetworkCredential cred = new NetworkCredential(userName, password, domain);
            Uri tfsUri = new Uri(tfsUrl);
            var tfs = new TfsTeamProjectCollection(tfsUri, cred);
            return tfs;
        }

        private WorkItemCollection checkPBIExists(string workItemTitle, TfsTeamProjectCollection tfs)
        {
            try
            {
                WorkItemStore wiStore = new WorkItemStore(tfs);
                var queryText = @"SELECT [System.Id], 
                                    [System.Title], 
                                    [Microsoft.VSTS.Common.BacklogPriority], 
                                    [System.AssignedTo], 
                                    [System.State], 
                                    [Microsoft.VSTS.Scheduling.RemainingWork], 
                                    [Microsoft.VSTS.CMMI.Blocked], 
                                    [System.WorkItemType] 
                                FROM WorkItems 
                                WHERE [System.TeamProject] = @project 
                                    AND [System.WorkItemType] IN ('Product Backlog Item')
                                    AND [System.State] <> 'Removed'
                                    AND [System.Title] Contains @workItemTitle         
                                    ORDER BY [Microsoft.VSTS.Common.Priority], [System.Id] ";

                Dictionary<string, object> parameters = new Dictionary<string, object>();
                parameters.Add("project", project);
                parameters.Add("workItemTitle", workItemTitle);

                var query = new Query(wiStore, queryText, parameters);
                var workItem = query.RunQuery();

                return workItem;
            }
            catch
            {
                return null;
            }
        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
        }

        #region VSTO generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(ThisAddIn_Startup);
            this.Shutdown += new System.EventHandler(ThisAddIn_Shutdown);
        }

        #endregion
    }
}
