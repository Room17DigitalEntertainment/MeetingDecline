using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using Outlook = Microsoft.Office.Interop.Outlook;
using Office = Microsoft.Office.Core;
using Microsoft.Office.Interop.Outlook;
using Room17.MeetingDecline.Util;
using Microsoft.Office.Tools.Ribbon;

namespace Room17.MeetingDecline
{
    public partial class AddIn
    {
        private Folders DeletedItemsFolder;
        internal static int[] SystemFoldersIDs;

        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            // event handler for new email
            this.Application.NewMailEx += Application_NewMailEx;

            // make sure that a deleted folder removes the meetingdecline rule 
            Folder deletedItemsFolder = (Folder)Application.Session.GetDefaultFolder(OlDefaultFolders.olFolderDeletedItems);
            DeletedItemsFolder = deletedItemsFolder.Folders; // keep a reference at class level so it wont be GCed and event handler lost
            DeletedItemsFolder.FolderAdd += DeletedItems_FolderAdd;

            // enumerate non user folder entry ids for later
            Array systemFolders = Enum.GetValues(typeof(OlDefaultFolders));
            string[] customFolders = new string[] { "Yammer Root", "Files", "Conversation History", "Social Activity Notifications", "Scheduled", "Quick Step Settings", "Archive", "Conversation Action Settings" };
            SystemFoldersIDs = new int[systemFolders.Length + customFolders.Length];
            int i;
            for (i = 0; i < systemFolders.Length; i++)
                try
                {
                    SystemFoldersIDs[i] = this.Application.Session.DefaultStore.GetDefaultFolder((OlDefaultFolders)systemFolders.GetValue(i))
                        .EntryID.GetHashCode();
                }
                catch { } // not all folders from OlDefaultFolders exist in outlook
            for (; i < customFolders.Length + systemFolders.Length; i++)
                try
                {
                    SystemFoldersIDs[i] = this.Application.Session.DefaultStore.GetRootFolder()
                        .Folders[customFolders[i - systemFolders.Length]].EntryID.GetHashCode();
                }
                catch { } // being a hardcoded list, we can't be 100% sure it always exists from app to app
        }

        /// <summary>
        /// Event handler for folder addition in another folder. In our case, it's triggered when 
        /// a folder has been moved to Deleted Items (aka trash)
        /// </summary>
        private void DeletedItems_FolderAdd(MAPIFolder Folder)
        {
            // check if we have settings
            if (Properties.Settings.Default.MeetingDeclineRules == null)
                return;

            // check if folder exists in settings, then remove it and save
            if (Properties.Settings.Default.MeetingDeclineRules.ContainsKey(Folder.EntryID))
            {
                Properties.Settings.Default.MeetingDeclineRules.Remove(Folder.EntryID);
                Properties.Settings.Default.Save();
            }
        }

        /// <summary>
        /// Event handler for every new email received, regardless of its type
        /// </summary>
        private void Application_NewMailEx(string EntryIDCollection)
        {
            // check if we have settings
            if (Properties.Settings.Default.MeetingDeclineRules == null)
                return;

            // get the meeting, if it's a meeting
            MeetingItem meetingItem = GetMeeting(EntryIDCollection);
            if (meetingItem == null)
                return;
            
            // get current meeting parent folder
            if (!(meetingItem.Parent is MAPIFolder parentFolder)) return;

            // check if parent folder is between settings
            if(Properties.Settings.Default.MeetingDeclineRules.ContainsKey(parentFolder.EntryID))
            {
                // check if rule it's active
                MeetingDeclineRule rule = Properties.Settings.Default.MeetingDeclineRules[parentFolder.EntryID];
                if (rule.IsActive)
                {
                    // if it's a Cancelation, delete it from calendar
                    if (meetingItem.Class == OlObjectClass.olMeetingCancellation)
                    {
                        meetingItem.Delete();
                        return;
                    }

                    // get associated appointment
                    AppointmentItem appointment = meetingItem.GetAssociatedAppointment(false);

                    // optional, send notification back to sender
                    appointment.ResponseRequested = rule.SendNotification;

                    // optional, add a meesage to the Body
                    if (!String.IsNullOrEmpty(rule.Message))
                        appointment.Body = rule.Message + Environment.NewLine + Environment.NewLine + appointment.Body;

                    // set decline to the meeting
                    MeetingItem responseMeeting = appointment.Respond(rule.Response, true);
                    // https://msdn.microsoft.com/en-us/VBA/Outlook-VBA/articles/appointmentitem-respond-method-outlook 
                    // says that Respond() will return a new meeting object for Tentative response

                    // send decline
                    (responseMeeting ?? meetingItem).Send(); 
                    (responseMeeting ?? meetingItem).Delete();
                }
            }
        }

        /// <summary>
        /// Get a MeetingItem based on EntryIDCollection, or null if it's not a meeting
        /// </summary>
        /// <param name="EntryIDCollection">The ID of the meeting</param>
        /// <returns>A MeetingItem that corresponds to the EntryID</returns>
        internal MeetingItem GetMeeting(string EntryIDCollection)
        {
            object item = null;
            try
            {
                item = Globals.AddIn.Application.Session.GetItemFromID(EntryIDCollection);
            }
            catch (System.Exception ex)
            {
                Logger.Error(String.Format("Could not retrieve info for message id:{0}. Error message is:{1}{2}",
                    EntryIDCollection, Environment.NewLine, ex.ToString()));
                return null;
            }

            MeetingItem meetingItem = item as MeetingItem;
            if (meetingItem == null)
            {
                Logger.Debug(String.Format("Message with id {0} is not a MeetingItem.", EntryIDCollection));
            }

            return meetingItem;
        }

        /// <summary>
        /// Activates this application Ribbon
        /// </summary>
        protected override Microsoft.Office.Core.IRibbonExtensibility CreateRibbonExtensibilityObject() => new Ribbon();

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
            // Note: Outlook no longer raises this event. If you have code that 
            //    must run when Outlook shuts down, see https://go.microsoft.com/fwlink/?LinkId=506785
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
