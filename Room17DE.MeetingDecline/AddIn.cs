using System;
using System.Collections.Generic;
using System.Threading;
using Microsoft.Office.Interop.Outlook;
using Room17DE.MeetingDecline.Util;

namespace Room17DE.MeetingDecline
{
    public partial class AddIn
    {
        private Folders DeletedItemsFolder;
        internal static int[] SystemFoldersIDs;
        private CancellationTokenSource CancelationTokenSource = new CancellationTokenSource();

        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            #region settings
            // check for settings upgrade after outlook update
            if(Properties.Settings.Default.UpgradeRequired)
            {
                Properties.Settings.Default.Upgrade();
                Properties.Settings.Default.UpgradeRequired = false;
                Properties.Settings.Default.Save();
            }

            // check if we have settings
            if (Properties.Settings.Default.MeetingDeclineRules == null)
            {
                Properties.Settings.Default.MeetingDeclineRules = new Dictionary<string, DeclineRule>();
                Properties.Settings.Default.Save();
            }
            if (Properties.Settings.Default.LastMailCheck == null)
            {
                Properties.Settings.Default.LastMailCheck = new Dictionary<string, DateTime>();
                Properties.Settings.Default.Save();
            }
            #endregion

            // make sure that a deleted folder removes the meetingdecline rule 
            Folder deletedItemsFolder = (Folder)Application.Session.GetDefaultFolder(OlDefaultFolders.olFolderDeletedItems);
            DeletedItemsFolder = deletedItemsFolder.Folders; // keep a reference at class level so it wont be GCed and event handler lost
            DeletedItemsFolder.FolderAdd += DeletedItems_FolderAdd;

            // enumerate non user folder entry ids for later
            #region hashes
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
            #endregion

            // timer thread for handling new mails in folders
            NewMailPeriodicTask.Run(new TimeSpan(0, 1 ,0), CancelationTokenSource.Token); // TODO: make interval configurable?

            // register to process exit event for cleanup
            AppDomain.CurrentDomain.ProcessExit += Addin_ProcessExit;
        }

        /// <summary>
        /// Event handler for folder addition in another folder. In our case, it's triggered when 
        /// a folder has been moved to Deleted Items (aka trash)
        /// </summary>
        private void DeletedItems_FolderAdd(MAPIFolder Folder)
        {
            // check if folder exists in settings, then remove it and save
            if (Properties.Settings.Default.MeetingDeclineRules.ContainsKey(Folder.EntryID))
            {
                Properties.Settings.Default.MeetingDeclineRules.Remove(Folder.EntryID);
                Properties.Settings.Default.Save();
            }

            // same for checked map too
            if (Properties.Settings.Default.LastMailCheck.ContainsKey(Folder.EntryID))
            {
                Properties.Settings.Default.LastMailCheck.Remove(Folder.EntryID);
                Properties.Settings.Default.Save();
            }
        }

        /// <summary>
        /// Activates this application Ribbon
        /// </summary>
        protected override Microsoft.Office.Core.IRibbonExtensibility CreateRibbonExtensibilityObject() => new Ribbon();

        #region shutdown
        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
            // Note: Outlook no longer raises this event. If you have code that 
            //    must run when Outlook shuts down, see https://go.microsoft.com/fwlink/?LinkId=506785
            AddinShutdown();
        }

        /// <summary>
        /// Event handler for AppDomain process exit
        /// </summary>
        private void Addin_ProcessExit(object sender, EventArgs e)
        {
            AddinShutdown();
        }

        /// <summary>
        /// Method that needs to be called when Outlook is shutting down
        /// </summary>
        private void AddinShutdown()
        {
            CancelationTokenSource.Cancel();
        }
        #endregion

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
