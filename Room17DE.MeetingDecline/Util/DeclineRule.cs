using Microsoft.Office.Interop.Outlook;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;

namespace Room17DE.MeetingDecline.Util
{
    /// <summary>
    /// POCO used for saving rules to settings
    /// </summary>
    [Serializable]
    class DeclineRule
    {
        // default values
        private bool _isActive = false;
        private bool _sendNotification = false;
        private OlMeetingResponse _response = OlMeetingResponse.olMeetingDeclined;

        public bool IsActive { get => _isActive; set => _isActive = value; }
        public bool SendNotification { get => _sendNotification; set => _sendNotification = value; }
        public OlMeetingResponse Response { get => _response; set => _response = value; }
        public string Message { get; set; }
    }

    /// <summary>
    /// Class for communicating with WPF and present rules data
    /// </summary>
    class DeclineRuleController
    {
        /// <summary>
        /// Populates DeclineRules with entries from settings
        /// </summary>
        /// <returns>An ObservableCollection used for data binding in WPF</returns>
        public ObservableCollection<DeclineRuleRecord> LoadData()
        {
            IDictionary<string, DeclineRule> rules = Properties.Settings.Default.MeetingDeclineRules;
            ObservableCollection<DeclineRuleRecord> results = new ObservableCollection<DeclineRuleRecord>();

            // get all folders
            MAPIFolder root = Globals.AddIn.Application.Session.DefaultStore.GetRootFolder();
            IEnumerable<MAPIFolder> allFolders = GetFolders(root);

            string toRemove = "\\\\" + Globals.AddIn.Application.Session.CurrentUser.AddressEntry.GetExchangeUser().PrimarySmtpAddress + "\\";

            // show all folders
            foreach (MAPIFolder folder in allFolders)
            {
                bool isActive = false;
                bool sendResponse = false;
                bool isDecline = true;

                // get folder setting
                if (rules.ContainsKey(folder.EntryID))
                {
                    DeclineRule rule = rules[folder.EntryID];
                    isActive = rule.IsActive;
                    sendResponse = rule.SendNotification;
                    isDecline = rule.Response == OlMeetingResponse.olMeetingDeclined;
                }

                // add new rule to be displayed
                results.Add(new DeclineRuleRecord()
                {
                    FolderID = folder.EntryID,
                    FolderName = folder.Name,
                    FolderPath = folder.FolderPath.Replace(toRemove, ""),
                    IsActive = isActive,
                    IsDecline = isDecline,
                    SendResponse = sendResponse
                });
            }

            return results;
        }

        /// <summary>
        /// Gets all folders in Outlook recursively
        /// </summary>
        /// <param name="folder">Root folder to start searching from</param>
        /// <returns>List of all folders found starting from root</returns>
        private IEnumerable<MAPIFolder> GetFolders(MAPIFolder folder)
        {
            if (folder.DefaultItemType == OlItemType.olMailItem && Array.IndexOf(AddIn.SystemFoldersIDs, folder.EntryID.GetHashCode()) < 0)
                if (folder.Folders.Count == 0)
                    yield return folder;
                else
                {
                    foreach (MAPIFolder subFolder in folder.Folders)
                        foreach (MAPIFolder result in GetFolders(subFolder))
                            yield return result;
                }
        }
    }

    /// <summary>
    /// POCO for keeping a deline rule settings
    /// </summary>
    class DeclineRuleRecord
    {
        public string FolderID { get; set; }
        public string FolderName { get; set; }
        public string FolderPath { get; set; }
        public bool IsActive { get; set; }
        public bool IsDecline { get; set; }
        public bool SendResponse { get; set; }
    }
}
