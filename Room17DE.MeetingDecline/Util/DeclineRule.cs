using Microsoft.Office.Interop.Outlook;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Collections.Specialized;
using System.ComponentModel;

namespace Room17DE.MeetingDecline.Util
{
    /// <summary>
    /// POCO used for saving rules to settings
    /// </summary>
    [Serializable]
    class DeclineRuleSetting
    {
        // default values
        private bool _isActive = false;
        private bool _sendResponse = false;
        private OlMeetingResponse _response = OlMeetingResponse.olMeetingDeclined;

        public bool IsActive { get => _isActive; set => _isActive = value; }
        public bool SendResponse { get => _sendResponse; set => _sendResponse = value; }
        public OlMeetingResponse Response { get => _response; set => _response = value; }
        public string Message { get; set; }
    }

    /// <summary>
    /// Class for communicating with WPF and present rules data
    /// </summary>
    static class DeclineRuleDao
    {
        /// <summary>
        /// Populates DeclineRules with entries from settings
        /// </summary>
        /// <returns>An ObservableCollection used for data binding in WPF</returns>
        public static ObservableCollection<DeclineRuleRecord> LoadData()
        {
            IDictionary<string, DeclineRuleSetting> rules = Properties.Settings.Default.MeetingDeclineRules;
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
                    DeclineRuleSetting rule = rules[folder.EntryID];
                    isActive = rule.IsActive;
                    sendResponse = rule.SendResponse;
                    isDecline = rule.Response == OlMeetingResponse.olMeetingDeclined;
                }

                // create new rule to be displayed
                DeclineRuleRecord item = new DeclineRuleRecord()
                {
                    FolderID = folder.EntryID,
                    FolderName = folder.Name,
                    FolderPath = folder.FolderPath.Replace(toRemove, ""),
                    IsActive = isActive,
                    IsDecline = isDecline,
                    SendResponse = sendResponse
                };

                // set property changes event handler for saving
                item.PropertyChanged += Rule_PropertyChanged;

                // add to collection
                results.Add(item);
            }

            return results;
        }

        /// <summary>
        /// Gets all folders in Outlook recursively
        /// </summary>
        /// <param name="folder">Root folder to start searching from</param>
        /// <returns>List of all folders found starting from root</returns>
        private static IEnumerable<MAPIFolder> GetFolders(MAPIFolder folder)
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

        /// <summary>
        /// Event handler used to save a decline rule change
        /// </summary>
        public static void Rule_PropertyChanged(object sender, PropertyChangedEventArgs e)
        {
            if (!(sender is DeclineRuleRecord record))
                return;

            IDictionary<string, DeclineRuleSetting> rules = Properties.Settings.Default.MeetingDeclineRules;
            if (rules.ContainsKey(record.FolderID))
            {
                // change existing rule setting
                DeclineRuleSetting rule = Properties.Settings.Default.MeetingDeclineRules[record.FolderID];
                rule.IsActive = record.IsActive;
                rule.SendResponse = record.SendResponse;
                rule.Response = record.IsDecline ? OlMeetingResponse.olMeetingDeclined : OlMeetingResponse.olMeetingTentative;
            }
            else
            {
                // set new rule settings
                Properties.Settings.Default.MeetingDeclineRules.Add(record.FolderID, new DeclineRuleSetting()
                {
                    IsActive = record.IsActive,
                    SendResponse = record.SendResponse,
                    Response = record.IsDecline ? OlMeetingResponse.olMeetingDeclined : OlMeetingResponse.olMeetingTentative
                });
            }

            // and save
            Properties.Settings.Default.Save();
        }
    }

    /// <summary>
    /// POCO for keeping and updating a deline rule settings
    /// </summary>
    class DeclineRuleRecord : INotifyPropertyChanged
    {
        private string _folderID;
        private string _folderName;
        private string _folderPath;
        private bool   _isActive;
        private bool   _isDecline;
        private bool   _sendResponse;

        public string FolderID     { get { return _folderID;     } set { _folderID     = value; OnPropertyChanged(); } }
        public string FolderName   { get { return _folderName;   } set { _folderName   = value; OnPropertyChanged(); } }
        public string FolderPath   { get { return _folderPath;   } set { _folderPath   = value; OnPropertyChanged(); } }
        public bool   IsActive     { get { return _isActive;     } set { _isActive     = value; OnPropertyChanged(); } }
        public bool   IsDecline    { get { return _isDecline;    } set { _isDecline    = value; OnPropertyChanged(); } }
        public bool   SendResponse { get { return _sendResponse; } set { _sendResponse = value; OnPropertyChanged(); } }

        public event PropertyChangedEventHandler PropertyChanged;
        private static readonly PropertyChangedEventArgs NULL_PROPERTY = new PropertyChangedEventArgs(String.Empty);
        protected void OnPropertyChanged() => PropertyChanged?.Invoke(this, NULL_PROPERTY);
    }
}
