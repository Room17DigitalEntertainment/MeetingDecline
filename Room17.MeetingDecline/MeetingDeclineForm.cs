using Microsoft.Office.Interop.Outlook;
using Room17.MeetingDecline.Util;
using System;
using System.Collections.Generic;
using System.Windows.Forms;

namespace Room17.MeetingDecline
{
    public partial class MeetingDeclineForm : Form
    {
        private IDictionary<string, MeetingDeclineSetting> Settings;

        public MeetingDeclineForm()
        {
            InitializeComponent();

            // read settings
            if (Properties.Settings.Default.MeetingDeclineSettings == null)
                Properties.Settings.Default.MeetingDeclineSettings = new Dictionary<string, MeetingDeclineSetting>();
            Settings = Properties.Settings.Default.MeetingDeclineSettings;
        }

        // TODO: ask user for default behavior on folder add

        /// <summary>
        /// Save configuration on item check/uncheck
        /// </summary>
        private void FoldersListBox_ItemCheck(object sender, ItemCheckEventArgs e)
        {
            CheckBoxEntry checkBox = foldersListBox.Items[e.Index] as CheckBoxEntry;
            if (!(checkBox.Value is MAPIFolder folder)) return;

            if (e.NewValue == CheckState.Checked)
                Settings[folder.EntryID] = new MeetingDeclineSetting() { FolderEntryID = folder.EntryID };
            else
                Settings.Remove(folder.EntryID);
            Properties.Settings.Default.Save();
        }

        /// <summary>
        /// Load folders to be shown by reading them from runtime + apply setting on it
        /// </summary>
        private void MeetingDeclinedForm_Load(object sender, EventArgs e)
        {
            // get all folders
            MAPIFolder root = Globals.AddIn.Application.Session.DefaultStore.GetRootFolder();
            IEnumerable<MAPIFolder> allFolders = GetFolders(root);

            // comfigure the list
            foldersListBox.DisplayMember = "Text";
            foldersListBox.ValueMember = "Value";

            // disable the item check event so it wont trigger when adding elements below
            this.foldersListBox.ItemCheck -= this.FoldersListBox_ItemCheck;

            // show all folders
            foreach (MAPIFolder folder in allFolders)
            {
                // get folder setting and show it
                if (Settings.ContainsKey(folder.EntryID))
                    foldersListBox.Items.Add(new CheckBoxEntry { Text = folder.Name, Value = folder }, true);
                // else show folder without check
                else
                    foldersListBox.Items.Add(new CheckBoxEntry { Text = folder.Name, Value = folder }, false);
            }

            // enable the item check event 
            this.foldersListBox.ItemCheck += this.FoldersListBox_ItemCheck;

            // TODO: load debug setting too
            // TODO: configure url label to open debug logs
        }

        /// <summary>
        /// Gets all folders in Outlook recursively
        /// </summary>
        /// <param name="folder">Root folder to start searching from</param>
        /// <returns>List of all folders found starting from root</returns>
        private IEnumerable<MAPIFolder> GetFolders(MAPIFolder folder)
        {
            // TODO: filter unwanted system folders
            if (folder.Folders.Count == 0)
                yield return folder;
            else
                foreach (MAPIFolder subFolder in folder.Folders)
                    foreach (MAPIFolder result in GetFolders(subFolder))
                        yield return result;
        }

        /// <summary>
        /// Handle OK button click event
        /// </summary>
        private void OK_Click(object sender, EventArgs e)
        {
            // save debug setting
            Properties.Settings.Default["Debug"] = debugCheckBox.Checked;
            // apply debug setting
            Utils.Logger.DEBUG = debugCheckBox.Checked;
        }
    }

    class CheckBoxEntry
    {
        public string Text { get; set; }
        public MAPIFolder Value { get; set; }
    }
}
