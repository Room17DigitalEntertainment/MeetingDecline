using Microsoft.Office.Interop.Outlook;
using Room17.MeetingDecline.Util;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Windows.Forms;

namespace Room17.MeetingDecline
{
    public partial class MeetingDeclineForm : Form
    {
        private IDictionary<string, MeetingDeclineRule> Rules;
        private Padding Padding1 = new Padding(3, 5, 3, 0);
        private Padding Padding2 = new Padding(3, 3, 3, 3);
        private Padding Padding3 = new Padding(3, 4, 3, 3);
        private Point Point1 = new Point(3, -1);
        private Point Point2 = new Point(96, -1);
        private Size Size1 = new Size(202, 24);

        public MeetingDeclineForm()
        {
            InitializeComponent();

            // read settings
            if (Properties.Settings.Default.MeetingDeclineRules == null)
                Properties.Settings.Default.MeetingDeclineRules = new Dictionary<string, MeetingDeclineRule>();
            Rules = Properties.Settings.Default.MeetingDeclineRules;

            // put an event handler to draw table lines
            rulesTablePanel.CellPaint += RulesTablePanel_CellPaint;
        }

        /// <summary>
        /// Event handler to draw lines between rows
        /// </summary>
        private void RulesTablePanel_CellPaint(object sender, TableLayoutCellPaintEventArgs e)
        {
            if (e.Row > 0)
                e.Graphics.DrawLine(Pens.Black, new Point(e.CellBounds.Left, e.CellBounds.Top - 5),
                    new Point(e.CellBounds.Right, e.CellBounds.Top - 5));
        }

        // TODO: ask user for default behavior on folder add (autodecline it or ignore)
        // TODO: choose between decline or tentative (default decline)
        // TODO: send notification back or not (default not)
        // TODO: send a message (default no message)

        /// <summary>
        /// Event handler for loading folders to be shown by reading them from runtime + apply setting on it
        /// </summary>
        private void MeetingDeclinedForm_Load(object sender, EventArgs e)
        {
            // get all folders
            MAPIFolder root = Globals.AddIn.Application.Session.DefaultStore.GetRootFolder();
            IEnumerable<MAPIFolder> allFolders = GetFolders(root);

            // add vetical scroll and remove horizontal scroll
            rulesTablePanel.HorizontalScroll.Maximum = 0;
            rulesTablePanel.AutoScroll = false;
            rulesTablePanel.VerticalScroll.Visible = false;
            rulesTablePanel.AutoScroll = true;

            // show all folders
            foreach (MAPIFolder folder in allFolders)
            {
                bool isActive = false;
                bool sendNotification = false;
                bool isDecline = true;
                string message = null;

                // get folder setting and show it
                if (Rules.ContainsKey(folder.EntryID))
                {
                    MeetingDeclineRule rule = Rules[folder.EntryID];
                    isActive = rule.IsActive;
                    sendNotification = rule.SendNotification;
                    isDecline = rule.Response == OlMeetingResponse.olMeetingDeclined;
                    message = rule.Message;
                }

                // add table row
                rulesTablePanel.RowCount++;
                rulesTablePanel.RowStyles.Add(new RowStyle(SizeType.Absolute, 40F));
                rulesTablePanel.Controls.Add(
                    new Label() { Text = folder.Name, Margin = Padding1, AutoSize = true }, 0, rulesTablePanel.RowCount - 1);
                rulesTablePanel.Controls.Add(
                    new CheckBox() { Text = "Enabled", Checked = isActive, Margin = Padding2, Tag = folder }, 1, rulesTablePanel.RowCount - 1);
                Panel panel = new Panel() { Margin = Padding3, Size = Size1 };
                RadioButton declineButton =
                    new RadioButton() { Text = "Decline", Margin = Padding2, Checked = isDecline, AutoSize = true, Location = Point1 };
                panel.Controls.Add(declineButton);
                panel.Controls.Add(
                    new RadioButton() { Text = "Tentative", Margin = Padding2, Checked = !isDecline, AutoSize = true, Location = Point2 });
                rulesTablePanel.Controls.Add(panel, 2, rulesTablePanel.RowCount - 1);
                rulesTablePanel.Controls.Add(
                    new CheckBox() { Text = "Send response", Checked = sendNotification, AutoSize = true }, 3, rulesTablePanel.RowCount - 1);
                // TODO: add Message too
            }

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
        /// Handle OK button click event and save rules
        /// </summary>
        private void OK_Click(object sender, EventArgs e)
        {
            for (int i = 0; i < rulesTablePanel.RowCount; i++)
            {
                for (int j = 1; j < rulesTablePanel.ColumnCount; j++) // ignore first column, aka the Label
                {
                    CheckBox enabledCheck = rulesTablePanel.GetControlFromPosition(i, j) as CheckBox;
                    Panel panel = rulesTablePanel.GetControlFromPosition(i, j+1) as Panel;
                    RadioButton declineButton = panel.Controls[0] as RadioButton;
                    CheckBox sendCheck = rulesTablePanel.GetControlFromPosition(i, j+2) as CheckBox;
                    MAPIFolder folder = enabledCheck.Tag as MAPIFolder;

                    Rules[folder.EntryID] = new MeetingDeclineRule() {
                        IsActive = enabledCheck.Checked,
                        Message = null,
                        Response = declineButton.Checked ? OlMeetingResponse.olMeetingDeclined : OlMeetingResponse.olMeetingTentative,
                        SendNotification = sendCheck.Checked
                    };
                }
            }

            Properties.Settings.Default.Save();
        }

        /// <summary>
        /// Handle debug checbok check event and save its state
        /// </summary>
        private void debugCheckBox_CheckedChanged(object sender, EventArgs e)
        {
            // save debug setting
            Properties.Settings.Default["Debug"] = debugCheckBox.Checked;
            // apply debug setting
            Util.Logger.DEBUG = debugCheckBox.Checked;
        }
    }

    class CheckBoxEntry
    {
        public string Text { get; set; }
        public MAPIFolder Value { get; set; }
    }
}
