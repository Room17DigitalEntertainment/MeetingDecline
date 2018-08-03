using Microsoft.Office.Interop.Outlook;
using Room17.MeetingDecline;
using Room17.MeetingDecline.Util;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Windows.Forms;

namespace Room17.Forms.MeetingDecline
{
    public partial class RulesForm : Form
    {
        private IDictionary<string, DeclineRule> Rules;
        private Padding Padding1 = new Padding(3, 5, 3, 0);
        private Padding Padding2 = new Padding(3, 3, 3, 3);
        private Padding Padding3 = new Padding(3, 4, 3, 3);
        private Point Point1 = new Point(3, -1);
        private Point Point2 = new Point(96, -1);
        private Size Size1 = new Size(202, 24);

        public RulesForm()
        {
            InitializeComponent();
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

        // TODO: tooltip for folder label with full folder path
        // TODO: show loading bar

        /// <summary>
        /// Event handler for loading folders to be shown by reading them from runtime + apply setting on it
        /// </summary>
        private void MeetingDeclinedForm_Load(object sender, EventArgs e)
        {
            // read settings
            if (Room17.MeetingDecline.Properties.Settings.Default.MeetingDeclineRules == null)
                Room17.MeetingDecline.Properties.Settings.Default.MeetingDeclineRules = new Dictionary<string, DeclineRule>();
            Rules = Room17.MeetingDecline.Properties.Settings.Default.MeetingDeclineRules;

            // put an event handler to draw table lines
            rulesTablePanel.CellPaint += RulesTablePanel_CellPaint;

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

                // get folder setting and show it
                if (Rules.ContainsKey(folder.EntryID))
                {
                    DeclineRule rule = Rules[folder.EntryID];
                    isActive = rule.IsActive;
                    sendNotification = rule.SendNotification;
                    isDecline = rule.Response == OlMeetingResponse.olMeetingDeclined;
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
                LinkLabel linkLabel = new LinkLabel() { Text = "Message", AutoSize = true, Margin = Padding1, Tag = folder.EntryID };
                linkLabel.LinkClicked += MessageLabel_LinkClicked;
                rulesTablePanel.Controls.Add(linkLabel, 4, rulesTablePanel.RowCount - 1);
            }
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
            for (int i = 1; i < rulesTablePanel.RowCount; i++)
            {
                CheckBox enabledCheck = rulesTablePanel.GetControlFromPosition(1, i) as CheckBox;
                Panel panel = rulesTablePanel.GetControlFromPosition(2, i) as Panel;
                RadioButton declineButton = panel.Controls[0] as RadioButton;
                CheckBox sendCheck = rulesTablePanel.GetControlFromPosition(3, i) as CheckBox;
                MAPIFolder folder = enabledCheck.Tag as MAPIFolder;

                // do not override existing rule, otherwise Message field is lost
                if (Rules.ContainsKey(folder.EntryID))
                {
                    Rules[folder.EntryID].IsActive = enabledCheck.Checked;
                    Rules[folder.EntryID].Response =
                        declineButton.Checked ? OlMeetingResponse.olMeetingDeclined : OlMeetingResponse.olMeetingTentative;
                    Rules[folder.EntryID].SendNotification = sendCheck.Checked;
                }
                else
                    Rules[folder.EntryID] = new DeclineRule()
                    {
                        IsActive = enabledCheck.Checked,
                        Response = declineButton.Checked ? OlMeetingResponse.olMeetingDeclined : OlMeetingResponse.olMeetingTentative,
                        SendNotification = sendCheck.Checked
                    };
            }

            Room17.MeetingDecline.Properties.Settings.Default.Save();
        }

        private void MessageLabel_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            if (!(sender is LinkLabel messageLink)) return;
            if (!(messageLink.Tag is string folderID)) return;
            
            // send folderID and Message to the input message form
            new Room17.MeetingDecline.Forms.DeclineMessageForm(folderID).ShowDialog();
        }
    }
}
