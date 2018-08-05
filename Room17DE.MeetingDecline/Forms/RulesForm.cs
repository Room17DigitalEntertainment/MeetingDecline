using Microsoft.Office.Interop.Outlook;
using Room17DE.MeetingDecline;
using Room17DE.MeetingDecline.Util;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Room17DE.Forms.MeetingDecline
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

        /// <summary>
        /// Event handler for loading folders to be shown by reading them from runtime + apply setting on it
        /// </summary>
        private async void MeetingDeclinedForm_Load(object sender, EventArgs e)
        {
            // load icon
            this.Icon = Room17DE.MeetingDecline.Properties.Resources.icon;

            // put an event handler to draw table lines
            rulesTablePanel.CellPaint += RulesTablePanel_CellPaint;

            // add vetical scroll and remove horizontal scroll
            rulesTablePanel.HorizontalScroll.Maximum = 0;
            rulesTablePanel.AutoScroll = false;
            rulesTablePanel.VerticalScroll.Visible = false;
            rulesTablePanel.AutoScroll = true;

            // hide table and disable OK button until rules are loaded
            rulesTablePanel.Visible = okButton.Enabled = false;

            // show a progress bar
            ProgressBar progressBar = new ProgressBar()
                { Style = ProgressBarStyle.Marquee, Parent = this, Size = new Size(400, 23), Anchor = AnchorStyles.Right | AnchorStyles.Left };
            progressBar.Left = (this.ClientSize.Width - progressBar.Width) / 2;
            progressBar.Top = (this.ClientSize.Height - progressBar.Height) / 2;

            // load data inside table async
            await LoadRules();

            // make table visible and OK button enabled
            rulesTablePanel.Visible = okButton.Enabled = true;

            progressBar.Dispose();
        }

        /// <summary>
        /// Asynchronously load and display Rules inside table layout panel
        /// </summary>
        /// <returns>A task to await or continue with</returns>
        private async Task<bool> LoadRules()
        {
            return await Task.Run(() =>
            {
                try
                {
                    // read settings
                    if (Room17DE.MeetingDecline.Properties.Settings.Default.MeetingDeclineRules == null)
                        Room17DE.MeetingDecline.Properties.Settings.Default.MeetingDeclineRules = new Dictionary<string, DeclineRule>();
                    Rules = Room17DE.MeetingDecline.Properties.Settings.Default.MeetingDeclineRules;

                    // get all folders
                    MAPIFolder root = Globals.AddIn.Application.Session.DefaultStore.GetRootFolder();
                    IEnumerable<MAPIFolder> allFolders = GetFolders(root);

                    string toRemove = "\\\\" + Globals.AddIn.Application.Session.CurrentUser.AddressEntry.GetExchangeUser().PrimarySmtpAddress + "\\";

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

                        // add a new row using UI thread
                        rulesTablePanel.Invoke(new System.Action(() =>
                        {
                            // add table row
                            rulesTablePanel.RowCount++;
                            rulesTablePanel.RowStyles.Add(new RowStyle(SizeType.Absolute, 40F));

                            // add label with folder name
                            Label folderLabel = new Label() { Text = folder.Name, Margin = Padding1, AutoSize = true };
                            ToolTip toolTip = new ToolTip() { ToolTipIcon = ToolTipIcon.None };
                            toolTip.SetToolTip(folderLabel, folder.FolderPath.Replace(toRemove, ""));
                            rulesTablePanel.Controls.Add(folderLabel, 0, rulesTablePanel.RowCount - 1);

                            // add checkbox for enable rule
                            rulesTablePanel.Controls.Add(
                                    new CheckBox() { Text = "Enabled", Checked = isActive, Margin = Padding2, Tag = folder }, 1, rulesTablePanel.RowCount - 1);

                            // add radio buttons for decline/tentative
                            Panel panel = new Panel() { Margin = Padding3, Size = Size1 };
                            RadioButton declineButton =
                                new RadioButton() { Text = "Decline", Margin = Padding2, Checked = isDecline, AutoSize = true, Location = Point1 };
                            panel.Controls.Add(declineButton);
                            panel.Controls.Add(
                                new RadioButton() { Text = "Tentative", Margin = Padding2, Checked = !isDecline, AutoSize = true, Location = Point2 });
                            rulesTablePanel.Controls.Add(panel, 2, rulesTablePanel.RowCount - 1);

                            // add checkbox for send response
                            rulesTablePanel.Controls.Add(
                                    new CheckBox() { Text = "Send response", Checked = sendNotification, AutoSize = true }, 3, rulesTablePanel.RowCount - 1);

                            // add link for setting a message
                            LinkLabel linkLabel = new LinkLabel() { Text = "Message", AutoSize = true, Margin = Padding1, Tag = folder.EntryID };
                            linkLabel.LinkClicked += MessageLabel_LinkClicked;
                            rulesTablePanel.Controls.Add(linkLabel, 4, rulesTablePanel.RowCount - 1);
                        }));
                    }

                    return true;
                }
                catch(InvalidOperationException) // when clicking Cancel and table is not loaded
                {
                    return false;
                }
            });
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

            Room17DE.MeetingDecline.Properties.Settings.Default.Save();
        }

        private void MessageLabel_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            if (!(sender is LinkLabel messageLink)) return;
            if (!(messageLink.Tag is string folderID)) return;
            
            // send folderID and Message to the input message form
            new Room17DE.MeetingDecline.Forms.DeclineMessageForm(folderID).ShowDialog();
        }
    }
}
