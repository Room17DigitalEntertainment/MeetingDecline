using Microsoft.Office.Interop.Outlook;
using Room17DE.MeetingDecline;
using Room17DE.MeetingDecline.Util;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Room17DE.Forms.MeetingDecline
{
    public partial class RulesForm : Form
    {
        private IDictionary<string, DeclineRule> Rules;
        private readonly Padding Padding1 = new Padding(3, 4, 3, 3);
        private readonly Point Point1 = new Point(3, 3);
        private Size Size1;
        private float RowHeight;

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

            // compute table row height by dpi; equation must be changed when doing massive UI update
            // 144 is dpi for 150%, 96 is dpi for 100%
            Graphics graphics = this.CreateGraphics();
            this.RowHeight = (45F-30) / (144 - 96) * graphics.DpiY;

            // set form width according to dpi; equation must be changed when doing massive UI update
            this.Width = (int)((924F-616)/(144-96) * graphics.DpiX);
            this.MinimumSize = new Size(this.Width, this.MinimumSize.Height);

            // panel height
            Size1 = new Size((int)((208.5F-139)/(144-96) * graphics.DpiX), 17);

            // put an event handler to draw table lines
            rulesTablePanel.CellPaint += RulesTablePanel_CellPaint;

            // add vetical scroll and remove horizontal scroll
            rulesTablePanel.HorizontalScroll.Maximum = 0;
            rulesTablePanel.AutoScroll = false;
            rulesTablePanel.VerticalScroll.Visible = false;
            rulesTablePanel.AutoScroll = true;

            // reset rows
            rulesTablePanel.RowCount = 0;
            rulesTablePanel.RowStyles.Clear();

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
                            rulesTablePanel.RowStyles.Add(new RowStyle(SizeType.Absolute, this.RowHeight));

                            // add label with folder name
                            Label folderLabel = new Label() { Text = folder.Name, Padding = Padding1, AutoSize = true };
                            ToolTip toolTip = new ToolTip() { ToolTipIcon = ToolTipIcon.None };
                            toolTip.SetToolTip(folderLabel, folder.FolderPath.Replace(toRemove, ""));
                            rulesTablePanel.Controls.Add(folderLabel, 0, rulesTablePanel.RowCount - 1);

                            // add checkbox for enable rule
                            rulesTablePanel.Controls.Add(
                                new CheckBox() { Text = "Enabled", Checked = isActive, Location = Point1, AutoSize = true, Tag = folder },
                                1, rulesTablePanel.RowCount - 1);

                            // add radio buttons for decline/tentative
                            Panel panel = new Panel() { Location = Point1, Size = Size1 };
                            RadioButton declineButton =
                                new RadioButton() { Text = "Decline", Checked = isDecline, AutoSize = true, Dock = DockStyle.Left };
                            panel.Controls.Add(declineButton);
                            panel.Controls.Add(
                                new RadioButton() { Text = "Tentative", Checked = !isDecline, AutoSize = true, Dock = DockStyle.Right });
                            rulesTablePanel.Controls.Add(panel, 2, rulesTablePanel.RowCount - 1);

                            // add checkbox for send response
                            rulesTablePanel.Controls.Add(
                                new CheckBox() { Text = "Send response", Checked = sendNotification, AutoSize = true }, 3, rulesTablePanel.RowCount - 1);

                            // add link for setting a message
                            LinkLabel linkLabel = new LinkLabel() { Text = "Message", AutoSize = true, Padding = Padding1, Tag = folder.EntryID };
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
