using System;
using System.Windows.Forms;

namespace Room17.MeetingDecline.Forms
{
    public partial class DeclineMessageForm : Form
    {
        private string FolderID;

        public DeclineMessageForm(string folderID)
        {
            FolderID = folderID;
            InitializeComponent();

            // show Message
            messageBox.Text = Properties.Settings.Default.MeetingDeclineRules[FolderID]?.Message;
        }

        /// <summary>
        /// Event handler to save in Settings the Message associated with current rule
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void OkButton_Click(object sender, EventArgs e)
        {
            // avoid NPE on fresh settings
            if (!Properties.Settings.Default.MeetingDeclineRules.ContainsKey(FolderID))
                Properties.Settings.Default.MeetingDeclineRules[FolderID] = new Util.DeclineRule();

            // save the message
            Properties.Settings.Default.MeetingDeclineRules[FolderID].Message = messageBox.Text;
            Properties.Settings.Default.Save();
        }
    }
}
