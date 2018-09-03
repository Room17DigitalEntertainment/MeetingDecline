using System.Collections.ObjectModel;
using System.Windows;
using System.Windows.Controls;

namespace Room17DE.MeetingDecline.Forms
{
    /// <summary>
    /// Interaction logic for RulesForm2.xaml
    /// </summary>
    public partial class RulesForm : Window
    {
        public RulesForm()
        {
            InitializeComponent();

            // load and bind data
            mainGrid.ItemsSource = Util.DeclineRuleDao.LoadData();
        }

        /// <summary>
        /// Event handler to show a dialog for add a response message when declining
        /// </summary>
        private void MessageButton_Click(object sender, RoutedEventArgs e)
        {
            if (!(sender is Button button)) return;
            if (!(button.Tag is string folderID)) return;

            // send folderID and Message to the input message form
            new DeclineMessageForm(folderID).ShowDialog();
        }

        /// <summary>
        /// Event handler to close the dialog when pressong OK button
        /// </summary>
        private void OKButton_Click(object sender, RoutedEventArgs e)
        {
            this.DialogResult = true;
            this.Close();
        }

        /// <summary>
        /// Event handler to close the dialog when pressong Cancel button
        /// </summary>
        private void CancelButton_Click(object sender, RoutedEventArgs e)
        {
            this.DialogResult = false;
            this.Close();
        }
    }
}
