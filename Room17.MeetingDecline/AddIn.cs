using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using Outlook = Microsoft.Office.Interop.Outlook;
using Office = Microsoft.Office.Core;
using Microsoft.Office.Interop.Outlook;
using Room17.MeetingDecline.Utils;
using Microsoft.Office.Tools.Ribbon;

namespace Room17.MeetingDecline
{
    public partial class AddIn
    {
        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            // event handler for new email
            this.Application.NewMailEx += Application_NewMailEx;
        }

        private void Application_NewMailEx(string EntryIDCollection)
        {
            MeetingItem meetingItem = CheckEmail(EntryIDCollection);
            if (meetingItem == null)
                return;


            // TODO: get excluded folders 
            // get current meeting folder
            // compare current meeting folder with excluded ones
            // if true, send decline

            // TODO: on folder delete event should remove entry in settings

            //meetingItem.GetAssociatedAppointment(false).Respond(OlMeetingResponse.olMeetingDeclined, true);
        }

        internal MeetingItem CheckEmail(string EntryIDCollection)
        {
            object item = null;
            try
            {
                item = Globals.AddIn.Application.Session.GetItemFromID(EntryIDCollection);
            }
            catch (System.Exception ex)
            {
                Logger.Error(String.Format("Could not retrieve info for message id:{0}. Error message is:{1}{2}",
                    EntryIDCollection, Environment.NewLine, ex.ToString()));
                return null;
            }

            MeetingItem meetingItem = item as MeetingItem;
            if (item == null)
            {
                Logger.Debug(String.Format("Message with id {0} is not a MeetingItem.", EntryIDCollection));
                return null;
            }

            return meetingItem;
        }

        protected override Microsoft.Office.Core.IRibbonExtensibility CreateRibbonExtensibilityObject() => new Ribbon();

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
            // Note: Outlook no longer raises this event. If you have code that 
            //    must run when Outlook shuts down, see https://go.microsoft.com/fwlink/?LinkId=506785
        }

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
