using Microsoft.Office.Interop.Outlook;
using System;
using System.Collections.Generic;
using System.Threading;
using System.Threading.Tasks;

namespace Room17DE.MeetingDecline.Util
{
    static class NewMailPeriodicTask
    {
        public static async Task Run(TimeSpan period, CancellationToken cancellationToken)
        {
            // keep scanning for new mails while thread is not stopped
            while (!cancellationToken.IsCancellationRequested)
            {
                await Task.Delay(period, cancellationToken);

                if (!cancellationToken.IsCancellationRequested)
                    CheckNewEmails(cancellationToken);
            }
        }

        /// <summary>
        /// Starting point of check and apply decline rules for new mails
        /// </summary>
        /// <param name="cancellationToken">A cancelation token to stop this method from executing in an async context</param>
        private static void CheckNewEmails(CancellationToken cancellationToken)
        {
            bool save = false;

            // get folders from rules
            IDictionary<string, DeclineRule> rules = Properties.Settings.Default.MeetingDeclineRules;
            IDictionary<string, DateTime> lastCheckedFolders = Properties.Settings.Default.LastMailCheck;

            foreach (string folderID in rules.Keys)
            {
                // skip inactive rules
                if (!rules[folderID].IsActive)
                    continue;

                // get folder
                MAPIFolder folder = Globals.AddIn.Application.Session.GetFolderFromID(folderID);
                if (folder == null)
                    continue; // user deleted folder

                // get last date when new mails were checked
                DateTime lastCheckedDate = DateTime.MinValue;
                if (lastCheckedFolders.ContainsKey(folderID))
                    lastCheckedDate = lastCheckedFolders[folderID];

                // make sure to not check last mail twice
                if (lastCheckedDate != DateTime.MinValue)
                    lastCheckedDate = lastCheckedDate.AddMinutes(+1);

                // stop expensive ephemeral execution
                if (cancellationToken.IsCancellationRequested)
                    return;

                // get new emails
                folder.Items.Sort("[ReceivedTime]");
                Items results = folder.Items.Restrict(String.Format("[ReceivedTime] >= '{0}' AND [Unread]=true",
                    lastCheckedDate.ToString("g")));
                DateTime? lastDate = null;

                foreach (object item in results)
                {
                    // get only meetings
                    if (!(item is MeetingItem meetingItem))
                        continue;

                    // save date
                    lastDate = meetingItem.ReceivedTime;

                    // stop expensive ephemeral execution
                    if (cancellationToken.IsCancellationRequested)
                        return;

                    // process meeting rule
                    try { ProcessRule(meetingItem, rules[folderID]); }
                    catch { continue; }
                }

                // update latest mail entry for this folder
                if (lastDate != null)
                {
                    lastCheckedFolders[folderID] = (DateTime)lastDate;
                    save = true;
                }
            }

            // save new last date
            if (save)
                Properties.Settings.Default.Save();
        }

        /// <summary>
        /// Applies the appropiate rule to the meeting parameter
        /// </summary>
        /// <param name="meetingItem">The meeting that needs processing</param>
        /// <param name="rule">the associated decline rule for the folder of this meetingItem</param>
        private static void ProcessRule(MeetingItem meetingItem, DeclineRule rule)
        {
            // if it's a Cancelation, delete it from calendar
            if (meetingItem.Class == OlObjectClass.olMeetingCancellation)
            {
                if (meetingItem.GetAssociatedAppointment(false) != null)
                { meetingItem.GetAssociatedAppointment(false).Delete(); return; }
                meetingItem.Delete(); return; // if deleted by user/app, delete the whole message
            }

            // get associated appointment
            AppointmentItem appointment = meetingItem.GetAssociatedAppointment(false);
            string globalAppointmentID = appointment.GlobalAppointmentID;

            // optional, send notification back to sender
            appointment.ResponseRequested &= rule.SendNotification;

            // set decline to the meeting
            MeetingItem responseMeeting = appointment.Respond(rule.Response, true);
            // https://msdn.microsoft.com/en-us/VBA/Outlook-VBA/articles/appointmentitem-respond-method-outlook 
            // says that Respond() will return a new meeting object for Tentative response

            // optional, add a meesage to the Body
            if (!String.IsNullOrEmpty(rule.Message))
                (responseMeeting ?? meetingItem).Body = rule.Message;

            // send decline
            //if(rule.Response == OlMeetingResponse.olMeetingDeclined)
            (responseMeeting ?? meetingItem).Send();

            // and delete the appointment if tentative
            if (rule.Response == OlMeetingResponse.olMeetingTentative)
                appointment.Delete();

            // after Sending the response, sometimes the appointment doesn't get deleted from calendar,
            // but appointmnent could become and invalid object, so we need to search for it and delete it
            AppointmentItem newAppointment = Globals.AddIn.Application.Session.GetDefaultFolder(OlDefaultFolders.olFolderCalendar).Items
                .Find("@SQL=\"http://schemas.microsoft.com/mapi/id/{6ED8DA90-450B-101B-98DA-00AA003F1305}/00030102\" = '"
                + globalAppointmentID + "' ");
            if (newAppointment != null)
                newAppointment.Delete();

        }
    }
}
