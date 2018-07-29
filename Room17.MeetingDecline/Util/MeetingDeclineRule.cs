using Microsoft.Office.Interop.Outlook;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Room17.MeetingDecline.Util
{
    [Serializable]
    class MeetingDeclineRule
    {
        // default values
        private bool _isActive = false;
        private bool _sendNotification = false;
        private OlMeetingResponse _response = OlMeetingResponse.olMeetingDeclined;

        public bool IsActive { get => _isActive; set => _isActive = value; }
        public bool SendNotification { get => _sendNotification; set => _sendNotification = value; }
        public OlMeetingResponse Response { get => _response; set => _response = value; }
        public string Message { get; set; }
    }
}
