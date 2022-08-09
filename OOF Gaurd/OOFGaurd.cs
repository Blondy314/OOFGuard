using System;
using System.Linq;
using Outlook = Microsoft.Office.Interop.Outlook;
using System.Windows.Forms;

namespace OOF_Gaurd
{
    public partial class OOFGaurd
    {
        private readonly string[] OofVerbs = new[] { "oof", "ooo", "vacation", "out of office" };

        private void ThisAddIn_Startup(object sender, EventArgs e)
        {
            Application.ItemSend += new Outlook.ApplicationEvents_11_ItemSendEventHandler(Application_ItemSend);
        }

        private void Application_ItemSend(object Item, ref bool Cancel)
        {
            var meeting = Item as Outlook.MeetingItem;

            if (meeting == null)
            {
                return;
            }

            var appointment = meeting.GetAssociatedAppointment(false);
            if (appointment == null)
            {
                return;
            }

            if (appointment.BusyStatus == Outlook.OlBusyStatus.olFree)
            {
                return;
            }

            // warn only when sending to multiple recipients (for instance, not to self)
            if (meeting.Recipients.Count == 1)
            {
                var rec = meeting.Recipients.Cast<Outlook.Recipient>().FirstOrDefault();

                // check this is not a DL
                if (rec != null && rec.DisplayType == Outlook.OlDisplayType.olUser)
                {
                    return;
                }
            }

            // warn about busy only if subject contains an OOF verb otherwise assume its a normal meeting
            if (appointment.BusyStatus == Outlook.OlBusyStatus.olBusy)
            {
                var subject = appointment.Subject.ToLower();

                if (!OofVerbs.Any(v => subject.ToLower().Contains(v)))
                {
                    return;
                }
            }

            var status = appointment.BusyStatus.ToString().Substring("ol".Length);

            var res = MessageBox.Show($"You are about to send a meeting with {status} status to {meeting.Recipients.Count} recepients.\n\n" +
                $"This will set all recipients to appear as {status} as well.\n\n" +
                "Are you sure?", "OOF Gaurd",
                MessageBoxButtons.YesNo,
                MessageBoxIcon.Question);

            if (res != DialogResult.Yes)
            {
                Cancel = true;
            }
        }

        private void ThisAddIn_Shutdown(object sender, EventArgs e)
        {

        }

        #region VSTO generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InternalStartup()
        {
            this.Startup += new EventHandler(ThisAddIn_Startup);
            this.Shutdown += new EventHandler(ThisAddIn_Shutdown);
        }

        #endregion
    }
}
