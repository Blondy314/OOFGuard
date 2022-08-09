using Microsoft.Office.Tools.Ribbon;
using Microsoft.Office.Interop.Outlook;
using System.Linq;

namespace OOF_Gaurd
{
    public partial class Ribbon1
    {
        private void Ribbon1_Load(object sender, RibbonUIEventArgs e)
        {
        }

        private void button1_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                Application application = Globals.OOFGaurd.Application;

                AppointmentItem app = (AppointmentItem)application.CreateItem(OlItemType.olAppointmentItem);

                var ns = application.GetNamespace("MAPI");
                var user = ns.CurrentUser;
                var name = user.Name;

                var account = ns.Accounts.Cast<Account>().FirstOrDefault();
                var addr = name;

                if (account != null)
                {
                    addr = account.SmtpAddress;
                }

                app.Subject = "OOF";

                if (!string.IsNullOrEmpty(name))
                {
                    app.Subject = $"{name.Split(' ')[0]} {app.Subject}";
                }

                app.AllDayEvent = true;
                app.Location = "OOF";

                app.Recipients.Add(addr);
                app.Recipients.ResolveAll();

                app.MeetingStatus = OlMeetingStatus.olMeeting;
                app.ResponseRequested = false;

                app.BusyStatus = OlBusyStatus.olFree;

                app.Display(true);
            }
            catch
            {
            }
        }
    }
}
