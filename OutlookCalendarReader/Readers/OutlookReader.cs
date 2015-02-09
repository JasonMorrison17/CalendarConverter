using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Office.Interop.Outlook;

namespace OutlookCalendarReader.Readers
{
    class OutlookReader : IReadableCalendar
    {
        public String UserName
        {
            get
            {
                return _user;
            }

            set
            {
                _user = value;
            }
        }

        public String Password
        {
            set
            {
                _password = value;
            }
        }

        private String _user { get; set; }

        private String _password { get; set; }

        public OutlookReader(String user, String password)
        {
            this._user = user;
            this._password = password;
        }

        public void readCalendar()
        {

        }

        public void getCalendarItemsInRange(DateTime startTime, DateTime endTime)
        {
            string filter = "[Start] >= '" + startTime.ToString("g") + "' AND [End] <= '" + endTime.ToString("g") + "'";
            Items calendarItems = getCalendarItems(filter);
            foreach (AppointmentItem apt in calendarItems)
            {
                if (apt.Categories != null)
                {
                    Console.WriteLine(apt.Categories + ": " + apt.Subject);
                }
                else
                {
                    Console.WriteLine(apt.Subject);
                }
            }
        }

        private Items getCalendarItems(string filter){
            Application oApp = null;
            NameSpace mapiNamespace = null;
            MAPIFolder calendarFolder = null;
            Items outlookCalendarItems = null;

            try
            {
                oApp = new Application();
                mapiNamespace = oApp.GetNamespace("MAPI"); ;
                mapiNamespace.Logon(_user, _password, false, false);
                calendarFolder = mapiNamespace.GetDefaultFolder(OlDefaultFolders.olFolderCalendar);
                outlookCalendarItems = calendarFolder.Items;
                outlookCalendarItems.IncludeRecurrences = true;
                //outlookCalendarItems.Sort("[Start]", Type.Missing);
                Items restrictItems = outlookCalendarItems.Restrict(filter);

                return restrictItems;
            }
            catch (System.Exception e)
            {
                throw e;
            }
            finally
            {
                oApp = null;
                mapiNamespace = null;
                calendarFolder = null;
                outlookCalendarItems = null;
            }
        }

        public void getAllCalendarItems()
        {
            Application oApp = null;
            NameSpace mapiNamespace = null;
            MAPIFolder calendarFolder = null;
            Items outlookCalendarItems = null;

            try
            {
                oApp = new Application();
                mapiNamespace = oApp.GetNamespace("MAPI"); ;
                mapiNamespace.Logon(_user, _password, false, false);
                calendarFolder = mapiNamespace.GetDefaultFolder(OlDefaultFolders.olFolderCalendar);
                outlookCalendarItems = calendarFolder.Items;
                outlookCalendarItems.IncludeRecurrences = true;

                foreach (AppointmentItem item in outlookCalendarItems)
                {
                    Console.WriteLine(item.Categories);
                    //if (item.IsRecurring)
                    //{
                    //    RecurrencePattern rp = item.GetRecurrencePattern();
                    //    DateTime first = new DateTime(2008, 8, 31, item.Start.Hour, item.Start.Minute, 0);
                    //    DateTime last = new DateTime(2008, 10, 1);
                    //    AppointmentItem recur = null;

                    //    for (DateTime cur = first; cur <= last; cur = cur.AddDays(1))
                    //    {
                    //        try
                    //        {
                    //            recur = rp.GetOccurrence(cur);                            
                    //            MessageBox.Show(recur.Subject + " -> " + cur.ToLongDateString());
                    //        }
                    //        catch
                    //        { }
                    //    }
                    //}
                    //else
                    //{
                    //    MessageBox.Show(item.Subject + " -> " + item.Start.ToLongDateString());
                    //}
                }
            }
            catch (System.Exception e)
            {
                throw e;
            }
            finally
            {
                oApp = null;
                mapiNamespace = null;
                calendarFolder = null;
                outlookCalendarItems = null;
            }

        }
    }
}
