using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using OutlookCalendarReader.Readers;
using OutlookCalendarReader.TimeSheets;

namespace OutlookCalendarReader
{
    class Program
    {
        static void Main(string[] args)
        {
            //OutlookReader reader = new OutlookReader("jasonmorrsion@fairfaxdatasystems.com", "Cooljkk10527");
            //DateTime start = new DateTime(2015, 1, 1);
            //DateTime end = new DateTime(2015, 1, 31);            
            //reader.getCalendarItemsInRange(start, end);

            ATTimeSheet ts = new ATTimeSheet("jasonmorrison@fairfaxdatasystems.com", "Cooljkk10527");


            Console.Write("Press any key to exit...");
            Console.ReadLine();
        }
    }
}
