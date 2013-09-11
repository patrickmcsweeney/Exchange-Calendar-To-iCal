using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Exchange.WebServices.Data;

namespace Exchange_Calendar_To_iCal
{
    class Program
    {
        static void Main(string[] args)
        {
            if (args.Length != 4)
            {
                Console.WriteLine("This application must be passed 4 arguments:");
                Console.WriteLine("  <username> - The username of the user requesting a calendar. (E.g. abc1d23).");
                Console.WriteLine("  <password> - The pass word of the user requesting a calendar.");
                Console.WriteLine("  <calendar> - The name of the calendar. (E.g. building 32 room 4049 would be b32r4049).");
                Console.WriteLine("  <writefile> - The file to write the iCal for the calendar to. (E.g. C:\\Users\\User\\b32r4049.ics or Windows or /home/user/b32r4049.ics on Linux).\n");
                Environment.Exit(1);
            }

            // Configure filtering options
            FilterSetting FilterOptions = new FilterSetting();
            FilterOptions.SearchStartDate = DateTime.Now.AddDays(-30);
            FilterOptions.SearchEndDate = DateTime.Now.AddDays(90);

            // Establish web service session
            ExchangeService WebService = ExchangeWebServiceWrapper.GetWebServiceInstance(args[0], args[1], "SOTON");

            // Retrieve calendar folder
            Console.WriteLine("Getting calendar...");
            Mailbox mailbox = new Mailbox(args[2]+"@soton.ac.uk");
            FolderId id = new FolderId(WellKnownFolderName.Calendar, mailbox);
            CalendarFolder Calendar = CalendarFolder.Bind(WebService, id);
            if (Calendar == null)
            {
                throw new Exception("Could not find calendar folder.");
            }

            // Retrieve appointments
            Console.WriteLine("Getting appointments..."); 
            FindItemsResults<Appointment> Appointments = Calendar.FindAppointments(new CalendarView(FilterOptions.SearchStartDate, FilterOptions.SearchEndDate));
            iCalWrapper iCal = new iCalWrapper();
            iCal.AddEventsToCalendar(Appointments, FilterOptions);

            // Write iCal file to disk
            Console.WriteLine("Writing to file "+args[3]+"..."); 
            iCal.GenerateFile(args[3]);
            Console.WriteLine("Done");
        }
    }
}
