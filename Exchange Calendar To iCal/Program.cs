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
            string PublicFolderPath = @"Professional Services\iSolutions\ITIL Processes\Change\Forward Schedule of Change";

            FilterSetting FilterOptions = new FilterSetting();
            // TODO : Read in the following as arugments
            //FilterOptions.SubjectFilter = "(iss-)|(Migrations)";
            FilterOptions.SubjectFilter = "(Building|Bld|Bldg) [0-9][0-9]";
            //string[] CategoryFilter = { "DBA", "EST", "DST" };            
            //FilterOptions.CategoryFilter = CategoryFilter;
            FilterOptions.SearchStartDate = DateTime.Now;
            FilterOptions.SearchEndDate = DateTime.Now.AddDays(10);

            // Initialise web service instance.
            // NOTE: This is an overloaded method - alternate versions take credentials/username+password - switch out if needed
            ExchangeService WebService = ExchangeWebServiceWrapper.GetWebServiceInstance();

            // Retrieve calendar folder
            Console.WriteLine("Getting calendar...");
            CalendarFolder Calendar = ExchangeWebServiceWrapper.GetCalendarFolder(WebService, PublicFolderPath);

            if (Calendar == null)
            {
                throw new Exception("Could not find calendar folder.");
            }

            // Retrieve appointments
            Console.WriteLine("Getting appointments..."); 
            FindItemsResults<Appointment> Appointments = Calendar.FindAppointments(new CalendarView(FilterOptions.SearchStartDate, FilterOptions.SearchEndDate));

            iCalWrapper iCal = new iCalWrapper();

            
            iCal.AddEventsToCalendar(Appointments, FilterOptions);

            Console.WriteLine(iCal.GenerateString());

            Console.WriteLine("\n\nDone");
            Console.ReadLine();
        }
    }
}
