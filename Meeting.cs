using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace RendToOutlook
{
    public class Meeting
    {
        private string subject;
        private string location;
        private string email;
        private DateTime startTime;
        private DateTime endTime;

        public Meeting( string sub, string loc, string start, string end, string mail )
        {
            subject = sub;
            location = loc;
            //converts the start and end time string to a DateTime object
            DateTime.TryParseExact( start, new string[] { "d/M/yyyy H:mm" }, System.Globalization.CultureInfo.InvariantCulture, System.Globalization.DateTimeStyles.None, out startTime );
            DateTime.TryParseExact( end, new string[] { "d/M/yyyy H:mm" }, System.Globalization.CultureInfo.InvariantCulture, System.Globalization.DateTimeStyles.None, out endTime );

            email = mail;
        }

        //Writes the called meeting's subject,location,start time,end time,and email address to the console 
        public void Display()
        {
            Console.WriteLine( subject );
            Console.WriteLine( location );
            Console.WriteLine( startTime );
            Console.WriteLine( endTime );
            Console.WriteLine( email );
        }

        public string Subject
        {
            get { return subject; }
            set { subject = value; }
        }

        public string Location
        {
            get { return location; }
            set { location = value; }
        }

        public string Email
        {
            get { return email; }
            set { email = value; }
        }

        public DateTime StartTime
        {
            get { return startTime; }
            set { startTime = value; }
        }

        public DateTime EndTime
        {
            get { return endTime; }
            set { endTime = value; }
        }
    }
}