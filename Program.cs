using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using Microsoft.Exchange.WebServices.Data;
using System.Text;
using System.Threading.Tasks;

namespace RendToOutlook
{
    class Program
    {
        static void Main( string[] args )
        {
            //Reads the roomData file and turns it into a multidimensional list
            string roomDataFilePath = @"C:\Users\rhoffer2\Desktop\CrestronRoomDisplay\RendToOutlook\RendToOutlook\data\roomData.csv";
            var roomData = File.ReadLines( roomDataFilePath ).Select( x => x.Split( ',' ) ).ToList();

            //Creates an empty list that will be populated with the rooms that have displays
            var roomsWithDisplays = new List<string>();

            //Reads the weeklySchedule file and turns it into a multidimensional list
            string weeklyScheduleDataFilePath = @"C:\Users\rhoffer2\Desktop\CrestronRoomDisplay\RendToOutlook\RendToOutlook\data\weeklySchedule.csv";
            var weeklyScheduleData = File.ReadLines( weeklyScheduleDataFilePath ).Select( x => x.Split( ',' ) ).ToList();

            //Create an array called rooms that contains the object Room() and is the size of the number of rooms that were read in
            Room[] rooms = new Room[roomData.Count];

            //Iterates through the read in data and creates an object for each individual room and adds the Rendezvous name to the roomsWithDisplays list
            for ( int index = 0; index < rooms.Length; index++ )
            {
                rooms[index] = new Room();

                rooms[index].Number = roomData[index][0];
                rooms[index].Name = roomData[index][1];
                rooms[index].Email = roomData[index][2];
                rooms[index].RendezvousName = roomData[index][3];
                roomsWithDisplays.Add( rooms[index].RendezvousName );

            }

            //Uncomment these lines to display the names of all the loaded in rooms
            //for ( int index = 0; index < rooms.Length; index++ )
            //{
            //    Console.WriteLine(rooms[index].Name);
            //}

            //Initiates the HashSet rowsToRemove which will be populated with the number of the location of meetings that dont have a display
            HashSet<int> rowsToRemove = new HashSet<int>();
            bool b;

            //Checks through the weeklyScheduleData and finds the location of meetings that dont have a display by checking rooms against the list of rooms with displays
            for ( int index = 0; index < weeklyScheduleData.Count; index++ )
            {
                if ( b = roomsWithDisplays.Any( weeklyScheduleData[index][0].Contains ) )
                {

                }
                else
                {
                    rowsToRemove.Add( index );
                }
            }

            //Creates a new list of meetings that are taking place in a room with a display
            var meetingsToSchedule = weeklyScheduleData.Where( ( arr, index ) => !rowsToRemove.Contains( index ) ).ToList();

            //Create an array called meetings that contains the object Meeting() and is the size of the number of meetings to be scheduled
            Meeting[] meetings = new Meeting[meetingsToSchedule.Count];

            //Initializing some variables that will be filled temporarily as data is iterated through
            string meetingsToScheduleRoomName, meetingsToScheduleStartDate, meetingsToScheduleStartTime, meetingsToScheduleEndDate, 
                meetingsToScheduleEndTime, meetingsToScheduleEventTitle;
            string tempSubject = null, tempLocation = null, tempStartTime = null, tempEndTime = null, tempEmail = null;


            //Iterates through the meetingsToSchedule list to grab the subject, start time and end time of the meetings to schedule. It then grabs the room number 
            //and email address from the appropriate room that meeting is scheduled in. Then it creates a new object of the class Meeting with that info.
            for ( int index = 0; index < meetings.Length; index++ )
            {
                //grabs the info from the meetingsToSchedule list and puts it in individual variables
                meetingsToScheduleRoomName = meetingsToSchedule[index][0];
                meetingsToScheduleStartDate = meetingsToSchedule[index][1];
                meetingsToScheduleStartTime = meetingsToSchedule[index][2];
                meetingsToScheduleEndDate = meetingsToSchedule[index][3];
                meetingsToScheduleEndTime = meetingsToSchedule[index][4];
                meetingsToScheduleEventTitle = meetingsToSchedule[index][5];

                //checks through the list of rooms to grab the appropriate room number and email address for each meeting
                for ( int roomSearchIndex = 0; roomSearchIndex < rooms.Length; roomSearchIndex++ )
                {
                    //This if statement checks the current room name with the names of the rooms in the rooms with displays list then finds the room number and email 
                    //when it gets to the proper room
                    if ( meetingsToScheduleRoomName == rooms[roomSearchIndex].RendezvousName )
                    {
                        tempLocation = "IISC " + rooms[roomSearchIndex].Number;
                        tempEmail = rooms[roomSearchIndex].Email;
                    }
                }

                //set the subject, start time, and end time
                tempSubject = meetingsToScheduleEventTitle;
                tempStartTime = meetingsToScheduleStartDate + " " + meetingsToScheduleStartTime;
                tempEndTime = meetingsToScheduleEndDate + " " + meetingsToScheduleEndTime;

                //create the object with the required info that was grabbed
                meetings[index] = new Meeting( tempSubject,tempLocation,tempStartTime,tempEndTime,tempEmail );

            }



            for ( int index = 0; index < meetings.Length; index++ )
            {
                meetings[index].Display();
                Console.WriteLine( "---------------------------------------------------------------------------------------------" );
            }
            Console.WriteLine( meetings.Length + " meetings ready to schedule." );

            Console.ReadKey();

            // Connects to the exchange server with currently logged in user credentials
            ExchangeService service = new ExchangeService( ExchangeVersion.Exchange2010_SP2 );
            service.UseDefaultCredentials = true;
            service.AutodiscoverUrl( "ryan.hoffer@utoledo.edu", RedirectionUrlValidationCallback );           

            for ( int index = 0; index < meetings.Length; index++ )
            {
                Appointment appointment = new Appointment(service);
                // Set the properties on the appointment object to create the appointment.
                appointment.Subject = meetings[index].Subject;
                appointment.Body = meetings[index].Subject;
                appointment.Start = meetings[index].StartTime;
                appointment.End = meetings[index].EndTime;
                appointment.Location = meetings[index].Location;
                string roomEmail = meetings[index].Email;

                // Save the appointment to the calendar.
                FolderId CalendartoSaveTo = new FolderId(WellKnownFolderName.Calendar, roomEmail);
                appointment.Save(CalendartoSaveTo, SendInvitationsMode.SendToNone);

                // Verify that the appointment was created by using the appointment's item ID.
                Item item = Item.Bind(service, appointment.Id, new PropertySet(ItemSchema.Subject));
                Console.WriteLine("\nAppointment created: " + item.Subject + "\n");
            }




            Console.Read();
            Console.Read();


        }

        private static bool RedirectionUrlValidationCallback( string redirectionUrl )
        {
            // The default for the validation callback is to reject the URL.
            bool result = false;

            Uri redirectionUri = new Uri( redirectionUrl );

            // Validate the contents of the redirection URL. In this simple validation
            // callback, the redirection URL is considered valid if it is using HTTPS
            // to encrypt the authentication credentials. 
            if ( redirectionUri.Scheme == "https" )
            {
                result = true;
            }
            return result;
        }
    }
}