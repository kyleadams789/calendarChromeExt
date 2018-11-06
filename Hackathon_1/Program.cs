using Google.Apis.Auth.OAuth2;
using Google.Apis.Calendar.v3;
using Google.Apis.Calendar.v3.Data;
using Google.Apis.Services;
using Google.Apis.Util.Store;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;

namespace CalendarQuickstart
{
    class Program
    {
        // If modifying these scopes, delete your previously saved credentials
        // at ~/.credentials/calendar-dotnet-quickstart.json
        static string[] Scopes = { CalendarService.Scope.CalendarReadonly };
        static string ApplicationName = "Google Calendar API .NET Quickstart";
        //create a list object to store the calendar events
        static List<Event> listOfEvents = new List<Event>();
        static List<String> listOfEmails = new List<String>();

        static void Main(string[] args)
        {
            UserCredential credential;

            using (var stream =
                new FileStream("credentials.json", FileMode.Open, FileAccess.Read))
            {
                string credPath = "token.json";
                credential = GoogleWebAuthorizationBroker.AuthorizeAsync(
                    GoogleClientSecrets.Load(stream).Secrets,
                    Scopes,
                    "user",
                    CancellationToken.None,
                    new FileDataStore(credPath, true)).Result;
                Console.WriteLine("Credential file saved to: " + credPath);
            }

            // Create Google Calendar API service.
            var service = new CalendarService(new BaseClientService.Initializer()
            {
                HttpClientInitializer = credential,
                ApplicationName = ApplicationName,
            });

            // Define parameters of request.
            EventsResource.ListRequest request = service.Events.List("primary");
            request.TimeMin = DateTime.Now;
            request.ShowDeleted = false;
            request.SingleEvents = true;
            request.MaxResults = 10;
            request.OrderBy = EventsResource.ListRequest.OrderByEnum.StartTime;

            // List events.
            Events events = request.Execute();
            Console.WriteLine("Upcoming events:");
            if (events.Items != null && events.Items.Count > 0)
            {
                foreach (var eventItem in events.Items)
                {
                    
                    string when = eventItem.Start.DateTime.ToString();

                    if (String.IsNullOrEmpty(when))
                        {
                            when = eventItem.Start.Date;
                            DateTime currTime = DateTime.Now;
                            DateTime appt = DateTime.Parse(when);
                        if ((currTime.Month == appt.Month && currTime.Day == appt.Day) || (currTime.AddDays(7).Day == appt.Day && currTime.AddDays(7).Month == appt.Month))
                        {
                            listOfEvents.Add(eventItem);
                            Console.WriteLine("{0} ({1})", eventItem.Summary, when,"\n");
                        }
                        else
                        {
                          Console.WriteLine("");
                        }
                    }
                    else
                    {
                        Console.WriteLine("No upcoming events found.");
                    }
                    
                }
                //List the attendees

                Console.WriteLine("\nList of Attendees: ");

                foreach (var eventItem in listOfEvents)
                {
                    int i = 0;
                    string when = eventItem.Start.DateTime.ToString();
                    if (String.IsNullOrEmpty(when))
                    {
                        when = eventItem.Start.Date;
                        DateTime currTime = DateTime.Now;
                        DateTime appt = DateTime.Parse(when);
                        if ((currTime.Month == appt.Month && currTime.Day == appt.Day) || (currTime.AddDays(7).Day == appt.Day && currTime.AddDays(7).Month == appt.Month))
                        {
                            while (i < eventItem.Attendees.Count())
                            {
                                string attendee = eventItem.Attendees[i].Email;
                                listOfEmails.Add(attendee);
                                Console.WriteLine(attendee);
                                i++;
                            }
                        }
                    }
                    else
                    {
                        Console.WriteLine("No Attendees Found");
                    }
                }
                Console.Read();
            }
        }
    }
}
