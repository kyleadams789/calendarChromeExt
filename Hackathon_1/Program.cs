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
using System.Net;
using System.Net.Mail;


namespace CalendarQuickstart
{
    class Program
    {
        // If modifying these scopes, delete your previously saved credentials
        // at ~/.credentials/calendar-dotnet-quickstart.json
        const string password = "storminmormon";
        static string[] Scopes = { CalendarService.Scope.CalendarReadonly };
        static string ApplicationName = "Google Calendar API .NET Quickstart";
        //create a list object to store the calendar events
        static List<Event> listOfEvents = new List<Event>();
        static List<String> listOfEmails = new List<String>();
        static List<Appointment> listOfAppointments = new List<Appointment>();

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
                        if ((currTime.AddDays(1).Month == appt.Month && currTime.AddDays(1).Day == appt.Day) || (currTime.AddDays(7).Day == appt.Day && currTime.AddDays(7).Month == appt.Month))
                        {                         
                            listOfEvents.Add(eventItem);
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

                foreach (var eventItem in listOfEvents)
                {
                    int i = 0;
                    string when = eventItem.Start.DateTime.ToString();
                    if (String.IsNullOrEmpty(when))
                    {
                        when = eventItem.Start.Date;
                        DateTime currTime = DateTime.Now;
                        DateTime appt = DateTime.Parse(when);
                        if ((currTime.AddDays(1).Month == appt.Month && currTime.AddDays(1).Day == appt.Day) || (currTime.AddDays(7).Day == appt.Day && currTime.AddDays(7).Month == appt.Month))
                        {

                            Console.WriteLine("{0} ({1})", eventItem.Summary, when, "\n");
                            Appointment newAppt = new Appointment();
                            newAppt.setAppt(eventItem.Summary);
                            newAppt.setDate(when);
                          
                            while (i < eventItem.Attendees.Count())
                            {
                                string attendee = eventItem.Attendees[i].Email;
                                if (attendee == "taylor.j.grover@gmail.com")
                                {
                                    i++;
                                }
                                else
                                {
                                    newAppt.fillEmails(attendee);
                                    listOfAppointments.Add(newAppt);
                                    Console.WriteLine(attendee);
                                    i++;
                                }
                            }
                        }
                    }
                    else
                    {
                        Console.WriteLine("No Attendees Found");
                    }
                }



                foreach (var apt in listOfAppointments)
                {
                    foreach (var email in apt.emailList)
                    {
                        var fromAddress = new MailAddress("taylor.j.grover@gmail.com", "Food & Care Coalition");
                        var toAddress = new MailAddress(email, "");
                        const string fromPassword = password;
                        const string subject = "Appointment Reminder";
                        string body = "Hello, I am emailing you to confirm your appointment \"" + apt.apptName + "\" on " + apt.Date + ". Please RSVP as soon as possible, I look forward to seeing you!";

                        var smtp = new SmtpClient
                        {
                            Host = "smtp.gmail.com",
                            Port = 587,
                            EnableSsl = true,
                            DeliveryMethod = SmtpDeliveryMethod.Network,
                            UseDefaultCredentials = false,
                            Credentials = new NetworkCredential(fromAddress.Address, fromPassword),
                            Timeout = 20000
                        };
                        using (var message = new MailMessage(fromAddress, toAddress)
                        {
                            Subject = subject,
                            Body = body
                        })
                        {
                            smtp.Send(message);
                        }
                    }
                }

                Console.Read();
            }
        }
    }
}
