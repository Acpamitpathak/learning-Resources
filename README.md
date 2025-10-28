
using Outlook = Microsoft.Office.Interop.Outlook;

public static void AddEventToOutlook(string subject, string body, DateTime start, DateTime end, string location, string requiredAttendees)
{
    try
    {
        // Create Outlook instance
        Outlook.Application outlookApp = new Outlook.Application();

        // Create a new appointment item
        Outlook.AppointmentItem appointment = (Outlook.AppointmentItem)
            outlookApp.CreateItem(Outlook.OlItemType.olAppointmentItem);

        // Set event details
        appointment.Subject = subject;
        appointment.Body = body;
        appointment.Location = location;
        appointment.Start = start;
        appointment.End = end;
        appointment.ReminderSet = true;
        appointment.ReminderMinutesBeforeStart = 15;

        // Mark as a meeting invite
        appointment.MeetingStatus = Outlook.OlMeetingStatus.olMeeting;

        // Add recipients
        appointment.RequiredAttendees = requiredAttendees;

        // Save and send automatically
        appointment.Save();   // saves directly to your Outlook Calendar
        appointment.Send();   // sends invite to attendees

        Console.WriteLine("✅ Meeting invite added and sent successfully!");
    }
    catch (Exception ex)
    {
        Console.WriteLine("❌ Error adding event to Outlook: " + ex.Message);
    }
}



# learning-Resources
React :  
https://www.youtube.com/watch?v=TtPXvEcE11E
https://www.youtube.com/watch?v=x4rFhThSX04&t=18061s
