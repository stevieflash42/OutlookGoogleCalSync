using Google.Apis.Auth.OAuth2;
using Google.Apis.Calendar.v3;
using Google.Apis.Calendar.v3.Data;
using Google.Apis.Requests;
using Google.Apis.Services;
using Ical.Net.CalendarComponents;
using Microsoft.Extensions.Configuration;
using Newtonsoft.Json;

class Program
{
    static async Task Main()
    {
        IConfigurationRoot configuration = new ConfigurationBuilder()
            .AddUserSecrets<Program>()
            .Build();

        string icsUrl = configuration["Outlook_ICS_URL"];
        string googleCalendarId = configuration["GoogleCalendarID"];

        CalendarService googleService = GetGoogleCalendarService(configuration["GoogleOAuth2ClientSecret"]);

        List<CalendarEvent> icalEvents = await LoadIcsEvents(icsUrl);

        string gCalTimeZone = (await googleService.Calendars.Get(googleCalendarId).ExecuteAsync()).TimeZone;

        //TODO: remove this
        await ClearGoogleCalendar(googleService, googleCalendarId);

        BatchRequest batchRequest = new BatchRequest(googleService);
        foreach (CalendarEvent ev in icalEvents)
        {
            Console.WriteLine($"{icalEvents.IndexOf(ev) + 1} / {icalEvents.Count}");
            await SyncEventToGoogle(ev, googleService, googleCalendarId, gCalTimeZone, batchRequest);
        }

        await batchRequest.ExecuteAsync();

        Console.WriteLine("Sync complete.");
    }

    static async Task<string> DownloadWithRedirect(string url)
    {
        using var handler = new HttpClientHandler { AllowAutoRedirect = true };
        using var client = new HttpClient(handler);

        client.DefaultRequestHeaders.UserAgent.ParseAdd("Mozilla/5.0 (Windows NT 10.0; Win64; x64)");

        var response = await client.GetAsync(url);
        response.EnsureSuccessStatusCode();
        return await response.Content.ReadAsStringAsync();
    }


    static async Task<List<CalendarEvent>> LoadIcsEvents(string url)
    {
        string icsContent = await DownloadWithRedirect(url);

        Ical.Net.Calendar? calendar = Ical.Net.Calendar.Load(icsContent);
        return calendar.Events
            .Where(e => !e.Summary.StartsWith("Declined:") && e.Summary != "Reminder to prep for: Platform-wide Weekly Sync")
            .ToList();
    }

    static CalendarService GetGoogleCalendarService(string strJsonClientSecret)
    {
        var jsonClientSecret = JsonConvert.DeserializeObject<GoogleClientSecrets>(strJsonClientSecret);

        UserCredential credential = GoogleWebAuthorizationBroker.AuthorizeAsync(
            jsonClientSecret.Secrets,
            [CalendarService.Scope.Calendar],
            "user",
            CancellationToken.None
        ).Result;

        return new CalendarService(new BaseClientService.Initializer
        {
            HttpClientInitializer = credential,
            ApplicationName = "ICS to Google Sync"
        });
    }

    static async Task SyncEventToGoogle(CalendarEvent icsEvent, CalendarService googleService, string calendarId, string strCalTimezone, BatchRequest batchRequest)
    {
        DateTimeOffset start = icsEvent.Start.AsUtc.ToUniversalTime();
        DateTimeOffset end = icsEvent.End.AsUtc.ToUniversalTime();

        EventDateTime startEventDateTime;
        if (icsEvent.IsAllDay)
        {
            startEventDateTime = new EventDateTime
            {
                Date = start.ToString("yyyy-MM-dd"),
                //TimeZone = strCalTimezone
            };
        }
        else
        {
            startEventDateTime = new EventDateTime
            {
                DateTimeDateTimeOffset = start,
                TimeZone = "UTC"
            };
        }

        EventDateTime endEventDateTime;
        if (icsEvent.IsAllDay)
        {
            endEventDateTime = new EventDateTime
            {
                Date = end.ToString("yyyy-MM-dd"),
                //TimeZone = strCalTimezone
            };
        }
        else
        {
            endEventDateTime = new EventDateTime
            {
                DateTimeDateTimeOffset = end,
                TimeZone = "UTC"
            };
        }


        Event gEvent = new Google.Apis.Calendar.v3.Data.Event
        {
            Summary = icsEvent.Summary,
            Description = icsEvent.Description,
            Location = icsEvent.Location,
            Start = startEventDateTime,
            End = endEventDateTime,
            //{FREQ=WEEKLY;UNTIL=20251128T150000Z;WKST=SU;BYDAY=FR}
            Recurrence = icsEvent.RecurrenceRules.Select(r => $"RRULE:{r}").ToList()
        };


        //// Optional: Try to find existing event by UID or summary+time
        //Event? existing = await FindExistingEvent(icsEvent, googleService, calendarId, start, end);

        //if (existing != null)
        //{
        //    gEvent.Id = existing.Id;
        //    batchRequest.Queue<Event>(googleService.Events.Update(gEvent, calendarId, gEvent.Id), (content, error, i, message) =>
        //    {
        //        //don't care
        //    });
        //}
        //else
        //{
        batchRequest.Queue<Event>(googleService.Events.Insert(gEvent, calendarId), (content, error, i, message) =>
        {
            //don't care
        });
        //}
    }

    static async Task ClearGoogleCalendar(
        CalendarService googleService,
        string calendarId)
    {
        HashSet<string> seenIDs = [];

        EventsResource.ListRequest request = googleService.Events.List(calendarId);
        request.SingleEvents = true;
        request.ShowDeleted = false;
        request.TimeMinDateTimeOffset = DateTimeOffset.MinValue;
        request.TimeMaxDateTimeOffset = DateTimeOffset.MaxValue;

        Events results = await request.ExecuteAsync();
        while (results.Items.Count > 0)
        {
            BatchRequest batchRequest = new BatchRequest(googleService);
            foreach (var e in results.Items)
            {
                var useThisID = e.RecurringEventId ?? e.Id;
                if (!seenIDs.Add(useThisID)) continue;
                EventsResource.DeleteRequest d = googleService.Events.Delete(calendarId, useThisID);
                batchRequest.Queue<Event>(d, (content, error, i, message) =>
                {
                    //don't care

                    if (content != null)
                    {

                    }

                    if (error != null)
                    {
                    }

                    if (message != null)
                    {
                    }
                });
            }
            await batchRequest.ExecuteAsync();

            results = await request.ExecuteAsync();
        }
    }

    static async Task<Google.Apis.Calendar.v3.Data.Event?> FindExistingEvent(
        CalendarEvent icsEvent,
        CalendarService googleService,
        string calendarId,
        DateTimeOffset start,
        DateTimeOffset end)
    {
        EventsResource.ListRequest request = googleService.Events.List(calendarId);
        request.TimeMinDateTimeOffset = start.AddMinutes(-1);
        request.TimeMaxDateTimeOffset = end.AddMinutes(1);
        request.Q = icsEvent.Summary;
        request.SingleEvents = true;

        Events results = await request.ExecuteAsync();

        return results.Items.FirstOrDefault(e =>
            e.Summary == icsEvent.Summary &&
            e.Start?.DateTimeDateTimeOffset == start &&
            e.End?.DateTimeDateTimeOffset == end
        );
    }


}
