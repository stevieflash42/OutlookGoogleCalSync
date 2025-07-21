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
            .SetBasePath(Directory.GetCurrentDirectory())
            .AddUserSecrets<Program>()
            .AddJsonFile("appsettings.json", optional: true, reloadOnChange: true)
            .Build();

        string icsUrl = configuration["Outlook_ICS_URL"];
        string googleCalendarId = configuration["GoogleCalendarID"];
        CalendarService googleService = GetGoogleCalendarService(configuration["GoogleOAuth2ClientSecret"]);

        List<CalendarEvent> icalEvents = await LoadIcsEvents(icsUrl);
        List<Event> gCalEvents = await GetExistingGCalEvents(googleService, googleCalendarId);
        string gCalTimeZone = (await googleService.Calendars.Get(googleCalendarId).ExecuteAsync()).TimeZone;

        // Build lookup dictionaries by synthetic key
        Dictionary<string, CalendarEvent> icalLookup = icalEvents.GroupBy(e => GenerateMatchKey(e)).ToDictionary(g => g.Key, g => g.First());
        Dictionary<string, Event> gCalLookup = gCalEvents.GroupBy(e => GenerateMatchKey(e)).ToDictionary(g => g.Key, g => g.First());

        BatchRequest batch = new(googleService);

        // INSERT missing events
        foreach (KeyValuePair<string, CalendarEvent> kvp in icalLookup)
        {
            if (gCalLookup.ContainsKey(kvp.Key))
                continue;

            SyncEventToGoogle(kvp.Value, googleService, googleCalendarId, gCalTimeZone, batch);
        }

        // DELETE orphaned events
        foreach (KeyValuePair<string, Event> kvp in gCalLookup)
        {
            if (icalLookup.ContainsKey(kvp.Key))
                continue;

            batch.Queue<Event>(
                googleService.Events.Delete(googleCalendarId, kvp.Value.Id),
                (content, error, i, message) => { /* no-op */ });
        }

        await batch.ExecuteAsync();
        Console.WriteLine("Sync complete.");
    }

    static string GenerateMatchKey(CalendarEvent e)
    {
        string start;
        string end;
        if (e.IsAllDay)
        {
            start = e.Start.ToString("yyyy-MM-dd");
            end = e.End.ToString("yyyy-MM-dd");
        }
        else
        {
            start = $"{e.Start.AsUtc:u}";start = $"{e.Start.AsUtc:u}";
            end = $"{e.End.AsUtc:u}"; end = $"{e.End.AsUtc:u}";
        }

        return $"{e.Summary}|{start}|{end}";
    }

    static string GenerateMatchKey(Event e)
    {
        string start = e.Start?.DateTimeDateTimeOffset?.ToUniversalTime().ToString("u") ?? e.Start?.Date;
        string end = e.End?.DateTimeDateTimeOffset?.ToUniversalTime().ToString("u") ?? e.End?.Date;
        return $"{e.Summary}|{start}|{end}";
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

    static void SyncEventToGoogle(CalendarEvent icsEvent, CalendarService googleService, string calendarId, string strCalTimezone, BatchRequest batchRequest)
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

        Event gEvent = new()
        {
            Summary = icsEvent.Summary,
            Description = icsEvent.Description,
            Location = icsEvent.Location,
            Start = startEventDateTime,
            End = endEventDateTime,
            //{FREQ=WEEKLY;UNTIL=20251128T150000Z;WKST=SU;BYDAY=FR}
            Recurrence = icsEvent.RecurrenceRules.Select(r => $"RRULE:{r}").ToList()
        };

        batchRequest.Queue<Event>(googleService.Events.Insert(gEvent, calendarId), (content, error, i, message) =>
        {
            //don't care
        });
    }

    static async Task<List<Event>> GetExistingGCalEvents(
        CalendarService googleService,
        string calendarId)
    {
        HashSet<string> seenIDs = [];
        List<Event> theReturn = [];

        string? strPageToken = null;

        do
        {
            EventsResource.ListRequest request = googleService.Events.List(calendarId);
            request.SingleEvents = true;
            request.ShowDeleted = false;
            request.TimeMinDateTimeOffset = DateTimeOffset.MinValue;
            request.TimeMaxDateTimeOffset = DateTimeOffset.MaxValue;
            request.PageToken = strPageToken;

            Events results = await request.ExecuteAsync();
            foreach (var e in results.Items)
            {
                var useThisID = e.RecurringEventId ?? e.Id;
                if (!seenIDs.Add(useThisID)) continue;
                theReturn.Add(e);
            }

            strPageToken = results.NextPageToken;
        }
        while (strPageToken != null);

        return theReturn;
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
            BatchRequest batchRequest = new(googleService);
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
}
