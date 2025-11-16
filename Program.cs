using Google.Apis.Auth.OAuth2;
using Google.Apis.Calendar.v3;
using Google.Apis.Calendar.v3.Data;
using Google.Apis.Requests;
using Google.Apis.Services;
using Ical.Net.CalendarComponents;
using Ical.Net.DataTypes;
using Microsoft.Extensions.Configuration;
using Newtonsoft.Json;
using System;
using System.Linq;
using System.Text.RegularExpressions;

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

        //await ClearGoogleCalendar(googleService, googleCalendarId);

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
            if (gCalLookup.TryGetValue(kvp.Key, out Event? gEvent))
            {
                UpdateGoogleEvent(gEvent, kvp.Value, googleService, googleCalendarId, gCalTimeZone, batch);
                continue;
            }

            CreateGoogleEvent(kvp.Value, googleService, googleCalendarId, gCalTimeZone, batch);
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
            start = $"{e.Start.AsUtc:u}"; start = $"{e.Start.AsUtc:u}";
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

    private static EventDateTime GetEventDateTimeFromCalDateTime(CalDateTime calDateTime, bool isAllDay)
    {
        if (isAllDay)
        {
            return new EventDateTime
            {
                Date = calDateTime.AsUtc.ToUniversalTime().ToString("yyyy-MM-dd")
            };
        }
        else
        {
            // Use the local datetime value from the iCal event
            DateTime dt = calDateTime.Value;
            string windowsTimeZoneId = calDateTime.TzId ?? "UTC";
            string ianaTimeZoneId = windowsTimeZoneId;

            // Try to convert Windows timezone ID to IANA format for Google Calendar
            if (!string.IsNullOrEmpty(calDateTime.TzId) &&
                TimeZoneInfo.TryConvertWindowsIdToIanaId(calDateTime.TzId, out string? ianaId))
            {
                ianaTimeZoneId = ianaId;
            }

            // Get the TimeZoneInfo to calculate the correct offset for this specific datetime
            TimeZoneInfo timeZone = TimeZoneInfo.FindSystemTimeZoneById(windowsTimeZoneId);

            // Get the UTC offset for this specific datetime (accounts for DST)
            TimeSpan offset = timeZone.GetUtcOffset(dt);

            // Create DateTimeOffset with the local time and correct offset
            DateTimeOffset dateTimeOffset = new DateTimeOffset(dt, offset);

            return new EventDateTime
            {
                DateTimeDateTimeOffset = dateTimeOffset,
                TimeZone = ianaTimeZoneId
            };
        }
    }

    static void UpdateGoogleEvent(Event gEvent, CalendarEvent icsEvent, CalendarService googleService, string calendarId, string strCalTimezone, BatchRequest batchRequest)
    {
        EventDateTime startEventDateTime = GetEventDateTimeFromCalDateTime(icsEvent.Start, icsEvent.IsAllDay);
        EventDateTime endEventDateTime = GetEventDateTimeFromCalDateTime(icsEvent.End, icsEvent.IsAllDay);

        List<string> recurrence = BuildRecurrenceList(icsEvent);

        bool bUpdated = false;

        //shouldn't happen??
        if (gEvent.Summary != icsEvent.Summary)
        {
            gEvent.Summary = icsEvent.Summary;
            bUpdated = true;
        }

        if (gEvent.Description != icsEvent.Description)
        {
            // Google Calendar may truncate descriptions, so check if this is a real change
            bool isRealChange = true;

            // If both are non-null, check if Google's version is a truncated version of the iCal description
            if (gEvent.Description != null && icsEvent.Description != null)
            {
                // Google Calendar truncates descriptions at around 8192 characters
                // If the Google description matches the start of the iCal description, it's likely truncated
                if (icsEvent.Description.StartsWith(gEvent.Description) &&
                    gEvent.Description.Length >= 8000) // Close to the limit
                {
                    isRealChange = false;
                }
            }

            if (isRealChange)
            {
                gEvent.Description = icsEvent.Description;
                bUpdated = true;
            }
        }

        if (gEvent.Location != icsEvent.Location &&
            string.IsNullOrWhiteSpace(gEvent.Location) !=
            string.IsNullOrWhiteSpace(icsEvent.Location))
        {
            gEvent.Location = icsEvent.Location;
            bUpdated = true;
        }

        // Compare DateTimeDateTimeOffset values and TimeZone
        if (gEvent.Start.DateTimeDateTimeOffset != startEventDateTime.DateTimeDateTimeOffset ||
            gEvent.Start.TimeZone != startEventDateTime.TimeZone)
        {
            gEvent.Start = startEventDateTime;
            bUpdated = true;
        }

        if (gEvent.End.DateTimeDateTimeOffset != endEventDateTime.DateTimeDateTimeOffset ||
            gEvent.End.TimeZone != endEventDateTime.TimeZone)
        {
            gEvent.End = endEventDateTime;
            bUpdated = true;
        }

        bool gEventHasRecurrence = gEvent.Recurrence != null && gEvent.Recurrence.Count != 0;
        bool iCalEventHasRecurrence = recurrence.Count != 0;
        if (gEventHasRecurrence != iCalEventHasRecurrence ||
            (gEventHasRecurrence && gEvent.Recurrence.Count != recurrence.Select(GetOrderedRecurrenceString).Union(gEvent.Recurrence.Select(GetOrderedRecurrenceString)).Count()))
        {
            gEvent.Recurrence = recurrence;
            bUpdated = true;

            ////https://www.nylas.com/blog/calendar-events-rrules/
            //batchRequest.Queue<Event>(
            //    googleService.Events.Delete(calendarId, gEvent.Id),
            //    (content, error, i, message) => { /* no-op */ });

            //CreateGoogleEvent(icsEvent, googleService, calendarId, strCalTimezone, batchRequest);

            //return;
        }

        if (bUpdated)
        {
            batchRequest.Queue<Event>(googleService.Events.Update(gEvent, calendarId, gEvent.Id), (content, error, i, message) =>
            {
                if (error != null)
                {
                }
                //don't care
            });
        }
    }

    private static string GetOrderedRecurrenceString(string strRecurr)
    {
       return string.Join(";", ConvertExDateToUtc(strRecurr).Split(";").OrderBy(str => str));
    }

    /// <summary>
    /// VIBED
    /// </summary>
    /// <param name="possibleExDateLine"></param>
    /// <returns></returns>
    public static string ConvertExDateToUtc(string possibleExDateLine)
    {
        if(!possibleExDateLine.Contains("EXDATE")) return possibleExDateLine;

        // Parse the timezone from TZID parameter
        var tzidMatch = Regex.Match(possibleExDateLine, @"TZID=([^:;]+)");
        if (!tzidMatch.Success)
        {
            return possibleExDateLine;
        }

        string tzid = tzidMatch.Groups[1].Value;
        TimeZoneInfo timeZone = TimeZoneInfo.FindSystemTimeZoneById(tzid);

        // Extract the dates portion after the colon
        var datesMatch = Regex.Match(possibleExDateLine, @":(.+)$");
        if (!datesMatch.Success)
        {
            return possibleExDateLine;
        }

        string datesStr = datesMatch.Groups[1].Value;
        string[] dates = datesStr.Split(',');

        // Convert each date to UTC
        var utcDates = dates.Select(dateStr =>
        {
            // Parse the local datetime (format: yyyyMMddTHHmmss)
            int year = int.Parse(dateStr.Substring(0, 4));
            int month = int.Parse(dateStr.Substring(4, 2));
            int day = int.Parse(dateStr.Substring(6, 2));
            int hour = int.Parse(dateStr.Substring(9, 2));
            int minute = int.Parse(dateStr.Substring(11, 2));
            int second = int.Parse(dateStr.Substring(13, 2));

            // Create DateTime in the specified timezone
            DateTime localTime = new DateTime(year, month, day, hour, minute, second, DateTimeKind.Unspecified);

            // Get the UTC offset for this specific datetime (accounts for DST)
            TimeSpan offset = timeZone.GetUtcOffset(localTime);

            // Convert to UTC
            DateTimeOffset localDto = new DateTimeOffset(localTime, offset);
            DateTime utcTime = localDto.UtcDateTime;

            // Format as yyyyMMddTHHmmssZ
            return utcTime.ToString("yyyyMMdd'T'HHmmss'Z'");
        });

        // Build the result
        return $"EXDATE;VALUE=DATE:{string.Join(",", utcDates)}";
    }

    private static List<string> BuildRecurrenceList(CalendarEvent icsEvent)
    {
        string strDateFormat = icsEvent.IsAllDay ? "yyyyMMdd" : "yyyyMMddTHHmmss";
        string strZFormatVal = icsEvent.IsAllDay ? "" : "Z";

        List<string> recurrence = icsEvent.RecurrenceRules.Select(r => $"RRULE:{r}").ToList();
        string exceptionDates = string.Join(",", icsEvent.ExceptionDates.GetAllDates().Select(exd => exd.AsUtc.ToUniversalTime().ToString(strDateFormat) + strZFormatVal));

        if (!string.IsNullOrWhiteSpace(exceptionDates))
        {
            //https://www.rfc-editor.org/rfc/rfc5545
            recurrence.Add($"EXDATE;VALUE=DATE:{exceptionDates}");
        }

        return recurrence;
    }

    static void CreateGoogleEvent(CalendarEvent icsEvent, CalendarService googleService, string calendarId, string strCalTimezone, BatchRequest batchRequest)
    {
        DateTimeOffset start = icsEvent.Start.AsUtc.ToUniversalTime();
        DateTimeOffset end = icsEvent.End.AsUtc.ToUniversalTime();

        EventDateTime startEventDateTime = GetEventDateTimeFromCalDateTime(icsEvent.Start, icsEvent.IsAllDay);
        EventDateTime endEventDateTime = GetEventDateTimeFromCalDateTime(icsEvent.End, icsEvent.IsAllDay);

        List<string> recurrence = BuildRecurrenceList(icsEvent);

        Event gEvent = new()
        {
            Summary = icsEvent.Summary,
            Description = icsEvent.Description,
            Location = icsEvent.Location,
            Start = startEventDateTime,
            End = endEventDateTime,
            //{FREQ=WEEKLY;UNTIL=20251128T150000Z;WKST=SU;BYDAY=FR}
            //https://developers.google.com/workspace/calendar/api/concepts/events-calendars#recurring_events
            Recurrence = recurrence
        };

        batchRequest.Queue<Event>(googleService.Events.Insert(gEvent, calendarId), (content, error, i, message) =>
        {
            if (error != null)
            {
            }
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
            request.SingleEvents = false;
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
