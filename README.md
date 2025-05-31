# EWS WorkMail

A Go client library for interacting with Amazon WorkMail through the Exchange Web Services (EWS) API. This package provides a clean, idiomatic Go interface for accessing and manipulating calendar items in Amazon WorkMail.

## Features

- Retrieve calendar items within a specified date range
- Check availability for specific time slots
- Find available time slots within a date range
- Create new calendar events with attendees
- Update existing calendar events
- Delete calendar events
- Full support for required and optional attendees
- Control over whether meeting invitations are sent to attendees
- Explicit timezone handling and conversion
- Error handling for all operations

## Installation

```bash
go get github.com/slav123/ews-workmail
```

## Usage

### Creating a client

```go
import "github.com/slav123/ews-workmail/ews"

// Create a new EWS client with your WorkMail credentials (uses local timezone by default)
client := ews.NewClient(
    "https://your-workmail-domain.awsapps.com/EWS/Exchange.asmx",
    "your-email@example.com",
    "your-password",
)

// Or create a client with a specific timezone
clientWithTZ, err := ews.NewClientWithTimezone(
    "https://your-workmail-domain.awsapps.com/EWS/Exchange.asmx",
    "your-email@example.com",
    "your-password",
    "America/New_York", // IANA timezone name
)
if err != nil {
    log.Fatalf("Error creating client with timezone: %v", err)
}

// You can also change the timezone of an existing client
if err := client.SetTimezone("Europe/London"); err != nil {
    log.Printf("Warning: Could not set timezone: %v", err)
}
```

### Retrieving calendar items

```go
// Get calendar items for the next 7 days
startDate := time.Now()
endDate := startDate.AddDate(0, 0, 7)

calendarItems, err := client.GetCalendarItems(startDate, endDate)
if err != nil {
    log.Fatalf("Error fetching calendar items: %v", err)
}

// Process the retrieved calendar items
for _, item := range calendarItems {
    // Parse the date strings into time.Time objects with proper timezone handling
    startTime, err := client.ParseDateTime(item.Start)
    if err != nil {
        log.Printf("Warning: Could not parse start time: %v", err)
    }
    
    endTime, err := client.ParseDateTime(item.End)
    if err != nil {
        log.Printf("Warning: Could not parse end time: %v", err)
    }
    
    fmt.Printf("Subject: %s\n", item.Subject)
    fmt.Printf("Start: %s (Parsed: %s)\n", item.Start, startTime.Format(time.RFC3339))
    fmt.Printf("End: %s (Parsed: %s)\n", item.End, endTime.Format(time.RFC3339))
    fmt.Printf("Duration: %s\n", endTime.Sub(startTime))
    fmt.Printf("Location: %s\n", item.Location)
}
```

### Creating a calendar event

```go
// Create a new calendar event
newEvent := ews.CalendarEvent{
    Subject:  "Team Meeting",
    Body:     "Discussing project status and next steps",
    Start:    time.Now().Add(24 * time.Hour),
    End:      time.Now().Add(25 * time.Hour),
    Location: "Conference Room A",
    RequiredAttendees: []ews.Attendee{
        {
            Name:  "John Doe",
            Email: "john@example.com",
        },
    },
    OptionalAttendees: []ews.Attendee{
        {
            Name:  "Jane Smith",
            Email: "jane@example.com",
        },
    },
    SendInvites: true, // Set to false to create the event without sending invitations
}

eventID, err := client.CreateCalendarEvent(newEvent)
if err != nil {
    log.Fatalf("Error creating calendar event: %v", err)
}
fmt.Printf("Created new event with ID: %s\n", *eventID)
```

### Updating a calendar event

```go
// Update an existing calendar event's start and end time
tomorrow := time.Now().Add(24 * time.Hour)
dayAfterTomorrow := time.Now().Add(48 * time.Hour)

updates := ews.EventUpdates{
    Start: &tomorrow,
    End:   &dayAfterTomorrow,
}

err := client.UpdateCalendarEvent("event-id-here", updates)
if err != nil {
    log.Fatalf("Error updating calendar event: %v", err)
}
```

### Deleting a calendar event

```go
// Delete a calendar event by its ID
err := client.DeleteCalendarEvent("event-id-here")
if err != nil {
    log.Fatalf("Error deleting calendar event: %v", err)
}
```

### Checking calendar slot availability

```go
// Check if a specific time slot is available (e.g., 13:00-14:00 today)
today := time.Now().Truncate(24 * time.Hour) // Start of today
slotStart := today.Add(13 * time.Hour)        // 13:00 today
slotEnd := today.Add(14 * time.Hour)          // 14:00 today

isAvailable, conflicts, err := client.CheckSlotAvailability(ews.TimeSlot{
	Start: slotStart,
	End:   slotEnd,
})

if err != nil {
	log.Printf("Error checking slot availability: %v\n", err)
} else if isAvailable {
	fmt.Println("The time slot is available!")
} else {
	fmt.Printf("The time slot is not available. Found %d conflicts:\n", len(conflicts))
	for _, conflict := range conflicts {
		fmt.Printf("  %s\n", conflict.Subject)
	}
}
```

### Finding available time slots

```go
// Find available slots of 30 minutes duration in the next 8 hours
slotDuration := 30 * time.Minute
periodStart := time.Now()
periodEnd := periodStart.Add(8 * time.Hour)

availableSlots, err := client.GetAvailableSlots(periodStart, periodEnd, slotDuration)
if err != nil {
	log.Printf("Error finding available slots: %v\n", err)
} else {
	fmt.Printf("Found %d available slots:\n", len(availableSlots))
	for i, slot := range availableSlots {
		fmt.Printf("  %d. %s - %s (%s)\n", 
			i+1, 
			slot.Start.Format("15:04"), 
			slot.End.Format("15:04"),
			slot.End.Sub(slot.Start))
	}
}
```

## Timezone Handling

The library provides explicit timezone handling to ensure consistent date and time management across different environments:

- **Client-Level Timezone Configuration**: Each EWSClient instance has a dedicated timezone setting
- **Timezone-Aware Formatting**: All date/time values are consistently formatted with the client's timezone
- **Parsing with Timezone Preservation**: Response date strings are parsed back to time.Time objects with timezone context preserved
- **Runtime Timezone Changes**: Change the client's timezone at any point with the `SetTimezone` method

### Timezone Methods

```go
// Format a time.Time with timezone offset for EWS API
dateStr := client.FormatDateWithTZ(myTime) 

// Format a time.Time without timezone info but in client's timezone
dateStr := client.FormatDateWithoutTZ(myTime) 

// Parse a datetime string from EWS response with timezone handling
timeObj, err := client.ParseDateTime(dateStr)

// Change the client's timezone
err := client.SetTimezone("Europe/Paris")
```

## Error Handling

All methods in this library return detailed error messages that include information about what went wrong. Always check for errors when calling library methods.

## Amazon WorkMail EWS URL Format

The EWS URL for Amazon WorkMail typically follows this format:
```
https://[organization-alias].awsapps.com/EWS/Exchange.asmx
```

Replace `[organization-alias]` with your WorkMail organization alias.

## Contributing

Contributions are welcome! Please feel free to submit a Pull Request.

## License

This project is licensed under the MIT License - see the LICENSE file for details.
