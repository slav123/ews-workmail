package main

import (
	"fmt"
	"log"
	"time"

	"github.com/slav123/ews-workmail/ews"
)

func main() {
	// Amazon WorkMail EWS URL and credentials
	url := "https://ews.mail.us-east-1.awsapps.com/EWS/Exchange.asmx"
	username := "slav@coreadvisory.com.au"
	password := "S5FaYaBcoFy3cp"

	// Create a new EWS client with local timezone
	client := ews.NewClient(url, username, password)

	// Example: Creating a client with specific timezone
	// Uncomment to use a specific timezone instead of local timezone
	/*
		clientWithTZ, err := ews.NewClientWithTimezone(url, username, password, "America/New_York")
		if err != nil {
			log.Fatalf("Error creating client with timezone: %v", err)
		}
		// Use clientWithTZ instead of client for the rest of the code
	*/

	// You can also change the timezone of an existing client
	// Uncomment to change the timezone
	/*
		if err := client.SetTimezone("Asia/Tokyo"); err != nil {
			log.Printf("Warning: Could not set timezone: %v", err)
		}
	*/

	// Define the time range for calendar items (e.g., next 7 days)
	now := time.Now()
	endDate := now.AddDate(0, 0, 7)

	fmt.Println("Fetching calendar items from", now.Format("2006-01-02"), "to", endDate.Format("2006-01-02"))

	// Handle potential errors
	calendarItems, err := client.GetCalendarItems(now, endDate)
	if err != nil {
		log.Fatalf("Error fetching calendar items: %v", err)
	}

	// Display fetched calendar items with parsed dates
	fmt.Printf("Found %d calendar items\n", len(calendarItems))
	for i, item := range calendarItems {
		// Parse the date strings to time.Time objects
		startTime, err := client.ParseDateTime(item.Start)
		if err != nil {
			log.Printf("Warning: Could not parse start time for item %d: %v", i+1, err)
		}

		endTime, err := client.ParseDateTime(item.End)
		if err != nil {
			log.Printf("Warning: Could not parse end time for item %d: %v", i+1, err)
		}

		fmt.Printf("%d. %s\n", i+1, item.Subject)
		fmt.Printf("   Start: %s (Parsed: %s)\n", item.Start, startTime.Format(time.RFC3339))
		fmt.Printf("   End: %s (Parsed: %s)\n", item.End, endTime.Format(time.RFC3339))
		fmt.Printf("   Duration: %s\n", endTime.Sub(startTime))
		fmt.Printf("   FreeBusy: %s\n", item.LegacyFreeBusy)
		fmt.Printf("   Location: %s\n", item.Location)
		fmt.Printf("   Organizer: %s (%s)\n",
			item.Organizer.Mailbox.Name,
			item.Organizer.Mailbox.EmailAddress)
		fmt.Println()
	}

	// Example: Creating a new calendar event
	// Using timezone-aware date handling

	// Get current time in the client's timezone
	currentTime := time.Now().In(client.TimeZone)
	fmt.Printf("Current time in %v: %s\n", client.TimeZone, currentTime.Format(time.RFC3339))

	// Example: Check if a specific time slot is available (e.g., 13:00-14:00 today)
	today := currentTime.Truncate(24 * time.Hour) // Start of today
	slotStart := today.Add(13 * time.Hour)        // 13:00 today
	slotEnd := today.Add(14 * time.Hour)          // 14:00 today

	fmt.Printf("\nChecking availability for time slot: %s to %s\n",
		slotStart.Format("15:04"), slotEnd.Format("15:04"))

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
		for i, conflict := range conflicts {
			fmt.Printf("  %d. %s\n", i+1, conflict.Subject)
		}
	}

	// Example: Find available slots of 30 minutes duration in the next 8 hours
	slotDuration := 30 * time.Minute
	periodStart := currentTime
	periodEnd := currentTime.Add(8 * time.Hour)

	fmt.Printf("\nFinding available %s slots between %s and %s\n",
		slotDuration,
		periodStart.Format("15:04"),
		periodEnd.Format("15:04"))

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

	// Example: Creating a new calendar event
	newEvent := ews.CalendarEvent{
		Subject:  "Team Meeting",
		Body:     "Discussing project status and next steps",
		Start:    currentTime.Add(24 * time.Hour),
		End:      currentTime.Add(25 * time.Hour),
		Location: "Conference Room A",
		RequiredAttendees: []ews.Attendee{
			{
				Name:  "John Doe",
				Email: "slav@gex.pl",
			},
		},
		// Set to true to send invitations, false to create without sending
		SendInvites: true,
	}

	eventID, err := client.CreateCalendarEvent(newEvent)
	if err != nil {
		log.Fatalf("Error creating calendar event: %v", err)
	}
	fmt.Printf("Created new event with ID: %s\n", *eventID)

	// Example: Updating a calendar event with new fields
	fmt.Println("\nUpdating the calendar event with new subject, body, status, and attendees...")
	
	// Wait a moment before updating
	time.Sleep(2 * time.Second)

	// Define the new values
	newSubject := "Updated Team Meeting"
	newBody := "Updated agenda: Project status, next steps, and budget review"
	newStatus := "Busy"
	newLocation := "Conference Room B"

	// Define new attendees
	newRequiredAttendee := ews.Attendee{
		Name:  "Alice Smith",
		Email: "alice@example.com",
	}

	newOptionalAttendee := ews.Attendee{
		Name:  "Bob Johnson",
		Email: "bob@example.com",
	}

	// Create the new start and end times
	newStartTime := currentTime.Add(25 * time.Hour) // Move the meeting an hour later
	newEndTime := currentTime.Add(26 * time.Hour)

	// Create an updates object with the new values
	updates := ews.EventUpdates{
		Subject:           &newSubject,
		Body:              &newBody,
		LegacyFreeBusy:    &newStatus,
		Location:          &newLocation,
		RequiredAttendees: []ews.Attendee{newRequiredAttendee},
		OptionalAttendees: []ews.Attendee{newOptionalAttendee},

		// You can also update the start and end times
		Start: &newStartTime,
		End:   &newEndTime,
	}

	// Update the event
	err = client.UpdateCalendarEvent(*eventID, updates)
	if err != nil {
		log.Fatalf("Error updating calendar event: %v", err)
	}

	fmt.Println("Calendar event updated successfully!")

	// Fetch the updated event to verify changes
	fmt.Println("\nFetching updated event to verify changes...")

	// Define the time range that includes our updated event
	verifyStart := currentTime.Add(24 * time.Hour)
	verifyEnd := currentTime.Add(27 * time.Hour)

	updatedCalendarItems, err := client.GetCalendarItems(verifyStart, verifyEnd)
	if err != nil {
		log.Fatalf("Error fetching updated calendar items: %v", err)
	}

	// Display the updated event
	foundUpdated := false
	for _, item := range updatedCalendarItems {
		if item.ItemId.Id == *eventID {
			foundUpdated = true
			fmt.Println("Updated Event Details:")
			fmt.Printf("  Subject: %s\n", item.Subject)
			fmt.Printf("  Location: %s\n", item.Location)
			fmt.Printf("  FreeBusy: %s\n", item.LegacyFreeBusy)

			// Parse the date strings to time.Time objects
			startTime, _ := client.ParseDateTime(item.Start)
			endTime, _ := client.ParseDateTime(item.End)

			fmt.Printf("  Start: %s\n", startTime.Format(time.RFC3339))
			fmt.Printf("  End: %s\n", endTime.Format(time.RFC3339))

			// Note: The body content and attendees may not be included in the GetCalendarItems response
			// depending on the EWS API's response format
			break
		}
	}

	if !foundUpdated {
		fmt.Println("Could not find the updated event in the calendar items.")
	}

	// Error handling example - try to handle errors gracefully
	fmt.Println("\nExample of error handling: Trying to update a non-existent event...")
	nonExistentID := "non-existent-id"
	err = client.UpdateCalendarEvent(nonExistentID, updates)
	if err != nil {
		fmt.Printf("Expected error occurred: %v\n", err)
		// Handle the error gracefully instead of fatally exiting
	}
}
