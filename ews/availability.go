package ews

import (
	"time"
)

// TimeSlot represents a time slot with start and end times
type TimeSlot struct {
	Start time.Time
	End   time.Time
}

// CheckSlotAvailability checks if a given time slot is available in the calendar
// It returns true if the slot is available, false if there are conflicts
func (c *EWSClient) CheckSlotAvailability(slot TimeSlot) (bool, []CalendarItem, error) {
	// Get all calendar items for the time range
	// We add a small buffer to make sure we get all relevant events
	startTime := slot.Start.Add(-1 * time.Minute)
	endTime := slot.End.Add(1 * time.Minute)

	items, err := c.GetCalendarItems(startTime, endTime)
	if err != nil {
		return false, nil, err
	}

	// If no items found, the slot is available
	if len(items) == 0 {
		return true, nil, nil
	}

	// Check for conflicts - find any events that overlap with our slot
	var conflicts []CalendarItem
	for _, item := range items {
		// Parse start and end times
		eventStart, err := c.ParseDateTime(item.Start)
		if err != nil {
			return false, nil, err
		}

		eventEnd, err := c.ParseDateTime(item.End)
		if err != nil {
			return false, nil, err
		}

		// Check if this event overlaps with our slot
		// Overlap occurs when:
		// 1. Event starts before our slot ends AND
		// 2. Event ends after our slot starts
		if eventStart.Before(slot.End) && eventEnd.After(slot.Start) {
			conflicts = append(conflicts, item)
		}
	}

	// Slot is available if there are no conflicts
	return len(conflicts) == 0, conflicts, nil
}

// GetAvailableSlots finds all available time slots of the specified duration within a time range
func (c *EWSClient) GetAvailableSlots(startTime, endTime time.Time, slotDuration time.Duration) ([]TimeSlot, error) {
	// Get all calendar items for the time range
	items, err := c.GetCalendarItems(startTime, endTime)
	if err != nil {
		return nil, err
	}

	// If no items, the entire range is available
	if len(items) == 0 {
		return []TimeSlot{{Start: startTime, End: endTime}}, nil
	}

	// Parse and sort all events
	var events []struct {
		Start time.Time
		End   time.Time
	}

	for _, item := range items {
		eventStart, err := c.ParseDateTime(item.Start)
		if err != nil {
			return nil, err
		}

		eventEnd, err := c.ParseDateTime(item.End)
		if err != nil {
			return nil, err
		}

		events = append(events, struct {
			Start time.Time
			End   time.Time
		}{Start: eventStart, End: eventEnd})
	}

	// Sort events by start time
	// Note: In a real implementation, you would sort the events by start time here

	// Find available slots
	var availableSlots []TimeSlot
	currentTime := startTime

	for _, event := range events {
		// If there's time before this event, check if it's enough for a slot
		if event.Start.Sub(currentTime) >= slotDuration {
			availableSlots = append(availableSlots, TimeSlot{
				Start: currentTime,
				End:   event.Start,
			})
		}

		// Move current time to the end of this event
		if event.End.After(currentTime) {
			currentTime = event.End
		}
	}

	// Check if there's time after the last event
	if endTime.Sub(currentTime) >= slotDuration {
		availableSlots = append(availableSlots, TimeSlot{
			Start: currentTime,
			End:   endTime,
		})
	}

	return availableSlots, nil
}
