package ews

import (
	"bytes"
	"encoding/xml"
	"fmt"
	"io"

	"net/http"
	"time"
)

// EWSClient represents a client for interacting with Amazon WorkMail EWS API
type EWSClient struct {
	URL      string
	Username string
	Password string
	Client   *http.Client
	// TimeZone location for consistent timezone handling
	TimeZone *time.Location
}

// NewClient creates a new EWS client with the provided credentials
// It uses the local timezone by default
func NewClient(url, username, password string) *EWSClient {
	return &EWSClient{
		URL:      url,
		Username: username,
		Password: password,
		Client: &http.Client{
			Timeout: 30 * time.Second,
		},
		TimeZone: time.Local, // Default to local timezone
	}
}

// NewClientWithTimezone creates a new EWS client with a specific timezone
func NewClientWithTimezone(url, username, password, timezone string) (*EWSClient, error) {
	loc, err := time.LoadLocation(timezone)
	if err != nil {
		return nil, fmt.Errorf("invalid timezone: %w", err)
	}

	return &EWSClient{
		URL:      url,
		Username: username,
		Password: password,
		Client: &http.Client{
			Timeout: 30 * time.Second,
		},
		TimeZone: loc,
	}, nil
}

// FormatDateWithTZ formats a time.Time with the client's timezone for EWS requests
func (c *EWSClient) FormatDateWithTZ(t time.Time) string {
	// Convert the time to the client's timezone
	inTZ := t.In(c.TimeZone)
	// Format with timezone offset for EWS API
	return inTZ.Format("2006-01-02T15:04:05-07:00")
}

// FormatDateWithoutTZ formats a time.Time without timezone info but in the client's timezone
func (c *EWSClient) FormatDateWithoutTZ(t time.Time) string {
	// Convert the time to the client's timezone
	inTZ := t.In(c.TimeZone)
	// Format without timezone offset
	return inTZ.Format("2006-01-02T15:04:05")
}

// ParseDateTime parses a datetime string from EWS response into a time.Time
// with the client's timezone
func (c *EWSClient) ParseDateTime(dateStr string) (time.Time, error) {
	// Try to parse with timezone first (with offset)
	t, err := time.Parse(time.RFC3339, dateStr)
	if err == nil {
		return t, nil
	}

	// Try with 'Z' timezone indicator (UTC)
	t, err = time.Parse("2006-01-02T15:04:05Z", dateStr)
	if err == nil {
		return t, nil
	}

	// Try with timezone offset without colon (e.g., -0700)
	t, err = time.Parse("2006-01-02T15:04:05-0700", dateStr)
	if err == nil {
		return t, nil
	}

	// If that fails, try without timezone and set the client's timezone
	t, err = time.Parse("2006-01-02T15:04:05", dateStr)
	if err != nil {
		return time.Time{}, fmt.Errorf("unable to parse date string '%s': %w", dateStr, err)
	}

	// Set the timezone to the client's timezone
	return time.Date(
		t.Year(), t.Month(), t.Day(),
		t.Hour(), t.Minute(), t.Second(), t.Nanosecond(),
		c.TimeZone,
	), nil
}

// SetTimezone changes the client's timezone
func (c *EWSClient) SetTimezone(timezone string) error {
	loc, err := time.LoadLocation(timezone)
	if err != nil {
		return fmt.Errorf("invalid timezone: %w", err)
	}
	c.TimeZone = loc
	return nil
}

// GetCalendarItems retrieves calendar items between the specified dates
func (c *EWSClient) GetCalendarItems(startDate, endDate time.Time) ([]CalendarItem, error) {
	// Format dates for EWS request using timezone-aware formatting
	startDateStr := c.FormatDateWithTZ(startDate)
	endDateStr := c.FormatDateWithTZ(endDate)

	// Prepare the SOAP envelope
	envelope := Envelope{
		XMLNS:  "http://schemas.xmlsoap.org/soap/envelope/",
		XMLNSt: "http://schemas.microsoft.com/exchange/services/2006/types",
		XMLNSm: "http://schemas.microsoft.com/exchange/services/2006/messages",
		Header: Header{
			ServerVersionInfo: ServerVersionInfo{
				Version: "Exchange2010",
			},
		},
		Body: Body{
			FindItem: &FindItemRequest{
				XMLNSm:    "http://schemas.microsoft.com/exchange/services/2006/messages",
				Traversal: "Shallow",
				ItemShape: ItemShape{
					BaseShape: "AllProperties",
				},
				CalendarView: CalendarView{
					StartDate: startDateStr,
					EndDate:   endDateStr,
				},
				ParentFolderIds: ParentFolderIds{
					DistinguishedFolderId: DistinguishedFolderId{
						Id: "calendar",
					},
				},
			},
		},
	}

	// Convert the envelope to XML
	xmlData, err := xml.MarshalIndent(envelope, "", "  ")
	if err != nil {
		return nil, fmt.Errorf("error marshalling request: %w", err)
	}

	// Create the HTTP request
	req, err := http.NewRequest(http.MethodPost, c.URL, bytes.NewReader(xmlData))
	if err != nil {
		return nil, fmt.Errorf("error creating request: %w", err)
	}

	// Set headers
	req.Header.Set("Content-Type", "text/xml; charset=utf-8")
	req.SetBasicAuth(c.Username, c.Password)

	// Send the request
	resp, err := c.Client.Do(req)
	if err != nil {
		return nil, fmt.Errorf("error sending request: %w", err)
	}
	defer resp.Body.Close()

	// Check response status
	if resp.StatusCode != http.StatusOK {
		body, _ := io.ReadAll(resp.Body)
		return nil, fmt.Errorf("unexpected status code: %d, body: %s", resp.StatusCode, string(body))
	}

	// Read and parse the response
	body, err := io.ReadAll(resp.Body)
	if err != nil {
		return nil, fmt.Errorf("error reading response body: %w", err)
	}

	// Parse the response
	var responseEnvelope ResponseEnvelope
	if err := xml.Unmarshal(body, &responseEnvelope); err != nil {
		return nil, fmt.Errorf("error unmarshalling response: %w", err)
	}

	// Check response code
	responseMessage := responseEnvelope.Body.FindItemResponse.ResponseMessages.FindItemResponseMessage
	if responseMessage.ResponseCode != "NoError" {
		return nil, fmt.Errorf("EWS error: %s", responseMessage.ResponseCode)
	}

	return responseMessage.RootFolder.Items.CalendarItem, nil
}

// CalendarEvent represents a calendar event to be created
type CalendarEvent struct {
	Subject           string
	Body              string
	Start             time.Time
	End               time.Time
	Location          string
	IsAllDay          bool
	RequiredAttendees []Attendee
	OptionalAttendees []Attendee
	SendInvites       bool // Controls whether meeting invitations are sent to attendees
}

// CreateCalendarEvent creates a new calendar event
func (c *EWSClient) CreateCalendarEvent(event CalendarEvent) (*string, error) {
	// Format dates using timezone-aware methods
	startStr := c.FormatDateWithoutTZ(event.Start)
	endStr := c.FormatDateWithoutTZ(event.End)

	// Prepare the SOAP envelope
	envelope := Envelope{
		XMLNS:  "http://schemas.xmlsoap.org/soap/envelope/",
		XMLNSt: "http://schemas.microsoft.com/exchange/services/2006/types",
		XMLNSm: "http://schemas.microsoft.com/exchange/services/2006/messages",
		Header: Header{
			ServerVersionInfo: ServerVersionInfo{
				Version: "Exchange2010",
			},
		},
		Body: Body{
			CreateItem: &CreateEventRequest{
				// Set SendMeetingInvitations based on the SendInvites flag
				SendMeetingInvitations: func() string {
					if event.SendInvites {
						return "SendToAllAndSaveCopy"
					}
					return "SendToNone"
				}(),
				SavedItemFolderId: SavedItemFolderId{
					DistinguishedFolderId: DistinguishedFolderId{
						Id: "calendar",
					},
				},
				Items: CreateEventItems{
					CalendarItem: CreateEventCalendarItem{
						XMLNSt:  "http://schemas.microsoft.com/exchange/services/2006/types",
						Subject: event.Subject,
						Body: ItemBody{
							BodyType: "Text",
							Content:  event.Body,
						},
						ReminderIsSet:   true,
						ReminderMinutes: 15,
						Start:           startStr,
						End:             endStr,
						IsAllDayEvent:   event.IsAllDay,
						LegacyFreeBusy:  "Busy",
						Location:        event.Location,
					},
				},
			},
		},
	}

	// Add required attendees if present
	if len(event.RequiredAttendees) > 0 {
		requiredAttendees := RequiredAttendees{
			Attendees: make([]AttendeeType, 0, len(event.RequiredAttendees)),
		}

		for _, attendee := range event.RequiredAttendees {
			requiredAttendees.Attendees = append(requiredAttendees.Attendees, AttendeeType{
				Mailbox: EmailAddress{
					Name:         attendee.Name,
					EmailAddress: attendee.Email,
					RoutingType:  "SMTP",
				},
			})
		}

		envelope.Body.CreateItem.Items.CalendarItem.RequiredAttendees = &requiredAttendees
	}

	// Add optional attendees if present
	if len(event.OptionalAttendees) > 0 {
		optionalAttendees := OptionalAttendees{
			Attendees: make([]AttendeeType, 0, len(event.OptionalAttendees)),
		}

		for _, attendee := range event.OptionalAttendees {
			optionalAttendees.Attendees = append(optionalAttendees.Attendees, AttendeeType{
				Mailbox: EmailAddress{
					Name:         attendee.Name,
					EmailAddress: attendee.Email,
					RoutingType:  "SMTP",
				},
			})
		}

		envelope.Body.CreateItem.Items.CalendarItem.OptionalAttendees = &optionalAttendees
	}

	// Convert the envelope to XML
	xmlData, err := xml.MarshalIndent(envelope, "", "  ")
	if err != nil {
		return nil, fmt.Errorf("error marshalling request: %w", err)
	}

	// Create the HTTP request
	req, err := http.NewRequest("POST", c.URL, bytes.NewReader(xmlData))
	if err != nil {
		return nil, fmt.Errorf("error creating request: %w", err)
	}

	// Set headers
	req.Header.Set("Content-Type", "text/xml; charset=utf-8")
	req.SetBasicAuth(c.Username, c.Password)

	// Send the request
	resp, err := c.Client.Do(req)
	if err != nil {
		return nil, fmt.Errorf("error sending request: %w", err)
	}
	defer resp.Body.Close()

	// Check response status
	if resp.StatusCode != http.StatusOK {
		body, _ := io.ReadAll(resp.Body)
		return nil, fmt.Errorf("unexpected status code: %d, body: %s", resp.StatusCode, string(body))
	}

	// Read and parse the response
	body, err := io.ReadAll(resp.Body)
	if err != nil {
		return nil, fmt.Errorf("error reading response body: %w", err)
	}

	// Parse the response
	var responseEnvelope CreateItemResponseEnvelope
	if err := xml.Unmarshal(body, &responseEnvelope); err != nil {
		return nil, fmt.Errorf("error unmarshalling response: %w", err)
	}

	// Check response code
	responseMessage := responseEnvelope.Body.CreateItemResponse.ResponseMessages.CreateItemResponseMessage
	if responseMessage.ResponseClass != "Success" {
		return nil, fmt.Errorf("EWS error: %s", responseMessage.ResponseCode)
	}

	// Return the ID of the created event
	if len(responseMessage.Items.CalendarItem) > 0 {
		return &responseMessage.Items.CalendarItem[0].ItemId.Id, nil
	}

	return nil, fmt.Errorf("no item ID returned")
}

// DeleteCalendarEvent deletes a calendar event by its ID
func (c *EWSClient) DeleteCalendarEvent(itemID string) error {
	// Prepare the SOAP envelope
	envelope := Envelope{
		XMLNS:  "http://schemas.xmlsoap.org/soap/envelope/",
		XMLNSt: "http://schemas.microsoft.com/exchange/services/2006/types",
		XMLNSm: "http://schemas.microsoft.com/exchange/services/2006/messages",
		Header: Header{
			ServerVersionInfo: ServerVersionInfo{
				Version: "Exchange2010",
			},
		},
		Body: Body{
			DeleteItem: &DeleteItemRequest{
				XMLNSm:                   "http://schemas.microsoft.com/exchange/services/2006/messages",
				DeleteType:               "HardDelete",
				SendMeetingCancellations: "SendToAllAndSaveCopy",
				ItemIds: DeleteItemIds{
					ItemId: []ItemId{
						{
							Id: itemID,
						},
					},
				},
			},
		},
	}

	// Convert the envelope to XML
	xmlData, err := xml.MarshalIndent(envelope, "", "  ")
	if err != nil {
		return fmt.Errorf("error marshalling request: %w", err)
	}

	// Create the HTTP request
	req, err := http.NewRequest(http.MethodPost, c.URL, bytes.NewReader(xmlData))
	if err != nil {
		return fmt.Errorf("error creating request: %w", err)
	}

	// Set headers
	req.Header.Set("Content-Type", "text/xml; charset=utf-8")
	req.SetBasicAuth(c.Username, c.Password)

	// Send the request
	resp, err := c.Client.Do(req)
	if err != nil {
		return fmt.Errorf("error sending request: %w", err)
	}
	defer resp.Body.Close()

	// Check response status
	if resp.StatusCode != http.StatusOK {
		body, _ := io.ReadAll(resp.Body)
		return fmt.Errorf("unexpected status code: %d, body: %s", resp.StatusCode, string(body))
	}

	return nil
}

// EventUpdates represents updates to an existing calendar event
type EventUpdates struct {
	Start             *time.Time
	End               *time.Time
	Subject           *string
	Body              *string
	LegacyFreeBusy    *string
	Location          *string
	RequiredAttendees []Attendee
	OptionalAttendees []Attendee
}

// UpdateCalendarEvent updates a calendar event by its ID
func (c *EWSClient) UpdateCalendarEvent(itemID string, updates EventUpdates) error {
	// Prepare the SOAP envelope
	envelope := Envelope{
		XMLNS:  "http://schemas.xmlsoap.org/soap/envelope/",
		XMLNSt: "http://schemas.microsoft.com/exchange/services/2006/types",
		XMLNSm: "http://schemas.microsoft.com/exchange/services/2006/messages",
		Header: Header{
			ServerVersionInfo: ServerVersionInfo{
				Version: "Exchange2010",
			},
		},
		Body: Body{
			UpdateItem: &UpdateItemRequest{
				XMLNSm:                 "http://schemas.microsoft.com/exchange/services/2006/messages",
				ConflictResolution:     "AlwaysOverwrite",
				SendMeetingInvitations: "SendToAllAndSaveCopy",
				MessageDisposition:     "SaveOnly",
				ItemChanges: ItemChanges{
					ItemChange: ItemChange{
						ItemId: ItemId{
							Id: itemID,
						},
						Updates: Updates{
							SetItemField: []SetItemField{},
						},
					},
				},
			},
		},
	}

	// Add the updates
	if updates.Start != nil {
		// Format with timezone-aware method
		startStr := c.FormatDateWithoutTZ(*updates.Start)
		envelope.Body.UpdateItem.ItemChanges.ItemChange.Updates.SetItemField = append(
			envelope.Body.UpdateItem.ItemChanges.ItemChange.Updates.SetItemField,
			SetItemField{
				FieldURI: FieldURI{
					FieldURI: "calendar:Start",
				},
				CalendarItem: UpdateCalendarItem{
					Start: &startStr,
				},
			},
		)
	}

	if updates.End != nil {
		// Format with timezone-aware method
		endStr := c.FormatDateWithoutTZ(*updates.End)
		envelope.Body.UpdateItem.ItemChanges.ItemChange.Updates.SetItemField = append(
			envelope.Body.UpdateItem.ItemChanges.ItemChange.Updates.SetItemField,
			SetItemField{
				FieldURI: FieldURI{
					FieldURI: "calendar:End",
				},
				CalendarItem: UpdateCalendarItem{
					End: &endStr,
				},
			},
		)
	}

	// Add Subject update if provided
	if updates.Subject != nil {
		envelope.Body.UpdateItem.ItemChanges.ItemChange.Updates.SetItemField = append(
			envelope.Body.UpdateItem.ItemChanges.ItemChange.Updates.SetItemField,
			SetItemField{
				FieldURI: FieldURI{
					FieldURI: "item:Subject",
				},
				CalendarItem: UpdateCalendarItem{
					Subject: updates.Subject,
				},
			},
		)
	}

	// Add Body (notes) update if provided
	if updates.Body != nil {
		envelope.Body.UpdateItem.ItemChanges.ItemChange.Updates.SetItemField = append(
			envelope.Body.UpdateItem.ItemChanges.ItemChange.Updates.SetItemField,
			SetItemField{
				FieldURI: FieldURI{
					FieldURI: "item:Body",
				},
				CalendarItem: UpdateCalendarItem{
					Body: &ItemBody{
						BodyType: "Text",
						Content:  *updates.Body,
					},
				},
			},
		)
	}

	// Add LegacyFreeBusy update if provided
	if updates.LegacyFreeBusy != nil {
		envelope.Body.UpdateItem.ItemChanges.ItemChange.Updates.SetItemField = append(
			envelope.Body.UpdateItem.ItemChanges.ItemChange.Updates.SetItemField,
			SetItemField{
				FieldURI: FieldURI{
					FieldURI: "calendar:LegacyFreeBusyStatus",
				},
				CalendarItem: UpdateCalendarItem{
					LegacyFreeBusy: updates.LegacyFreeBusy,
				},
			},
		)
	}

	// Add Location update if provided
	if updates.Location != nil {
		envelope.Body.UpdateItem.ItemChanges.ItemChange.Updates.SetItemField = append(
			envelope.Body.UpdateItem.ItemChanges.ItemChange.Updates.SetItemField,
			SetItemField{
				FieldURI: FieldURI{
					FieldURI: "calendar:Location",
				},
				CalendarItem: UpdateCalendarItem{
					Location: updates.Location,
				},
			},
		)
	}

	// Add Required Attendees if provided
	if len(updates.RequiredAttendees) > 0 {
		requiredAttendees := RequiredAttendees{
			Attendees: make([]AttendeeType, len(updates.RequiredAttendees)),
		}

		for i, attendee := range updates.RequiredAttendees {
			requiredAttendees.Attendees[i] = AttendeeType{
				Mailbox: EmailAddress{
					Name:         attendee.Name,
					EmailAddress: attendee.Email,
					RoutingType:  "SMTP",
				},
			}
		}

		envelope.Body.UpdateItem.ItemChanges.ItemChange.Updates.SetItemField = append(
			envelope.Body.UpdateItem.ItemChanges.ItemChange.Updates.SetItemField,
			SetItemField{
				FieldURI: FieldURI{
					FieldURI: "calendar:RequiredAttendees",
				},
				CalendarItem: UpdateCalendarItem{
					RequiredAttendees: &requiredAttendees,
				},
			},
		)
	}

	// Add Optional Attendees if provided
	if len(updates.OptionalAttendees) > 0 {
		optionalAttendees := OptionalAttendees{
			Attendees: make([]AttendeeType, len(updates.OptionalAttendees)),
		}

		for i, attendee := range updates.OptionalAttendees {
			optionalAttendees.Attendees[i] = AttendeeType{
				Mailbox: EmailAddress{
					Name:         attendee.Name,
					EmailAddress: attendee.Email,
					RoutingType:  "SMTP",
				},
			}
		}

		envelope.Body.UpdateItem.ItemChanges.ItemChange.Updates.SetItemField = append(
			envelope.Body.UpdateItem.ItemChanges.ItemChange.Updates.SetItemField,
			SetItemField{
				FieldURI: FieldURI{
					FieldURI: "calendar:OptionalAttendees",
				},
				CalendarItem: UpdateCalendarItem{
					OptionalAttendees: &optionalAttendees,
				},
			},
		)
	}

	// Convert the envelope to XML
	xmlData, err := xml.MarshalIndent(envelope, "", "  ")
	if err != nil {
		return fmt.Errorf("error marshalling request: %w", err)
	}

	// Create the HTTP request
	req, err := http.NewRequest("POST", c.URL, bytes.NewReader(xmlData))
	if err != nil {
		return fmt.Errorf("error creating request: %w", err)
	}

	// Set headers
	req.Header.Set("Content-Type", "text/xml; charset=utf-8")
	req.SetBasicAuth(c.Username, c.Password)

	// Send the request
	resp, err := c.Client.Do(req)
	if err != nil {
		return fmt.Errorf("error sending request: %w", err)
	}
	defer resp.Body.Close()

	// Check response status
	if resp.StatusCode != http.StatusOK {
		body, _ := io.ReadAll(resp.Body)
		return fmt.Errorf("unexpected status code: %d, body: %s", resp.StatusCode, string(body))
	}

	// Read and parse the response
	body, err := io.ReadAll(resp.Body)
	if err != nil {
		return fmt.Errorf("error reading response body: %w", err)
	}

	// Parse the response
	var responseEnvelope UpdateItemResponseEnvelope
	if err := xml.Unmarshal(body, &responseEnvelope); err != nil {
		return fmt.Errorf("error unmarshalling response: %w", err)
	}

	// Check response code
	responseMessage := responseEnvelope.Body.UpdateItemResponse.ResponseMessages.UpdateItemResponseMessage
	if responseMessage.ResponseClass != "Success" {
		return fmt.Errorf("EWS error: %s", responseMessage.ResponseCode)
	}

	return nil
}
