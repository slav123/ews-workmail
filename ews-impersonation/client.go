package ewsimpersonation

import (
	"bytes"
	"context"
	"encoding/xml"
	"fmt"
	"io"
	"log"
	"net/http"
	"sync"
	"time"

	"github.com/aws/aws-sdk-go-v2/aws"
	"github.com/aws/aws-sdk-go-v2/config"
	"github.com/aws/aws-sdk-go-v2/service/workmail"
)

const (
	// tokenRefreshBuffer is a buffer to proactively refresh the token before it expires.
	tokenRefreshBuffer = 5 * time.Minute
	// defaultEWSVersion is the EWS schema version to target.
	defaultEWSVersion = "Exchange2010_SP2" // Or another version as appropriate
)

// ImpersonationClient facilitates EWS calls using AWS WorkMail impersonation.
type ImpersonationClient struct {
	awsRegion           string
	workmailOrgID       string
	impersonationRoleID string
	ewsEndpoint         string

	httpClient *http.Client
	timeZone   *time.Location

	awsWorkMailClient *workmail.Client

	tokenLock    sync.Mutex
	currentToken *string
	tokenExpiry  time.Time
}

// NewImpersonationClient creates a new client for EWS with impersonation.
// It loads AWS credentials using the default chain (environment, shared credentials, IAM roles).
func NewImpersonationClient(ctx context.Context, awsRegion, workmailOrgID, impersonationRoleID, ewsEndpoint string) (*ImpersonationClient, error) {
	cfg, err := config.LoadDefaultConfig(ctx, config.WithRegion(awsRegion))
	if err != nil {
		return nil, fmt.Errorf("failed to load AWS config: %w", err)
	}

	return NewImpersonationClientWithAWSConfig(cfg, awsRegion, workmailOrgID, impersonationRoleID, ewsEndpoint)
}

// NewImpersonationClientWithAWSConfig creates a new client using a provided AWS config.
func NewImpersonationClientWithAWSConfig(cfg aws.Config, awsRegion, workmailOrgID, impersonationRoleID, ewsEndpoint string) (*ImpersonationClient, error) {
	wmClient := workmail.NewFromConfig(cfg)

	client := &ImpersonationClient{
		awsRegion:           awsRegion,
		workmailOrgID:       workmailOrgID,
		impersonationRoleID: impersonationRoleID,
		ewsEndpoint:         ewsEndpoint,
		httpClient: &http.Client{
			Timeout: 30 * time.Second,
		},
		timeZone:          time.Local, // Default to local timezone
		awsWorkMailClient: wmClient,
	}

	// Optionally, load a specific timezone if needed, similar to original client
	// loc, err := time.LoadLocation("America/New_York")
	// if err == nil {
	// 	client.timeZone = loc
	// }

	return client, nil
}

// getToken retrieves a valid EWS access token, refreshing if necessary.
func (c *ImpersonationClient) getToken(ctx context.Context) (string, error) {
	c.tokenLock.Lock()
	defer c.tokenLock.Unlock()

	if c.currentToken != nil && time.Now().Before(c.tokenExpiry.Add(-tokenRefreshBuffer)) {
		return *c.currentToken, nil
	}

	log.Println("Refreshing AWS WorkMail EWS impersonation token...")
	resp, err := c.awsWorkMailClient.AssumeImpersonationRole(ctx, &workmail.AssumeImpersonationRoleInput{
		OrganizationId:      &c.workmailOrgID,
		ImpersonationRoleId: &c.impersonationRoleID,
	})
	if err != nil {
		return "", fmt.Errorf("failed to assume impersonation role: %w", err)
	}

	if resp.Token == nil || resp.ExpiresIn == nil {
		return "", fmt.Errorf("AssumeImpersonationRole response missing token or expiry")
	}

	c.currentToken = resp.Token
	c.tokenExpiry = time.Now().Add(time.Duration(*resp.ExpiresIn) * time.Second)
	log.Printf("Successfully refreshed EWS token. Expires at: %s", c.tokenExpiry.Format(time.RFC3339))

	return *c.currentToken, nil
}

// doRequest performs the actual EWS request.
func (c *ImpersonationClient) doRequest(ctx context.Context, soapAction, targetUserEmail string, requestBody interface{}, responseBody interface{}) error {
	token, err := c.getToken(ctx)
	if err != nil {
		return fmt.Errorf("failed to get EWS token: %w", err)
	}

	envelope := Envelope{
		XMLNS:  "http://schemas.xmlsoap.org/soap/envelope/",
		XMLNSt: "http://schemas.microsoft.com/exchange/services/2006/types",
		XMLNSm: "http://schemas.microsoft.com/exchange/services/2006/messages",
		Header: Header{
			ServerVersionInfo: ServerVersionInfo{
				Version: defaultEWSVersion,
			},
			ExchangeImpersonation: &ExchangeImpersonationType{
				ConnectingSID: ConnectingSIDType{
					PrimarySmtpAddress: targetUserEmail,
				},
			},
		},
		Body: Body{},
	}

	// Dynamically set the correct field in the Body struct based on the requestBody type
	switch r := requestBody.(type) {
	case *FindItemRequest:
		envelope.Body.FindItem = r
	case *CreateEventRequest:
		envelope.Body.CreateItem = r
	case *DeleteItemRequest:
		envelope.Body.DeleteItem = r
	case *UpdateItemRequest:
		envelope.Body.UpdateItem = r
	default:
		return fmt.Errorf("unsupported request body type: %T", requestBody)
	}

	xmlData, err := xml.MarshalIndent(envelope, "", "  ")
	if err != nil {
		return fmt.Errorf("error marshalling EWS request: %w", err)
	}

	// Log the XML for UpdateItem requests for debugging
	// if soapAction == "http://schemas.microsoft.com/exchange/services/2006/messages/UpdateItem" {
	// 	log.Printf("DEBUG: EWS UpdateItem Request XML:\n%s\n", string(xmlData))
	// }

	req, err := http.NewRequestWithContext(ctx, "POST", c.ewsEndpoint, bytes.NewReader(xmlData))
	if err != nil {
		return fmt.Errorf("error creating EWS HTTP request: %w", err)
	}

	req.Header.Set("Content-Type", "text/xml; charset=utf-8")
	req.Header.Set("SOAPAction", soapAction)
	req.Header.Set("Authorization", "Bearer "+token)

	resp, err := c.httpClient.Do(req)
	if err != nil {
		return fmt.Errorf("error sending EWS request: %w", err)
	}
	defer resp.Body.Close()

	bodyBytes, err := io.ReadAll(resp.Body)
	if err != nil {
		return fmt.Errorf("error reading EWS response body: %w", err)
	}

	if resp.StatusCode != http.StatusOK {
		return fmt.Errorf("EWS request failed with status %d: %s", resp.StatusCode, string(bodyBytes))
	}

	if err := xml.Unmarshal(bodyBytes, responseBody); err != nil {
		return fmt.Errorf("error unmarshalling EWS response: %w. Response body: %s", err, string(bodyBytes))
	}

	return nil
}

// FormatDateWithTZ formats a time.Time with the client's timezone for EWS requests
func (c *ImpersonationClient) FormatDateWithTZ(t time.Time) string {
	inTZ := t.In(c.timeZone)
	return inTZ.Format("2006-01-02T15:04:05-07:00")
}

// ParseDateTime parses a datetime string from EWS response into a time.Time
func (c *ImpersonationClient) ParseDateTime(dateStr string) (time.Time, error) {
	formats := []string{
		time.RFC3339,
		"2006-01-02T15:04:05Z", // UTC
		"2006-01-02T15:04:05",  // No timezone, assume client's timezone
	}

	var t time.Time
	var err error

	for _, format := range formats {
		t, err = time.Parse(format, dateStr)
		if err == nil {
			if format == "2006-01-02T15:04:05" { // If no timezone info, set to client's timezone
				return time.Date(t.Year(), t.Month(), t.Day(), t.Hour(), t.Minute(), t.Second(), t.Nanosecond(), c.timeZone), nil
			}
			return t.In(c.timeZone), nil // Ensure it's in client's timezone
		}
	}
	return time.Time{}, fmt.Errorf("unable to parse date string '%s' with known formats: %w", dateStr, err)
}

// SetTimezone changes the client's timezone
func (c *ImpersonationClient) SetTimezone(timezone string) error {
	loc, err := time.LoadLocation(timezone)
	if err != nil {
		return fmt.Errorf("invalid timezone: %w", err)
	}
	c.timeZone = loc
	return nil
}

// GetCalendarItems retrieves calendar items for the target user between the specified dates.
func (c *ImpersonationClient) GetCalendarItems(ctx context.Context, startDate, endDate time.Time, targetUserEmail string) ([]CalendarItem, error) {
	startDateStr := c.FormatDateWithTZ(startDate)
	endDateStr := c.FormatDateWithTZ(endDate)

	request := &FindItemRequest{
		XMLNSm:    "http://schemas.microsoft.com/exchange/services/2006/messages",
		Traversal: "Shallow",
		ItemShape: ItemShape{
			BaseShape: "AllProperties", // Or "IdOnly", "Default"
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
	}

	var responseEnvelope ResponseEnvelope
	soapAction := "http://schemas.microsoft.com/exchange/services/2006/messages/FindItem"

	err := c.doRequest(ctx, soapAction, targetUserEmail, request, &responseEnvelope)
	if err != nil {
		return nil, err
	}

	respMsg := responseEnvelope.Body.FindItemResponse.ResponseMessages.FindItemResponseMessage
	if respMsg.ResponseCode != "NoError" {
		return nil, fmt.Errorf("EWS error in GetCalendarItems: %s", respMsg.ResponseCode)
	}

	// Convert internal CalendarItem types to a more usable format if needed, or return directly.
	// For now, returning the direct parsed items.
	// Note: The CalendarItem struct in types.go will need Start/End parsed to time.Time by the caller or a helper.
	return respMsg.RootFolder.Items.CalendarItem, nil
}

// CreateCalendarEvent creates a new calendar event for the target user.
// sendMeetingInvitations can be "SendToNone", "SendOnlyToAll", "SendToAllAndSaveCopy".
func (c *ImpersonationClient) CreateCalendarEvent(ctx context.Context, event CalendarEvent, sendMeetingInvitations string, targetUserEmail string) (*ItemId, error) {
	xmlNSt := "http://schemas.microsoft.com/exchange/services/2006/types"

	calItem := CreateEventCalendarItem{
		XMLNSt:          xmlNSt,
		Subject:         event.Subject,
		Body:            ItemBody{BodyType: "Text", Content: event.Body}, // Assuming Text body type
		ReminderIsSet:   true,                                            // Default, can be made configurable
		ReminderMinutes: 15,                                              // Default, can be made configurable
		Start:           c.FormatDateWithTZ(event.Start),
		End:             c.FormatDateWithTZ(event.End),
		IsAllDayEvent:   event.IsAllDay,
		LegacyFreeBusy:  "Busy", // Default, can be made configurable (e.g. Free, Tentative, OOF)
		Location:        event.Location,
	}

	if len(event.RequiredAttendees) > 0 {
		calItem.RequiredAttendees = &RequiredAttendees{}
		for _, ra := range event.RequiredAttendees {
			calItem.RequiredAttendees.Attendees = append(calItem.RequiredAttendees.Attendees, AttendeeType{
				Mailbox: EmailAddress{Name: ra.Name, EmailAddress: ra.Email, RoutingType: "SMTP"},
			})
		}
	}
	if len(event.OptionalAttendees) > 0 {
		calItem.OptionalAttendees = &OptionalAttendees{}
		for _, oa := range event.OptionalAttendees {
			calItem.OptionalAttendees.Attendees = append(calItem.OptionalAttendees.Attendees, AttendeeType{
				Mailbox: EmailAddress{Name: oa.Name, EmailAddress: oa.Email, RoutingType: "SMTP"},
			})
		}
	}

	request := &CreateEventRequest{
		SendMeetingInvitations: sendMeetingInvitations,
		SavedItemFolderId: SavedItemFolderId{
			DistinguishedFolderId: DistinguishedFolderId{Id: "calendar"},
		},
		Items: CreateEventItems{
			CalendarItem: calItem,
		},
	}

	var responseEnvelope CreateItemResponseEnvelope
	soapAction := "http://schemas.microsoft.com/exchange/services/2006/messages/CreateItem"

	err := c.doRequest(ctx, soapAction, targetUserEmail, request, &responseEnvelope)
	if err != nil {
		return nil, err
	}

	respMsg := responseEnvelope.Body.CreateItemResponse.ResponseMessages.CreateItemResponseMessage
	if respMsg.ResponseClass != "Success" {
		return nil, fmt.Errorf("EWS error creating event: %s. Code: %s", respMsg.ResponseClass, respMsg.ResponseCode)
	}

	if responseEnvelope.Body.CreateItemResponse.ResponseMessages.CreateItemResponseMessage.ResponseClass != "Success" {
		return nil, fmt.Errorf("EWS CreateItem failed: ResponseClass=%s, ResponseCode=%s", responseEnvelope.Body.CreateItemResponse.ResponseMessages.CreateItemResponseMessage.ResponseClass, responseEnvelope.Body.CreateItemResponse.ResponseMessages.CreateItemResponseMessage.ResponseCode)
	}

	if responseEnvelope.Body.CreateItemResponse.ResponseMessages.CreateItemResponseMessage.Items.CalendarItem == nil || len(responseEnvelope.Body.CreateItemResponse.ResponseMessages.CreateItemResponseMessage.Items.CalendarItem) == 0 {
		return nil, fmt.Errorf("created item ID not found in EWS response")
	}

	if len(responseEnvelope.Body.CreateItemResponse.ResponseMessages.CreateItemResponseMessage.Items.CalendarItem) > 0 {
		return &responseEnvelope.Body.CreateItemResponse.ResponseMessages.CreateItemResponseMessage.Items.CalendarItem[0].ItemId, nil
	}
	return nil, fmt.Errorf("created item ID not found in EWS response")
}

// UpdateCalendarEvent updates an existing calendar event for the target user.
// conflictResolution can be "NeverOverwrite", "AutoResolve", "AlwaysOverwrite".
// sendMeetingInvitationsOrCancellations can be "SendToNone", "SendOnlyToChanged", "SendOnlyToAll", "SendToAllAndSaveCopy", "SendToChangedAndSaveCopy".
func (c *ImpersonationClient) UpdateCalendarEvent(ctx context.Context, itemId string, changeKey string, updates EventUpdates, conflictResolution, sendMeetingInvitationsOrCancellations, targetUserEmail string) error {
	var itemChanges []SetItemField

	if updates.Subject != nil {
		itemChanges = append(itemChanges, SetItemField{
			FieldURI:     FieldURI{FieldURI: "item:Subject"},
			CalendarItem: UpdateCalendarItem{Subject: updates.Subject},
		})
	}
	if updates.Body != nil {
		itemChanges = append(itemChanges, SetItemField{
			FieldURI:     FieldURI{FieldURI: "item:Body"},
			CalendarItem: UpdateCalendarItem{Body: &ItemBody{BodyType: "Text", Content: *updates.Body}},
		})
	}
	if updates.Start != nil {
		startStr := c.FormatDateWithTZ(*updates.Start)
		itemChanges = append(itemChanges, SetItemField{
			FieldURI:     FieldURI{FieldURI: "calendar:Start"},
			CalendarItem: UpdateCalendarItem{Start: &startStr},
		})
	}
	if updates.End != nil {
		endStr := c.FormatDateWithTZ(*updates.End)
		itemChanges = append(itemChanges, SetItemField{
			FieldURI:     FieldURI{FieldURI: "calendar:End"},
			CalendarItem: UpdateCalendarItem{End: &endStr},
		})
	}
	if updates.Location != nil {
		itemChanges = append(itemChanges, SetItemField{
			FieldURI:     FieldURI{FieldURI: "calendar:Location"},
			CalendarItem: UpdateCalendarItem{Location: updates.Location},
		})
	}
	if updates.LegacyFreeBusy != nil {
		itemChanges = append(itemChanges, SetItemField{
			FieldURI:     FieldURI{FieldURI: "calendar:LegacyFreeBusyStatus"},
			CalendarItem: UpdateCalendarItem{LegacyFreeBusy: updates.LegacyFreeBusy},
		})
	}

	if len(updates.RequiredAttendees) > 0 {
		ra := &RequiredAttendees{}
		for _, attendee := range updates.RequiredAttendees {
			ra.Attendees = append(ra.Attendees, AttendeeType{Mailbox: EmailAddress{Name: attendee.Name, EmailAddress: attendee.Email, RoutingType: "SMTP"}})
		}
		itemChanges = append(itemChanges, SetItemField{
			FieldURI:     FieldURI{FieldURI: "calendar:RequiredAttendees"},
			CalendarItem: UpdateCalendarItem{RequiredAttendees: ra},
		})
	}
	if len(updates.OptionalAttendees) > 0 {
		oa := &OptionalAttendees{}
		for _, attendee := range updates.OptionalAttendees {
			oa.Attendees = append(oa.Attendees, AttendeeType{Mailbox: EmailAddress{Name: attendee.Name, EmailAddress: attendee.Email, RoutingType: "SMTP"}})
		}
		itemChanges = append(itemChanges, SetItemField{
			FieldURI:     FieldURI{FieldURI: "calendar:OptionalAttendees"},
			CalendarItem: UpdateCalendarItem{OptionalAttendees: oa},
		})
	}

	if len(itemChanges) == 0 {
		return fmt.Errorf("no updates provided for calendar event")
	}

	request := &UpdateItemRequest{
		XMLNSm:                 "http://schemas.microsoft.com/exchange/services/2006/messages",
		ConflictResolution:     conflictResolution,
		SendMeetingInvitations: sendMeetingInvitationsOrCancellations,
		MessageDisposition:     "SaveOnly",
		ItemChanges: ItemChanges{
			ItemChange: ItemChange{
				ItemId:  ItemId{Id: itemId, ChangeKey: changeKey},
				Updates: Updates{SetItemField: itemChanges},
			},
		},
	}

	var responseEnvelope UpdateItemResponseEnvelope
	soapAction := "http://schemas.microsoft.com/exchange/services/2006/messages/UpdateItem"

	err := c.doRequest(ctx, soapAction, targetUserEmail, request, &responseEnvelope)
	if err != nil {
		return err
	}

	respMsg := responseEnvelope.Body.UpdateItemResponse.ResponseMessages.UpdateItemResponseMessage
	if respMsg.ResponseClass != "Success" {
		return fmt.Errorf("EWS error updating event: %s. Code: %s", respMsg.ResponseClass, respMsg.ResponseCode)
	}

	return nil
}

// DeleteCalendarEvent deletes a calendar event for the target user.
// deleteType can be "HardDelete", "SoftDelete", "MoveToDeletedItems".
// sendMeetingCancellations can be "SendToNone", "SendOnlyToAll", "SendToAllAndSaveCopy".
func (c *ImpersonationClient) DeleteCalendarEvent(ctx context.Context, itemId string, changeKey string, deleteType, sendMeetingCancellations, targetUserEmail string) error {
	request := &DeleteItemRequest{
		XMLNSm:                   "http://schemas.microsoft.com/exchange/services/2006/messages",
		DeleteType:               deleteType,
		SendMeetingCancellations: sendMeetingCancellations,
		ItemIds: DeleteItemIds{
			ItemId: []ItemId{{Id: itemId, ChangeKey: changeKey}},
		},
	}

	var responseEnvelope DeleteItemResponseEnvelope // Make sure this type is defined in types.go
	soapAction := "http://schemas.microsoft.com/exchange/services/2006/messages/DeleteItem"

	err := c.doRequest(ctx, soapAction, targetUserEmail, request, &responseEnvelope)
	if err != nil {
		return err
	}

	respMsg := responseEnvelope.Body.DeleteItemResponse.ResponseMessages.DeleteItemResponseMessage
	if respMsg.ResponseClass != "Success" {
		return fmt.Errorf("EWS error deleting event: %s. Code: %s", respMsg.ResponseClass, respMsg.ResponseCode)
	}

	return nil
}
