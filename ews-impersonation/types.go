package ewsimpersonation

import (
	"encoding/xml"
	"time"
)

// LegacyFreeBusyStatus represents the free/busy status of a calendar item
type LegacyFreeBusyStatus string

// LegacyFreeBusyStatus constants
const (
	FreeBusyFree      LegacyFreeBusyStatus = "Free"      // Time slot is available
	FreeBusyTentative LegacyFreeBusyStatus = "Tentative" // Time slot is tentatively booked
	FreeBusyBusy      LegacyFreeBusyStatus = "Busy"      // Time slot is busy
	FreeBusyOOF       LegacyFreeBusyStatus = "OOF"       // User is Out of Office
	FreeBusyNoData    LegacyFreeBusyStatus = "NoData"    // Status is unknown
)

// ExchangeImpersonationType defines the structure for the EWS impersonation header
type ExchangeImpersonationType struct {
	XMLName       xml.Name          `xml:"t:ExchangeImpersonation"`
	ConnectingSID ConnectingSIDType `xml:"t:ConnectingSID"`
}

// ConnectingSIDType specifies the user to impersonate
type ConnectingSIDType struct {
	PrimarySmtpAddress string `xml:"t:PrimarySmtpAddress"`
}

// SOAP envelope structures
type Envelope struct {
	XMLName xml.Name `xml:"s:Envelope"`
	XMLNS   string   `xml:"xmlns:s,attr"`
	XMLNSt  string   `xml:"xmlns:t,attr"`
	XMLNSm  string   `xml:"xmlns:m,attr"`
	Header  Header   `xml:"s:Header"`
	Body    Body     `xml:"s:Body"`
}

// Response structures
type ResponseEnvelope struct {
	XMLName xml.Name     `xml:"Envelope"`
	Body    ResponseBody `xml:"Body"`
}

type ResponseBody struct {
	FindItemResponse FindItemResponse `xml:"FindItemResponse"`
}

type FindItemResponse struct {
	ResponseMessages ResponseMessages `xml:"ResponseMessages"`
}

type ResponseMessages struct {
	FindItemResponseMessage FindItemResponseMessage `xml:"FindItemResponseMessage"`
}

type FindItemResponseMessage struct {
	ResponseCode string     `xml:"ResponseCode"`
	RootFolder   RootFolder `xml:"RootFolder"`
}

type RootFolder struct {
	Items      Items `xml:"Items"`
	TotalItems int   `xml:"TotalItemsInView,attr"`
}

type Items struct {
	CalendarItem []CalendarItem `xml:"CalendarItem"`
}

type CalendarItem struct {
	ItemId         ItemId `xml:"ItemId"`
	Subject        string `xml:"Subject"`
	Start          string `xml:"Start"`
	End            string `xml:"End"`
	Location       string `xml:"Location"`
	IsAllDayEvent  bool   `xml:"IsAllDayEvent,omitempty"`
	LegacyFreeBusy LegacyFreeBusyStatus `xml:"LegacyFreeBusyStatus,omitempty"`
	Organizer      struct {
		Mailbox struct {
			Name         string `xml:"Name"`
			EmailAddress string `xml:"EmailAddress"`
			RoutingType  string `xml:"RoutingType"`
		} `xml:"Mailbox"`
	} `xml:"Organizer"`
}

type ItemId struct {
	Id        string `xml:"Id,attr"`
	ChangeKey string `xml:"ChangeKey,attr,omitempty"`
}

// Header now includes ExchangeImpersonation
type Header struct {
	ServerVersionInfo     ServerVersionInfo          `xml:"t:RequestServerVersion"`
	ExchangeImpersonation *ExchangeImpersonationType `xml:"t:ExchangeImpersonation,omitempty"`
}

type ServerVersionInfo struct {
	Version string `xml:"Version,attr"`
}

type Body struct {
	FindItem   *FindItemRequest    `xml:"m:FindItem,omitempty"`
	CreateItem *CreateEventRequest `xml:"m:CreateItem,omitempty"`
	DeleteItem *DeleteItemRequest  `xml:"m:DeleteItem,omitempty"`
	UpdateItem *UpdateItemRequest  `xml:"m:UpdateItem,omitempty"`
}

type FindItemRequest struct {
	XMLName         xml.Name        `xml:"m:FindItem"`
	XMLNSm          string          `xml:"xmlns:m,attr"`
	Traversal       string          `xml:"Traversal,attr"`
	ItemShape       ItemShape       `xml:"m:ItemShape"`
	CalendarView    CalendarView    `xml:"m:CalendarView"`
	ParentFolderIds ParentFolderIds `xml:"m:ParentFolderIds"`
}

type ItemShape struct {
	BaseShape string `xml:"t:BaseShape"`
}

type CalendarView struct {
	StartDate string `xml:"StartDate,attr"`
	EndDate   string `xml:"EndDate,attr"`
}

type ParentFolderIds struct {
	DistinguishedFolderId DistinguishedFolderId `xml:"t:DistinguishedFolderId"`
}

type DistinguishedFolderId struct {
	Id string `xml:"Id,attr"`
}

type CreateEventRequest struct {
	XMLName                xml.Name          `xml:"m:CreateItem"`
	SendMeetingInvitations string            `xml:"SendMeetingInvitations,attr"`
	SavedItemFolderId      SavedItemFolderId `xml:"m:SavedItemFolderId"`
	Items                  CreateEventItems  `xml:"m:Items"`
}

type CreateEventItems struct {
	CalendarItem CreateEventCalendarItem `xml:"t:CalendarItem"`
}

type CreateEventCalendarItem struct {
	XMLNSt            string             `xml:"xmlns,attr"`
	Subject           string             `xml:"t:Subject"`
	Body              ItemBody           `xml:"t:Body"`
	ReminderIsSet     bool               `xml:"t:ReminderIsSet"`
	ReminderMinutes   int                `xml:"t:ReminderMinutesBeforeStart"`
	Start             string             `xml:"t:Start"`
	End               string             `xml:"t:End"`
	IsAllDayEvent     bool               `xml:"t:IsAllDayEvent"`
	LegacyFreeBusy    LegacyFreeBusyStatus `xml:"t:LegacyFreeBusyStatus"`
	Location          string             `xml:"t:Location,omitempty"`
	RequiredAttendees *RequiredAttendees `xml:"t:RequiredAttendees,omitempty"`
	OptionalAttendees *OptionalAttendees `xml:"t:OptionalAttendees,omitempty"`
}

type ItemBody struct {
	BodyType string `xml:"BodyType,attr"`
	Content  string `xml:",chardata"`
}

type SavedItemFolderId struct {
	DistinguishedFolderId DistinguishedFolderId `xml:"t:DistinguishedFolderId"`
}

type DeleteItemRequest struct {
	XMLName                  xml.Name      `xml:"m:DeleteItem"`
	XMLNSm                   string        `xml:"xmlns:m,attr"`
	DeleteType               string        `xml:"DeleteType,attr"`
	SendMeetingCancellations string        `xml:"SendMeetingCancellations,attr"`
	ItemIds                  DeleteItemIds `xml:"m:ItemIds"`
}

type DeleteItemIds struct {
	ItemId []ItemId `xml:"t:ItemId"`
}

// CreateItem response structures
type CreateItemResponseEnvelope struct {
	XMLName xml.Name           `xml:"Envelope"`
	Body    CreateItemResponse `xml:"Body"`
}

type CreateItemResponse struct {
	CreateItemResponse CreateItemResponseMessage `xml:"CreateItemResponse"`
}

type CreateItemResponseMessage struct {
	ResponseMessages ResponseMessagesCreate `xml:"ResponseMessages"`
}

type ResponseMessagesCreate struct {
	CreateItemResponseMessage CreateItemResponseMessageType `xml:"CreateItemResponseMessage"`
}

type CreateItemResponseMessageType struct {
	ResponseClass string     `xml:"ResponseClass,attr"`
	ResponseCode  string     `xml:"ResponseCode"`
	Items         ItemsArray `xml:"Items"`
}

type ItemsArray struct {
	CalendarItem []struct {
		ItemId ItemId `xml:"ItemId"`
	} `xml:"CalendarItem"`
}

type UpdateItemRequest struct {
	XMLName                xml.Name    `xml:"m:UpdateItem"`
	XMLNSm                 string      `xml:"xmlns:m,attr"`
	ConflictResolution     string      `xml:"ConflictResolution,attr"`
	SendMeetingInvitations string      `xml:"SendMeetingInvitationsOrCancellations,attr"`
	MessageDisposition     string      `xml:"MessageDisposition,attr,omitempty"`
	ItemChanges            ItemChanges `xml:"m:ItemChanges"`
}

type ItemChanges struct {
	ItemChange ItemChange `xml:"t:ItemChange"`
}
type ItemChange struct {
	ItemId  ItemId  `xml:"t:ItemId"`
	Updates Updates `xml:"t:Updates"`
}
type Updates struct {
	SetItemField []SetItemField `xml:"t:SetItemField"`
}
type SetItemField struct {
	FieldURI     FieldURI           `xml:"t:FieldURI"`
	CalendarItem UpdateCalendarItem `xml:"t:CalendarItem"`
}

type FieldURI struct {
	FieldURI string `xml:"FieldURI,attr"`
}

type UpdateCalendarItem struct {
	Start             *string            `xml:"t:Start,omitempty"`
	End               *string            `xml:"t:End,omitempty"`
	Subject           *string            `xml:"t:Subject,omitempty"`
	Body              *ItemBody          `xml:"t:Body,omitempty"`
	LegacyFreeBusy    *LegacyFreeBusyStatus `xml:"t:LegacyFreeBusyStatus,omitempty"`
	Location          *string            `xml:"t:Location,omitempty"`
	RequiredAttendees *RequiredAttendees `xml:"t:RequiredAttendees,omitempty"`
	OptionalAttendees *OptionalAttendees `xml:"t:OptionalAttendees,omitempty"`
}

type RequiredAttendees struct {
	Attendees []AttendeeType `xml:"t:Attendee"`
}

type AttendeeType struct {
	Mailbox EmailAddress `xml:"t:Mailbox"`
}

type EmailAddress struct {
	Name         string `xml:"t:Name"`
	EmailAddress string `xml:"t:EmailAddress"`
	RoutingType  string `xml:"t:RoutingType"`
}

type OptionalAttendees struct {
	Attendees []AttendeeType `xml:"t:Attendee"`
}

// Attendee represents a meeting attendee for client-side operations
type Attendee struct {
	Name  string
	Email string
}

// UpdateItem response structures
type UpdateItemResponseEnvelope struct {
	XMLName xml.Name               `xml:"Envelope"`
	Body    UpdateItemResponseBody `xml:"Body"`
}

type UpdateItemResponseBody struct {
	UpdateItemResponse UpdateItemResponseMessage `xml:"UpdateItemResponse"`
}

type UpdateItemResponseMessage struct {
	ResponseMessages UpdateResponseMessages `xml:"ResponseMessages"`
}

type UpdateResponseMessages struct {
	UpdateItemResponseMessage UpdateItemResponseMessageType `xml:"UpdateItemResponseMessage"`
}

type UpdateItemResponseMessageType struct {
	ResponseClass string `xml:"ResponseClass,attr"`
	ResponseCode  string `xml:"ResponseCode"`
}

// CalendarEvent represents a calendar event to be created (client-side struct)
type CalendarEvent struct {
	Subject           string
	Body              string
	Start             time.Time
	End               time.Time
	Location          string
	IsAllDay          bool
	RequiredAttendees []Attendee
	OptionalAttendees []Attendee
	SendInvites       bool
}

// EventUpdates represents updates to an existing calendar event (client-side struct)
type EventUpdates struct {
	Start             *time.Time
	End               *time.Time
	Subject           *string
	Body              *string // Assumes ItemBody with BodyType="Text"
	LegacyFreeBusy    *LegacyFreeBusyStatus
	Location          *string
	RequiredAttendees []Attendee
	OptionalAttendees []Attendee
}

// DeleteItem response structures
type DeleteItemResponseEnvelope struct {
	XMLName xml.Name               `xml:"Envelope"`
	Body    DeleteItemResponseBody `xml:"Body"`
}

type DeleteItemResponseBody struct {
	DeleteItemResponse DeleteItemResponseMessage `xml:"DeleteItemResponse"`
}

type DeleteItemResponseMessage struct {
	ResponseMessages DeleteItemResponseMessages `xml:"ResponseMessages"`
}

type DeleteItemResponseMessages struct {
	DeleteItemResponseMessage DeleteItemResponseMessageType `xml:"DeleteItemResponseMessage"`
}

type DeleteItemResponseMessageType struct {
	ResponseClass string `xml:"ResponseClass,attr"`
	ResponseCode  string `xml:"ResponseCode"`
}
