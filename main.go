package main

import (
	"context"
	"fmt"
	"log"
	"os"
	"time"

	"github.com/joho/godotenv"
	"github.com/slav123/ews-workmail/ews"                             // Original EWS client
	impersonation "github.com/slav123/ews-workmail/ews-impersonation" // New Impersonation client
)

func main() {
	err := godotenv.Load()
	if err != nil {
		log.Println("No .env file found or error loading it, relying on environment variables set externally.")
	}

	// --- Configuration for Basic EWS Authentication ---
	ewsURL := os.Getenv("EWS_URL")
	ewsUsername := os.Getenv("EWS_USERNAME")
	ewsPassword := os.Getenv("EWS_PASSWORD")

	// --- Configuration for EWS Impersonation ---
	awsRegion := os.Getenv("AWS_REGION")
	workmailOrgID := os.Getenv("WORKMAIL_ORG_ID")
	impersonationRoleID := os.Getenv("IMPERSONATION_ROLE_ID")     // Using user-provided name
	impersonatedUserEmail := os.Getenv("IMPERSONATED_USER_EMAIL") // Using user-provided name
	// EWS_URL from above is also used as the endpoint for impersonation client

	// AWS Credentials will be picked up by the SDK's default credential chain
	// (e.g., from AWS_ACCESS_KEY_ID, AWS_SECRET_ACCESS_KEY, AWS_SESSION_TOKEN env vars)

	ctx := context.Background()
	startDate := time.Now().AddDate(0, 0, -2) // Last 2 days
	endDate := time.Now().AddDate(0, 0, 2)    // Next 2 days

	fmt.Println("=============================================")
	fmt.Println("TESTING BASIC EWS AUTHENTICATION (ews package)")
	fmt.Println("=============================================")
	if ewsURL == "" || ewsUsername == "" || ewsPassword == "" {
		log.Println("Skipping Basic EWS Auth test: EWS_URL, EWS_USERNAME, or EWS_PASSWORD environment variables not set.")
	} else {
		basicClient := ews.NewClient(ewsURL, ewsUsername, ewsPassword)
		// basicClient.SetTimezone("Australia/Sydney") // Optional: Set timezone if needed

		fmt.Printf("Attempting to get calendar items for user: %s (via basic auth)\n", ewsUsername)
		basicItems, err := basicClient.GetCalendarItems(startDate, endDate)
		if err != nil {
			log.Printf("Error getting calendar items (basic auth): %v\n", err)
		} else {
			if len(basicItems) == 0 {
				fmt.Println("No calendar items found (basic auth).")
			} else {
				fmt.Printf("Found %d calendar items (basic auth):\n", len(basicItems))
				for _, item := range basicItems {
					startTime, _ := basicClient.ParseDateTime(item.Start)
					endTime, _ := basicClient.ParseDateTime(item.End)
					fmt.Printf("  Subject: %s, Start: %s, End: %s, ID: %s\n",
						item.Subject,
						startTime.Format(time.RFC1123),
						endTime.Format(time.RFC1123),
						item.ItemId.Id,
					)
				}
			}
		}
	}

	fmt.Println("\n======================================================")
	fmt.Println("TESTING EWS IMPERSONATION (ews-impersonation package)")
	fmt.Println("======================================================")
	if ewsURL == "" || awsRegion == "" || workmailOrgID == "" || impersonationRoleID == "" || impersonatedUserEmail == "" {
		log.Println("Skipping EWS Impersonation test: EWS_URL, AWS_REGION, WORKMAIL_ORG_ID, IMPERSONATION_ROLE_ID, or IMPERSONATED_USER_EMAIL environment variables not set.")
	} else {
		impersonationClient, err := impersonation.NewImpersonationClient(ctx, awsRegion, workmailOrgID, impersonationRoleID, ewsURL) // Using EWS_URL as endpoint
		if err != nil {
			log.Printf("Error creating impersonation client: %v\n", err)
		} else {
			// impersonationClient.SetTimezone("Australia/Sydney") // Optional: Set timezone if needed

			fmt.Printf("Attempting to get calendar items for user: %s (via impersonation for %s)\n", ewsUsername, impersonatedUserEmail)
			impersonatedItems, err := impersonationClient.GetCalendarItems(ctx, startDate, endDate, impersonatedUserEmail)
			if err != nil {
				log.Printf("Error getting calendar items (impersonation): %v\n", err)
			} else {
				if len(impersonatedItems) == 0 {
					fmt.Println("No calendar items found (impersonation).")
				} else {
					fmt.Printf("Found %d calendar items (impersonation for %s):\n", len(impersonatedItems), impersonatedUserEmail)
					for _, item := range impersonatedItems {
						startTime, _ := impersonationClient.ParseDateTime(item.Start)
						endTime, _ := impersonationClient.ParseDateTime(item.End)
						fmt.Printf("  Subject: %s, Start: %s, End: %s, ID: %s\n",
							item.Subject,
							startTime.Format(time.RFC1123),
							endTime.Format(time.RFC1123),
							item.ItemId.Id,
						)
					}
				}
			}
			// --- Test CreateCalendarEvent (Impersonation) ---
				fmt.Println("\nAttempting to create a calendar event (impersonation)...")
				newEvent := impersonation.CalendarEvent{
					Subject:  "Impersonated Test Event - DO NOT DELETE MANUALLY",
					Body:     "This is a test event created via impersonation.",
					Start:    time.Now().Add(24 * time.Hour), // Tomorrow
					End:      time.Now().Add(25 * time.Hour), // Tomorrow + 1 hour
					Location: "Virtual Test Room",
				}
				createdEventItem, err := impersonationClient.CreateCalendarEvent(ctx, newEvent, "SendToNone", impersonatedUserEmail)
				if err != nil {
					log.Printf("Error creating calendar event (impersonation): %v\n", err)
				} else {
					fmt.Printf("Successfully created event (impersonation), ID: %s, ChangeKey: %s\n", createdEventItem.Id, createdEventItem.ChangeKey)

					// --- Test UpdateCalendarEvent (Impersonation) ---
					fmt.Println("\nAttempting to update the calendar event (impersonation) - Subject ONLY...")
					updatedSubject := "Impersonated Test Event - Subject UPDATED ONLY"
					// updatedBody := "The body of this test event has been updated via impersonation."
					// updatedLocation := "Updated Virtual Test Room"
					updates := impersonation.EventUpdates{
						Subject:  &updatedSubject,
						// Body:     &updatedBody,
						// Location: &updatedLocation,
						// You can also update Start, End, Attendees, etc.
					}
					err = impersonationClient.UpdateCalendarEvent(ctx, createdEventItem.Id, createdEventItem.ChangeKey, updates, "AutoResolve", "SendToNone", impersonatedUserEmail)
					if err != nil {
						log.Printf("Error updating calendar event (impersonation): %v\n", err)
					} else {
						fmt.Printf("Successfully updated event (impersonation), ID: %s\n", createdEventItem.Id)
					}

					// --- Test DeleteCalendarEvent (Impersonation) ---
					fmt.Println("\nAttempting to delete the calendar event (impersonation)...")
					err = impersonationClient.DeleteCalendarEvent(ctx, createdEventItem.Id, createdEventItem.ChangeKey, "MoveToDeletedItems", "SendToNone", impersonatedUserEmail)
					if err != nil {
						log.Printf("Error deleting calendar event (impersonation): %v\n", err)
					} else {
						fmt.Printf("Successfully deleted event (impersonation), ID: %s\n", createdEventItem.Id)
					}
				}
		}
	}
	fmt.Println("\n=============================================")
	fmt.Println("TESTING COMPLETE")
	fmt.Println("=============================================")
}
