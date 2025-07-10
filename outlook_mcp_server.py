import datetime
import os
import win32com.client
from typing import List, Optional, Dict, Any
from mcp.server.fastmcp import FastMCP, Context

# Initialize FastMCP server
mcp = FastMCP("outlook-assistant")

# Constants
MAX_DAYS = 30
# Email cache for storing retrieved emails by number
email_cache = {}
# Calendar cache for storing retrieved appointments by number
calendar_cache = {}

# Helper functions
def connect_to_outlook():
    """Connect to Outlook application using COM"""
    try:
        outlook = win32com.client.Dispatch("Outlook.Application")
        namespace = outlook.GetNamespace("MAPI")
        return outlook, namespace
    except Exception as e:
        raise Exception(f"Failed to connect to Outlook: {str(e)}")

def get_folder_by_name(namespace, folder_name: str):
    """Get a specific Outlook folder by name"""
    try:
        # First check inbox subfolder
        inbox = namespace.GetDefaultFolder(6)  # 6 is the index for inbox folder
        
        # Check inbox subfolders first (most common)
        for folder in inbox.Folders:
            if folder.Name.lower() == folder_name.lower():
                return folder
                
        # Then check all folders at root level
        for folder in namespace.Folders:
            if folder.Name.lower() == folder_name.lower():
                return folder
            
            # Also check subfolders
            for subfolder in folder.Folders:
                if subfolder.Name.lower() == folder_name.lower():
                    return subfolder
                    
        # If not found
        return None
    except Exception as e:
        raise Exception(f"Failed to access folder {folder_name}: {str(e)}")

def format_email(mail_item) -> Dict[str, Any]:
    """Format an Outlook mail item into a structured dictionary"""
    try:
        # Extract recipients
        recipients = []
        if mail_item.Recipients:
            for i in range(1, mail_item.Recipients.Count + 1):
                recipient = mail_item.Recipients(i)
                try:
                    recipients.append(f"{recipient.Name} <{recipient.Address}>")
                except:
                    recipients.append(f"{recipient.Name}")
        
        # Format the email data
        email_data = {
            "id": mail_item.EntryID,
            "conversation_id": mail_item.ConversationID if hasattr(mail_item, 'ConversationID') else None,
            "subject": mail_item.Subject,
            "sender": mail_item.SenderName,
            "sender_email": mail_item.SenderEmailAddress,
            "received_time": mail_item.ReceivedTime.strftime("%Y-%m-%d %H:%M:%S") if mail_item.ReceivedTime else None,
            "recipients": recipients,
            "body": mail_item.Body,
            "has_attachments": mail_item.Attachments.Count > 0,
            "attachment_count": mail_item.Attachments.Count if hasattr(mail_item, 'Attachments') else 0,
            "unread": mail_item.UnRead if hasattr(mail_item, 'UnRead') else False,
            "importance": mail_item.Importance if hasattr(mail_item, 'Importance') else 1,
            "categories": mail_item.Categories if hasattr(mail_item, 'Categories') else ""
        }
        return email_data
    except Exception as e:
        raise Exception(f"Failed to format email: {str(e)}")

def clear_email_cache():
    """Clear the email cache"""
    global email_cache
    email_cache = {}

def clear_calendar_cache():
    """Clear the calendar cache"""
    global calendar_cache
    calendar_cache = {}

def format_appointment(appointment) -> Dict[str, Any]:
    """Format an Outlook appointment item into a structured dictionary"""
    try:
        # Extract attendees
        attendees = []
        if hasattr(appointment, 'Recipients') and appointment.Recipients:
            for i in range(1, appointment.Recipients.Count + 1):
                recipient = appointment.Recipients(i)
                try:
                    attendees.append(f"{recipient.Name} <{recipient.Address}>")
                except:
                    attendees.append(f"{recipient.Name}")
        
        # Format the appointment data
        appointment_data = {
            "id": appointment.EntryID,
            "subject": appointment.Subject,
            "start_time": appointment.Start.strftime("%Y-%m-%d %H:%M:%S") if appointment.Start else None,
            "end_time": appointment.End.strftime("%Y-%m-%d %H:%M:%S") if appointment.End else None,
            "location": appointment.Location if hasattr(appointment, 'Location') else "",
            "organizer": appointment.Organizer if hasattr(appointment, 'Organizer') else "",
            "attendees": attendees,
            "body": appointment.Body if hasattr(appointment, 'Body') else "",
            "is_all_day": appointment.AllDayEvent if hasattr(appointment, 'AllDayEvent') else False,
            "is_recurring": appointment.IsRecurring if hasattr(appointment, 'IsRecurring') else False,
            "reminder_minutes": appointment.ReminderMinutesBeforeStart if hasattr(appointment, 'ReminderMinutesBeforeStart') else 0,
            "categories": appointment.Categories if hasattr(appointment, 'Categories') else "",
            "importance": appointment.Importance if hasattr(appointment, 'Importance') else 1,
            "busy_status": appointment.BusyStatus if hasattr(appointment, 'BusyStatus') else 2
        }
        return appointment_data
    except Exception as e:
        raise Exception(f"Failed to format appointment: {str(e)}")

def get_appointments_from_calendar(calendar, days: int, search_term: Optional[str] = None):
    """Get appointments from calendar with optional search filter"""
    appointments_list = []
    
    # Calculate the date range
    now = datetime.datetime.now()
    # For calendar appointments, look from today to future days
    start_date = now.replace(hour=0, minute=0, second=0, microsecond=0)  # Start from today 00:00:00
    end_date = (start_date + datetime.timedelta(days=days)).replace(hour=23, minute=59, second=59, microsecond=0)  # Look forward exactly 'days', end at 23:59:59    
    try:
        # Get appointments in date range
        calendar_items = calendar.Items
        calendar_items.Sort("[Start]", False)  # Sort by start time
        
        # Process appointments without date filter first, then filter manually
        for item in calendar_items:
            try:
                # Check if appointment is within our date range
                if hasattr(item, 'Start') and item.Start:
                    item_start = item.Start.replace(tzinfo=None)
                    if not (start_date <= item_start <= end_date):
                        continue
                
                # Manual search filter if needed
                if search_term:
                    search_terms = [term.strip().lower() for term in search_term.split(" OR ")]
                    
                    # Check if any of the search terms match
                    found_match = False
                    for term in search_terms:
                        if (term in item.Subject.lower() or 
                            term in item.Location.lower() or 
                            term in item.Body.lower()):
                            found_match = True
                            break
                    
                    if not found_match:
                        continue
                
                # Format and add the appointment
                appointment_data = format_appointment(item)
                appointments_list.append(appointment_data)
                
            except Exception as e:
                print(f"Warning: Error processing appointment: {str(e)}")
                continue
                
    except Exception as e:
        print(f"Error retrieving appointments: {str(e)}")
        
    return appointments_list

def get_emails_from_folder(folder, days: int, search_term: Optional[str] = None):
    """Get emails from a folder with optional search filter"""
    emails_list = []
    
    # Calculate the date threshold
    now = datetime.datetime.now()
    threshold_date = now - datetime.timedelta(days=days)
    
    try:
        # Set up filtering
        folder_items = folder.Items
        folder_items.Sort("[ReceivedTime]", True)  # Sort by received time, newest first
        
        # If we have a search term, apply it
        if search_term:
            # Handle OR operators in search term
            search_terms = [term.strip() for term in search_term.split(" OR ")]
            
            # Try to create a filter for subject, sender name or body
            try:
                # Build SQL filter with OR conditions for each search term
                sql_conditions = []
                for term in search_terms:
                    sql_conditions.append(f"\"urn:schemas:httpmail:subject\" LIKE '%{term}%'")
                    sql_conditions.append(f"\"urn:schemas:httpmail:fromname\" LIKE '%{term}%'")
                    sql_conditions.append(f"\"urn:schemas:httpmail:textdescription\" LIKE '%{term}%'")
                
                filter_term = f"@SQL=" + " OR ".join(sql_conditions)
                folder_items = folder_items.Restrict(filter_term)
            except:
                # If filtering fails, we'll do manual filtering later
                pass
        
        # Process emails
        count = 0
        for item in folder_items:
            try:
                if hasattr(item, 'ReceivedTime') and item.ReceivedTime:
                    # Convert to naive datetime for comparison
                    received_time = item.ReceivedTime.replace(tzinfo=None)
                    
                    # Skip emails older than our threshold
                    if received_time < threshold_date:
                        continue
                    
                    # Manual search filter if needed
                    if search_term and folder_items == folder.Items:  # If we didn't apply filter earlier
                        # Handle OR operators in search term for manual filtering
                        search_terms = [term.strip().lower() for term in search_term.split(" OR ")]
                        
                        # Check if any of the search terms match
                        found_match = False
                        for term in search_terms:
                            if (term in item.Subject.lower() or 
                                term in item.SenderName.lower() or 
                                term in item.Body.lower()):
                                found_match = True
                                break
                        
                        if not found_match:
                            continue
                    
                    # Format and add the email
                    email_data = format_email(item)
                    emails_list.append(email_data)
                    count += 1
            except Exception as e:
                print(f"Warning: Error processing email: {str(e)}")
                continue
                
    except Exception as e:
        print(f"Error retrieving emails: {str(e)}")
        
    return emails_list

# MCP Tools
@mcp.tool()
def list_folders() -> str:
    """
    List all available mail folders in Outlook
    
    Returns:
        A list of available mail folders
    """
    try:
        # Connect to Outlook
        _, namespace = connect_to_outlook()
        
        result = "Available mail folders:\n\n"
        
        # List all root folders and their subfolders
        for folder in namespace.Folders:
            result += f"- {folder.Name}\n"
            
            # List subfolders
            for subfolder in folder.Folders:
                result += f"  - {subfolder.Name}\n"
                
                # List subfolders (one more level)
                try:
                    for subsubfolder in subfolder.Folders:
                        result += f"    - {subsubfolder.Name}\n"
                except:
                    pass
        
        return result
    except Exception as e:
        return f"Error listing mail folders: {str(e)}"

@mcp.tool()
def list_recent_emails(days: int = 7, folder_name: Optional[str] = None) -> str:
    """
    List email titles from the specified number of days
    
    Args:
        days: Number of days to look back for emails (max 30)
        folder_name: Name of the folder to check (if not specified, checks the Inbox)
        
    Returns:
        Numbered list of email titles with sender information
    """
    if not isinstance(days, int) or days < 1 or days > MAX_DAYS:
        return f"Error: 'days' must be an integer between 1 and {MAX_DAYS}"
    
    try:
        # Connect to Outlook
        _, namespace = connect_to_outlook()
        
        # Get the appropriate folder
        if folder_name:
            folder = get_folder_by_name(namespace, folder_name)
            if not folder:
                return f"Error: Folder '{folder_name}' not found"
        else:
            folder = namespace.GetDefaultFolder(6)  # Default inbox
        
        # Clear previous cache
        clear_email_cache()
        
        # Get emails from folder
        emails = get_emails_from_folder(folder, days)
        
        # Format the output and cache emails
        folder_display = f"'{folder_name}'" if folder_name else "Inbox"
        if not emails:
            return f"No emails found in {folder_display} from the last {days} days."
        
        result = f"Found {len(emails)} emails in {folder_display} from the last {days} days:\n\n"
        
        # Cache emails and build result
        for i, email in enumerate(emails, 1):
            # Store in cache
            email_cache[i] = email
            
            # Format for display
            result += f"Email #{i}\n"
            result += f"Subject: {email['subject']}\n"
            result += f"From: {email['sender']} <{email['sender_email']}>\n"
            result += f"Received: {email['received_time']}\n"
            result += f"Read Status: {'Read' if not email['unread'] else 'Unread'}\n"
            result += f"Has Attachments: {'Yes' if email['has_attachments'] else 'No'}\n\n"
        
        result += "To view the full content of an email, use the get_email_by_number tool with the email number."
        return result
    
    except Exception as e:
        return f"Error retrieving email titles: {str(e)}"

@mcp.tool()
def search_emails(search_term: str, days: int = 7, folder_name: Optional[str] = None) -> str:
    """
    Search emails by contact name or keyword within a time period
    
    Args:
        search_term: Name or keyword to search for
        days: Number of days to look back (max 30)
        folder_name: Name of the folder to search (if not specified, searches the Inbox)
        
    Returns:
        Numbered list of matching email titles
    """
    if not search_term:
        return "Error: Please provide a search term"
        
    if not isinstance(days, int) or days < 1 or days > MAX_DAYS:
        return f"Error: 'days' must be an integer between 1 and {MAX_DAYS}"
    
    try:
        # Connect to Outlook
        _, namespace = connect_to_outlook()
        
        # Get the appropriate folder
        if folder_name:
            folder = get_folder_by_name(namespace, folder_name)
            if not folder:
                return f"Error: Folder '{folder_name}' not found"
        else:
            folder = namespace.GetDefaultFolder(6)  # Default inbox
        
        # Clear previous cache
        clear_email_cache()
        
        # Get emails matching search term
        emails = get_emails_from_folder(folder, days, search_term)
        
        # Format the output and cache emails
        folder_display = f"'{folder_name}'" if folder_name else "Inbox"
        if not emails:
            return f"No emails matching '{search_term}' found in {folder_display} from the last {days} days."
        
        result = f"Found {len(emails)} emails matching '{search_term}' in {folder_display} from the last {days} days:\n\n"
        
        # Cache emails and build result
        for i, email in enumerate(emails, 1):
            # Store in cache
            email_cache[i] = email
            
            # Format for display
            result += f"Email #{i}\n"
            result += f"Subject: {email['subject']}\n"
            result += f"From: {email['sender']} <{email['sender_email']}>\n"
            result += f"Received: {email['received_time']}\n"
            result += f"Read Status: {'Read' if not email['unread'] else 'Unread'}\n"
            result += f"Has Attachments: {'Yes' if email['has_attachments'] else 'No'}\n\n"
        
        result += "To view the full content of an email, use the get_email_by_number tool with the email number."
        return result
    
    except Exception as e:
        return f"Error searching emails: {str(e)}"

@mcp.tool()
def get_email_by_number(email_number: int) -> str:
    """
    Get detailed content of a specific email by its number from the last listing
    
    Args:
        email_number: The number of the email from the list results
        
    Returns:
        Full details of the specified email
    """
    try:
        if not email_cache:
            return "Error: No emails have been listed yet. Please use list_recent_emails or search_emails first."
        
        if email_number not in email_cache:
            return f"Error: Email #{email_number} not found in the current listing."
        
        email_data = email_cache[email_number]
        
        # Connect to Outlook to get the full email content
        _, namespace = connect_to_outlook()
        
        # Retrieve the specific email
        email = namespace.GetItemFromID(email_data["id"])
        if not email:
            return f"Error: Email #{email_number} could not be retrieved from Outlook."
        
        # Format the output
        result = f"Email #{email_number} Details:\n\n"
        result += f"Subject: {email_data['subject']}\n"
        result += f"From: {email_data['sender']} <{email_data['sender_email']}>\n"
        result += f"Received: {email_data['received_time']}\n"
        result += f"Recipients: {', '.join(email_data['recipients'])}\n"
        result += f"Has Attachments: {'Yes' if email_data['has_attachments'] else 'No'}\n"
        
        if email_data['has_attachments']:
            result += "Attachments:\n"
            for i in range(1, email.Attachments.Count + 1):
                attachment = email.Attachments(i)
                result += f"  - {attachment.FileName}\n"
        
        result += "\nBody:\n"
        result += email_data['body']
        
        result += "\n\nTo reply to this email, use the reply_to_email_by_number tool with this email number."
        
        return result
    
    except Exception as e:
        return f"Error retrieving email details: {str(e)}"

@mcp.tool()
def reply_to_email_by_number(email_number: int, reply_text: str, save_as_draft: bool = True) -> str:
    """
    Reply to a specific email by its number from the last listing
    
    Args:
        email_number: The number of the email from the list results
        reply_text: The text content for the reply
        save_as_draft: If True, save as draft instead of sending immediately (default: False)
        
    Returns:
        Status message indicating success or failure
    """
    try:
        if not email_cache:
            return "Error: No emails have been listed yet. Please use list_recent_emails or search_emails first."
        
        if email_number not in email_cache:
            return f"Error: Email #{email_number} not found in the current listing."
        
        email_id = email_cache[email_number]["id"]
        
        # Connect to Outlook
        outlook, namespace = connect_to_outlook()
        
        # Retrieve the specific email
        email = namespace.GetItemFromID(email_id)
        if not email:
            return f"Error: Email #{email_number} could not be retrieved from Outlook."
        
        # Create reply
        reply = email.Reply()
        reply.Body = reply_text
        
        # Handle string 'true'/'false' as well as boolean values
        if isinstance(save_as_draft, str):
            save_as_draft = save_as_draft.lower() in ['true', '1', 'yes']
        
        if save_as_draft:
            # Save as draft
            reply.Save()
            return f"Reply saved as draft for: {email.SenderName} <{email.SenderEmailAddress}>"
        else:
            # Send the reply
            reply.Send()
            return f"Reply sent successfully to: {email.SenderName} <{email.SenderEmailAddress}>"
    
    except Exception as e:
        return f"Error replying to email: {str(e)}"

@mcp.tool()
def compose_email(recipient_email: str, subject: str, body: str, cc_email: Optional[str] = None, save_as_draft: bool = True) -> str:
    """
    Compose and send a new email
    
    Args:
        recipient_email: Email address of the recipient
        subject: Subject line of the email
        body: Main content of the email
        cc_email: Email address for CC (optional)
        save_as_draft: If True, save as draft instead of sending immediately (default: False)
        
    Returns:
        Status message indicating success or failure
    """
    try:
        # Connect to Outlook
        outlook, _ = connect_to_outlook()
        
        # Create a new email
        mail = outlook.CreateItem(0)  # 0 is the value for a mail item
        mail.Subject = subject
        mail.To = recipient_email
        
        if cc_email:
            mail.CC = cc_email
        
        # Add signature to the body
        mail.Body = body
        
        # Handle string 'true'/'false' as well as boolean values
        if isinstance(save_as_draft, str):
            save_as_draft = save_as_draft.lower() in ['true', '1', 'yes']
        
        if save_as_draft:
            # Save as draft
            mail.Save()
            return f"Email saved as draft for: {recipient_email}"
        else:
            # Send the email
            mail.Send()
            return f"Email sent successfully to: {recipient_email}"
    
    except Exception as e:
        return f"Error composing email: {str(e)}"

# Calendar Tools
@mcp.tool()
def list_calendar_appointments(days: int = 14) -> str:
    """
    List calendar appointments from the specified number of days (past and future)
    
    Args:
        days: Number of days to look (half past, half future, max 60)
        
    Returns:
        Numbered list of appointments with basic information
    """
    if not isinstance(days, int) or days < 1 or days > 60:
        return f"Error: 'days' must be an integer between 1 and 60"
    
    try:
        # Connect to Outlook
        _, namespace = connect_to_outlook()
        
        # Get the default calendar folder
        calendar = namespace.GetDefaultFolder(9)  # 9 is the calendar folder
        
        # Clear previous cache
        clear_calendar_cache()
        
        # Get appointments from calendar
        appointments = get_appointments_from_calendar(calendar, days)
        
        if not appointments:
            return f"No appointments found in the next {days} days."
        
        result = f"Found {len(appointments)} appointments in the next {days} days:\n\n"
        
        # Cache appointments and build result
        for i, appointment in enumerate(appointments, 1):
            # Store in cache
            calendar_cache[i] = appointment
            
            # Format for display
            result += f"Appointment #{i}\n"
            result += f"Subject: {appointment['subject']}\n"
            result += f"Start: {appointment['start_time']}\n"
            result += f"End: {appointment['end_time']}\n"
            result += f"Location: {appointment['location']}\n"
            result += f"All Day: {'Yes' if appointment['is_all_day'] else 'No'}\n"
            result += f"Recurring: {'Yes' if appointment['is_recurring'] else 'No'}\n\n"
        
        result += "To view the full details of an appointment, use the get_appointment_by_number tool with the appointment number."
        return result
    
    except Exception as e:
        return f"Error retrieving calendar appointments: {str(e)}"

@mcp.tool()
def search_calendar_appointments(search_term: str, days: int = 14) -> str:
    """
    Search calendar appointments by keyword within a time period
    
    Args:
        search_term: Keyword to search for in subject, location, or body
        days: Number of days to look (half past, half future, max 60)
        
    Returns:
        Numbered list of matching appointments
    """
    if not search_term:
        return "Error: Please provide a search term"
        
    if not isinstance(days, int) or days < 1 or days > 60:
        return f"Error: 'days' must be an integer between 1 and 60"
    
    try:
        # Connect to Outlook
        _, namespace = connect_to_outlook()
        
        # Get the default calendar folder
        calendar = namespace.GetDefaultFolder(9)  # 9 is the calendar folder
        
        # Clear previous cache
        clear_calendar_cache()
        
        # Get appointments matching search term
        appointments = get_appointments_from_calendar(calendar, days, search_term)
        
        if not appointments:
            return f"No appointments matching '{search_term}' found in the next {days} days."
        
        result = f"Found {len(appointments)} appointments matching '{search_term}' in the next {days} days:\n\n"
        
        # Cache appointments and build result
        for i, appointment in enumerate(appointments, 1):
            # Store in cache
            calendar_cache[i] = appointment
            
            # Format for display
            result += f"Appointment #{i}\n"
            result += f"Subject: {appointment['subject']}\n"
            result += f"Start: {appointment['start_time']}\n"
            result += f"End: {appointment['end_time']}\n"
            result += f"Location: {appointment['location']}\n"
            result += f"All Day: {'Yes' if appointment['is_all_day'] else 'No'}\n"
            result += f"Recurring: {'Yes' if appointment['is_recurring'] else 'No'}\n\n"
        
        result += "To view the full details of an appointment, use the get_appointment_by_number tool with the appointment number."
        return result
    
    except Exception as e:
        return f"Error searching calendar appointments: {str(e)}"

@mcp.tool()
def get_appointment_by_number(appointment_number: int) -> str:
    """
    Get detailed information of a specific appointment by its number from the last listing
    
    Args:
        appointment_number: The number of the appointment from the list results
        
    Returns:
        Full details of the specified appointment
    """
    try:
        if not calendar_cache:
            return "Error: No appointments have been listed yet. Please use list_calendar_appointments or search_calendar_appointments first."
        
        if appointment_number not in calendar_cache:
            return f"Error: Appointment #{appointment_number} not found in the current listing."
        
        appointment_data = calendar_cache[appointment_number]
        
        # Format the output
        result = f"Appointment #{appointment_number} Details:\n\n"
        result += f"Subject: {appointment_data['subject']}\n"
        result += f"Start Time: {appointment_data['start_time']}\n"
        result += f"End Time: {appointment_data['end_time']}\n"
        result += f"Location: {appointment_data['location']}\n"
        result += f"Organizer: {appointment_data['organizer']}\n"
        
        if appointment_data['attendees']:
            result += f"Attendees: {', '.join(appointment_data['attendees'])}\n"
        
        result += f"All Day Event: {'Yes' if appointment_data['is_all_day'] else 'No'}\n"
        result += f"Recurring: {'Yes' if appointment_data['is_recurring'] else 'No'}\n"
        result += f"Reminder: {appointment_data['reminder_minutes']} minutes before\n"
        result += f"Categories: {appointment_data['categories']}\n"
        
        # Busy status mapping
        busy_status_map = {0: 'Free', 1: 'Tentative', 2: 'Busy', 3: 'Out of Office'}
        result += f"Busy Status: {busy_status_map.get(appointment_data['busy_status'], 'Unknown')}\n"
        
        if appointment_data['body']:
            result += f"\nDescription:\n{appointment_data['body']}\n"
        
        return result
    
    except Exception as e:
        return f"Error retrieving appointment details: {str(e)}"

@mcp.tool()
def create_calendar_appointment(subject: str, start_time: str, end_time: str, location: Optional[str] = None, body: Optional[str] = None, attendees: Optional[str] = None) -> str:
    """
    Create a new calendar appointment
    
    Args:
        subject: Subject/title of the appointment
        start_time: Start time in format 'YYYY-MM-DD HH:MM'
        end_time: End time in format 'YYYY-MM-DD HH:MM'
        location: Location of the appointment (optional)
        body: Description/body of the appointment (optional)
        attendees: Comma-separated list of attendee email addresses (optional)
        
    Returns:
        Status message indicating success or failure
    """
    try:
        # Connect to Outlook
        outlook, _ = connect_to_outlook()
        
        # Create a new appointment
        appointment = outlook.CreateItem(1)  # 1 is the value for an appointment item
        appointment.Subject = subject
        
        # Parse and set start/end times
        try:
            appointment.Start = datetime.datetime.strptime(start_time, '%Y-%m-%d %H:%M')
            appointment.End = datetime.datetime.strptime(end_time, '%Y-%m-%d %H:%M')
        except ValueError:
            return "Error: Invalid date format. Please use 'YYYY-MM-DD HH:MM' format."
        
        if location:
            appointment.Location = location
        
        if body:
            appointment.Body = body
        
        # Add attendees if provided
        if attendees:
            attendee_list = [email.strip() for email in attendees.split(',')]
            for attendee in attendee_list:
                if attendee:
                    appointment.Recipients.Add(attendee)
        
        # Save the appointment
        appointment.Save()
        
        return f"Calendar appointment '{subject}' created successfully for {start_time} - {end_time}"
    
    except Exception as e:
        return f"Error creating calendar appointment: {str(e)}"

# Run the server
if __name__ == "__main__":
    print("Starting Outlook MCP Server...")
    print("Connecting to Outlook...")
    
    try:
        # Test Outlook connection
        outlook, namespace = connect_to_outlook()
        inbox = namespace.GetDefaultFolder(6)  # 6 is inbox
        calendar = namespace.GetDefaultFolder(9)  # 9 is calendar
        print(f"Successfully connected to Outlook. Inbox has {inbox.Items.Count} items.")
        print(f"Calendar folder accessed successfully.")
        
        # Run the MCP server
        print("Starting MCP server. Press Ctrl+C to stop.")
        mcp.run()
    except Exception as e:
        print(f"Error starting server: {str(e)}")
