import win32com.client

# Create an instance of the Outlook application
outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")

# Get the unread folder
unread_folder = outlook.GetDefaultFolder(6)

# Get all unread emails in the folder, sorted by oldest to newest
unread_emails = unread_folder.Items.Restrict("[Unread]=True")
unread_emails.Sort("[ReceivedTime]")

# Loop through the unread emails
for email in unread_emails:
    # Open the oldest email
    email_display = email.Display()

    # Read the email content
    email_body = email.Body

    # Create a new file and save the email content to it
    with open('1.txt', 'w') as file:
        file.write(email_body)

    # Close the email
    email_display.Close(0)

    # Exit the loop after processing the oldest email
    break
