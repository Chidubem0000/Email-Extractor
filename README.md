# Email-Extractor
Contact Email Extractor 
Email Receiver Extractor

This Python script is designed to connect to an Office 365 email account, list available folders, and then extract and save the email addresses of the receivers (To:) from a specified email folder. Below are the instructions on how to use this script:
Prerequisites

Before using this script, make sure you have the following prerequisites in place:

    Python: Ensure you have Python installed on your system.

    Required Libraries: You will need the following Python libraries installed. You can install them using pip:
        imaplib: This library provides functionality to interact with IMAP servers.
        email: This library is used for working with email messages.

    Office 365 Email Account: You need to have access to an Office 365 email account, and you will need your email address and password.

Usage Instructions

    Clone or download the script to your local machine.

    Open the script in a text editor or integrated development environment (IDE).

    Replace the following placeholders with your Office 365 email and password:
        username: Replace with your Office 365 email address.
        password: Replace with your email account password.

    Specify the folder you want to extract email addresses from by setting the folder_name_to_select variable. Make sure to use lowercase for the folder name.

    Set the output_directory variable to the directory where you want to save the extracted email addresses. If the directory doesn't exist, it will be created.

    Save your changes and run the script. It will connect to your Office 365 account, select the specified folder, and extract the email addresses from the emails in that folder.

    The extracted email addresses will be saved in a text file named receivers.txt in the specified output_directory.

    The script has a rate-limiting feature that waits for a specified interval (default: 10 seconds) after processing a certain number of emails (default: 50 emails). You can adjust these parameters in the fetch_emails_with_limit function if needed.

    The script will print information about the progress and any errors encountered during execution.

    Example

Here's an example of how to use the script:

username = "your_office365_email@example.com"
password = "your_password"
folder_name_to_select = "sent"  # Replace with the folder you want to select
output_directory = "email_addresses"

# Rest of the script remains unchanged




Ensure you keep your email and password secure and do not share them with others. This script is intended for personal use and may require modifications for use in production environments.
