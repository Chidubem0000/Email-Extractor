import os
import imaplib
import email
import concurrent.futures

def list_available_folders(username, password):
    # Connect to the IMAP server with a timeout
    imap_server = imaplib.IMAP4_SSL("outlook.office365.com", timeout=None)  # Increased timeout to None for infinite
    imap_server.login(username, password)  # Pass the username and password as strings

    # Get a list of available mailbox folders
    status, mailbox_list = imap_server.list()
    available_folders = []

    if status == "OK":
        for folder_info in mailbox_list:
            folder_name_info = folder_info.decode("utf-8").split()[-1].lower()
            available_folders.append(folder_name_info)

    # Close the connection
    imap_server.logout()

    return available_folders

def fetch_and_save_receiver_emails(imap_server, email_id, output_directory):
    try:
        status, email_data = imap_server.fetch(email_id, "(BODY[HEADER.FIELDS (To)])")
        if status == "OK":
            email_message = email.message_from_bytes(email_data[0][1])
            email_receiver = email_message["to"]
            print(f"Receiver: {email_receiver}")

            # Save the receiver email address in a text file
            filename = os.path.join(output_directory, "receivers.txt")
            with open(filename, "a") as receiver_file:
                receiver_file.write(email_receiver + "\n")
    except Exception as e:
        print(f"Error fetching email {email_id.decode('utf-8')}: {e}")

def select_and_save_receiver_emails(username, password, folder_name, output_directory):
    # Connect to the IMAP server with a timeout
    imap_server = imaplib.IMAP4_SSL("outlook.office365.com", timeout=None)  # Increased timeout to None for infinite
    imap_server.login(username, password)  # Pass the username and password as strings

    # Convert the folder name to lowercase
    folder_name = folder_name.lower()

    # Get a list of available mailbox folders
    available_folders = list_available_folders(username, password)

    # Check if the specified folder exists
    if folder_name not in available_folders:
        print(f"Folder '{folder_name}' not found.")
        print("Available folders:")
        for folder in available_folders:
            print(folder)
        return

    # Select the specified folder (case-insensitive)
    status, _ = imap_server.select(folder_name)
    if status == "OK":
        # You can now perform operations on the selected folder
        print(f"Selected folder: {folder_name}")

        # Fetch emails with the "\Seen" flag
        status, email_ids = imap_server.search(None, 'SEEN')
        if status == "OK":
            email_id_list = email_ids[0].split()

            # Use multithreading for faster extraction
            with concurrent.futures.ThreadPoolExecutor(max_workers=5) as executor:
                email_id_list = [email_id for email_id in email_id_list]  # Convert to a list if not already
                for email_id in email_id_list:
                    executor.submit(fetch_and_save_receiver_emails, imap_server, email_id, output_directory)

        # Close the selected mailbox
        imap_server.close()

    # Logout from the IMAP server
    imap_server.logout()

if __name__ == "__main__":
    # Replace with your Office 365 email and password
    username = "Sherri-87@hotmail.com"
    password = "Teheresa12"

    # List available folders
    available_folders = list_available_folders(username, password)
    print("Available folders:")
    for folder in available_folders:
        print(folder)

    # Select and save receiver email addresses from a specific folder (case-insensitive)
    folder_name_to_select = "sent"  # Replace with the folder you want to select (use lowercase)
    output_directory = "email_raw_data"  # Change this to your desired directory

    if not os.path.exists(output_directory):
        os.makedirs(output_directory)  # Create the directory if it doesn't exist

    select_and_save_receiver_emails(username, password, folder_name_to_select, output_directory)
