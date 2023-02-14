#! Copyright (c) Bryan Hunter 2022

# Built-in modules:
import getpass
from imap_tools import MailBox
from email import *
import os
from os import path
from time import sleep
from datetime import datetime
from logging import basicConfig, getLogger, INFO, DEBUG

# Third-party modules:
from dotenv import load_dotenv
from progress.bar import Bar
from wordcloud import WordCloud

# Load the environment variables:
load_dotenv()

# Prompt the user for confirmation:
def confirm(prompt: str, default: bool = False):
    if default is None:
        prompt += " [y/n]: "
    elif default:
        prompt += " [Y/n]: "
    else:
        prompt += " [y/N]: "
    while True:
        choice = input(prompt).lower()
        if default is not None and choice == '':
            return default
        elif choice in ['y', 'yes']:
            return True
        elif choice in ['n', 'no']:
            return False
        else:
            print("Please respond with 'yes' or 'no' (or 'y' or 'n').\n")
logger = getLogger("Missionary Word Cloud Generator")

# Set up logging:
def setup_logger():
    # Check to make sure the logs folder exists:
    if not path.exists(os.getcwd() + os.sep + "logs"):
        os.mkdir(os.getcwd() + os.sep + "logs")
    basicConfig(level=INFO, format='%(asctime)s - %(name)s - %(levelname)s:==> %(message)s',
                filename=os.getcwd() + os.sep + "logs" + os.sep + 'missionary_word_cloud.log', filemode='w', datefmt='%m/%d/%Y %I:%M:%S %p')
    if os.getenv("WC_DEBUG").lower == "true" or os.getenv("WC_DEBUG") == "1":
        logger.setLevel(DEBUG)

# Receive input from the user:
def get_input(input_str: str, default: str = None, password: bool = False):
    if default is None and not password:
        return input(input_str)
    if password:
        return getpass.getpass(input_str + " [{}]: ".format(default)) or default
    else:
        return input(input_str + " [{}]: ".format(default)) or default
    
setup_logger()
print("Welcome to the Missionary Word Cloud Generator!\n")
print("This program will generate a word cloud from the emails in your specified folder.")
print("Please note that this program will not delete any emails from your inbox.")
print("You will be prompted to enter your Gmail App password.")
print("You can find more information on app passwords here: https://support.google.com/accounts/answer/185833?hl=en")
print("Also, please make sure that you have enabled IMAP in your Gmail settings.")
print("You can find more information on that here: https://support.google.com/mail/answer/7126229?hl=en")
print("If you have any questions, please contact me at admin@themoddersden.com.")
print("Thank you for using this program!\n")
print("Press enter to continue...")
input()

# Variables:
subfolder = get_input("Enter the subfolder to search: ", default="wordcloud")
logger.debug(f"Subfolder: {subfolder}")

cloud_outfile = get_input("Enter the path to save the word cloud: ", default=os.getcwd() + os.sep + "wordcloud.png")
logger.debug(f"Cloud outfile: {cloud_outfile}")

email_outpath = get_input("Enter the path to save the emails: ", default=os.getcwd() + os.sep + "emails" + os.sep)
logger.debug(f"Email outpath: {email_outpath}")

email = get_input("Enter your Gmail address: ", password=False)
logger.debug(f"Email: {email}")

password = get_input("Enter your Gmail App password: ", password=True)
logger.debug(f"Password hash: {hash(password)}")

sender_address = get_input("Enter the sender's email address: ", default="missionary@missionary.org")
logger.debug(f"Sender address: {sender_address}")

email_search_topic = get_input("Enter a query to help filter the emails: ", default=None)
logger.debug(f"Email search topic: {email_search_topic}")

ignore_replies = confirm("Ignore replies?", default=True)
logger.debug(f"Ignore replies: {ignore_replies}")

# Get the current date and time:
def get_date_time() -> str:
    now = datetime.now(get_timezone())
    new_now = now.strftime("%m/%d/%Y, %H:%M:%S")
    current_tz = get_timezone()
    logger.debug(f"Now: {new_now}")
    logger.debug(f"Timezone: {current_tz}")
    return new_now

# Get the local timezone:
def get_timezone() -> datetime.tzinfo:
    return datetime.now().astimezone().tzinfo

# Get all emails from the specified folder:
def get_emails(sub_folder: str = None) -> list:

    my_messages = []

    # Connect to the mailbox:
    with MailBox('imap.gmail.com', timeout=90).login(email, password) as mailbox:
        with Bar('Processing Emails') as bar:
            for msg in mailbox.fetch(mark_seen=False):
                mailbox.folder.set(sub_folder)
                if msg.from_ == os.getenv(sender_address) and msg.subject.startswith(email_search_topic):
                    logger.debug(msg.from_, msg.subject)
                    if ignore_replies and msg.subject.startswith("Re:") or msg.subject.startswith("RE:"):
                        logger.debug("Ignoring reply...")
                    else:
                        logger.debug(msg.date, msg.subject,
                                     len(msg.text or msg.html))
                        my_messages.append(msg.text or msg.html)
                        bar.next()

        # Return the list of messages:
        return my_messages

# Export each email to a text file:
def export_emails(emails: list):
    count = 0
    for email in emails:
        # Try to write the files:
        try:
            # Check if the path exists:
            if not path.exists(email_outpath):
                os.mkdir(email_outpath)
            # Write the files:
            with open(email_outpath + f"Missionary_Email_{count}.txt", "w") as f:
                f.write(email)
                count += 1
        # Handle errors:
        except (FileNotFoundError, OSError) as e:
            logger.debug(
                "Error writing file. Please check the file path and try again.")
            print(e)
            logger.critical(f"Stack trace: {e}")
            input("Press enter to exit...")
            exit(1)

# Make a word cloud from a list of emails:
def make_word_cloud(emails: list):

    # Join the different processed titles together.
    long_string = ','.join(emails)

    # Create a WordCloud object
    logger.info("Generating word cloud...")

    wordcloud = WordCloud(background_color="white", max_words=5000, width=1920, height=1080,
                          contour_width=3, contour_color='steelblue', collocations=True)  # , stopwords=stopwords

    # Generate a word cloud
    wordcloud.generate(long_string)
    wordcloud.to_file(cloud_outfile)

    # Finish up:
    logger.info(f"Word cloud saved as '{cloud_outfile}'")
    sleep(3)
    logger.info("Done!")

# Main function:
def main():
    if confirm("Do you wish to continue?", default=None):
        logger.debug("Getting emails...")
        emails = get_emails(sub_folder="wordcloud")
        export_emails(emails)
        make_word_cloud(emails)
    else:
        logger.debug("User cancelled program.")
        print("Program cancelled.")
        sleep(3)
        print("Goodbye!")
        logger.debug("Exiting program...")
        input("Press enter to exit...")
        exit(0)

# Run the main function:
if __name__ == "__main__":
    main()