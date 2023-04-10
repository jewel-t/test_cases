import win32com.client
import csv
import time

start_time = time.time()
index = 0

def get_subordinates(user, level, writer):
    """
    Recursively retrieves the subordinates of a given user and prints the hierarchy tree to console and CSV file.
    """

    # Check if user exists
    if user is None:
        return

    # Get the user's name and title
    name = user.Name
    title = user.JobTitle

    # Print the user's details to console
    global index
    print(str(index) + " " * level + "- " + name + " (" + title + ")")
    index += 1

    # Write the user's details to CSV file
    writer.writerow([name, title, user.Manager.Name if user.Manager else ""])

    # Get the user's subordinates
    subordinates = user.GetDirectReports()

    # Recursively call this function on each subordinate
    for subordinate in subordinates:
        get_subordinates(subordinate, level + 1, writer)


# Connect to Outlook
outlook = win32com.client.Dispatch("Outlook.Application")
namespace = outlook.GetNamespace("MAPI")

# Get the root user (CEO)
ceo = namespace.CreateRecipient("CEO's Email Address").AddressEntry.GetExchangeUser()

# Open the CSV file
with open("hierarchy.csv", "w", newline="") as csvfile:
    writer = csv.writer(csvfile)

    # Write the header row to CSV file
    writer.writerow(["Name", "Title", "Manager"])

    # Recursively get the subordinates of the CEO and print the hierarchy tree to console and CSV file
    get_subordinates(ceo, 0, writer)

    end_time = time.time()

elapsed_time = end_time - start_time

minutes, seconds = divmod(elapsed_time, 60)
hours, minutes = divmod(minutes, 60)

print(f"Time taken: {hours:.0f} hours, {minutes:.0f} minutes, {seconds:.2f} seconds")
