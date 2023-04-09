import win32com.client

outlook = win32com.client.Dispatch("Outlook.Application")
namespace = outlook.GetNamespace("MAPI")

def get_subordinates(email_address, level=0):
    recipient = namespace.CreateRecipient(email_address)
    if recipient.Resolve():
        exchange_user = recipient.AddressEntry.GetExchangeUser()
        if exchange_user:
            print("|   " * level + "+-- " + exchange_user.Name)
            for subordinate in exchange_user.GetDirectReports():
                subordinate_email = subordinate.GetExchangeUser().PrimarySmtpAddress
                get_subordinates(subordinate_email, level=level+1)
        else:
            print("|   " * level + "+-- " + email_address + " (no ExchangeUser found)")
    else:
        print("|   " * level + "+-- " + email_address + " (unable to resolve)")
        
# Replace with the email address of the employee whose subordinates you want to retrieve
get_subordinates("jsmith@example.com")
