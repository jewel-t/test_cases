import win32com.client

outlook = win32com.client.Dispatch("Outlook.Application")
namespace = outlook.GetNamespace("MAPI")

email_address = "jsmith@example.com"  # Replace with the email address
recipient = namespace.CreateRecipient(email_address)
if recipient.Resolve():
    exchange_user = recipient.AddressEntry.GetExchangeUser()
    if exchange_user:
        print(f"Alias: {exchange_user.Alias}")
        print(f"Name: {exchange_user.Name}")
        print(f"Department: {exchange_user.Department}")
    else:
        print(f"No ExchangeUser found for {email_address}")
else:
    print(f"Unable to resolve {email_address}")
