import win32com.client

outlook = win32com.client.Dispatch("Outlook.Application")
namespace = outlook.GetNamespace("MAPI")
gal = namespace.AddressLists.Item("Global Address List")
employee_email = "john.smith@example.com"  # Replace with employee email address
subordinate_dict = {}

def find_subordinates(employee_email, depth=0):
    try:
        employee = gal.AddressEntries.GetExchangeUserFromSMTPAddress(employee_email)
        subordinates = employee.GetDirectReports()
        subordinate_dict[employee_email] = [subordinate.PrimarySmtpAddress for subordinate in subordinates]
        if subordinates:
            for subordinate in subordinates:
                find_subordinates(subordinate.PrimarySmtpAddress, depth+1)
    except Exception as e:
        print(f"Error finding subordinates for employee {employee_email}: {e}")

find_subordinates(employee_email)
print(subordinate_dict)
