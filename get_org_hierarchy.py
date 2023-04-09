import win32com.client

outlook = win32com.client.Dispatch("Outlook.Application")
namespace = outlook.GetNamespace("MAPI")
gal = namespace.AddressLists.Item("Global Address List")
employee_email = "john.smith@example.com"  # Replace with employee email address
subordinate_dict = {}

visited = set()  # Keep track of visited employees

def find_subordinates(employee_email):
    try:
        employee = gal.AddressEntries.GetExchangeUserFromSMTPAddress(employee_email)
        subordinates = employee.GetDirectReports()
        subordinate_dict[employee_email] = [subordinate.PrimarySmtpAddress for subordinate in subordinates]
        visited.add(employee_email)  # Add the current employee to the visited set
        if subordinates:
            for subordinate in subordinates:
                if subordinate.PrimarySmtpAddress not in visited:  # Only visit new employees
                    find_subordinates(subordinate.PrimarySmtpAddress)
    except AttributeError:
        print(f"AttributeError: Email address {employee_email} is not associated with an Exchange user.")
    except Exception as e:
        print(f"Error finding subordinates for employee {employee_email}: {e}")

find_subordinates(employee_email)
print(subordinate_dict)
