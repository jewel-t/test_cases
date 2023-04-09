import win32com.client

outlook = win32com.client.Dispatch("Outlook.Application")
namespace = outlook.GetNamespace("MAPI")
gal = namespace.AddressLists.Item("Global Address List")
employee_id = "John.Smith@company.com"  # Update with your own employee ID
subordinate_dict = {}  # This dictionary will store the subordinate information

def find_subordinates(employee_id):
    try:
        employee = gal.AddressEntries(employee_id).GetExchangeUser()
        subordinates = employee.GetDirectReports()
        subordinate_dict[employee_id] = [subordinate.PrimarySmtpAddress for subordinate in subordinates]
        for subordinate in subordinates:
            find_subordinates(subordinate.PrimarySmtpAddress)
    except Exception as e:
        print(f"Error finding subordinates for employee {employee_id}: {e}")

find_subordinates(employee_id)
print(subordinate_dict)
