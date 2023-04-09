import win32com.client

outlook = win32com.client.Dispatch("Outlook.Application")
namespace = outlook.GetNamespace("MAPI")
gal = namespace.AddressLists.Item("Global Address List")
employee_alias = "John.Smith"  # Update with your own employee alias
subordinate_dict = {}  # This dictionary will store the subordinate information

def find_subordinates(employee_alias):
    try:
        employee = gal.AddressEntries(employee_alias).GetExchangeUser()
        subordinates = employee.GetDirectReports()
        subordinate_dict[employee_alias] = [subordinate.Alias for subordinate in subordinates]
        for subordinate in subordinates:
            find_subordinates(subordinate.Alias)
    except Exception as e:
        print(f"Error finding subordinates for employee {employee_alias}: {e}")

find_subordinates(employee_alias)
print(subordinate_dict)
