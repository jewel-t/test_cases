import win32com.client

# Connect to Outlook application
outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")

# Get the Global Address List
gal = outlook.Session.GetGlobalAddressList()

# Define a recursive function to traverse the organization hierarchy
def traverse_org_hierarchy(employee_code, direct_reports):
    print(f"Searching for employee with code {employee_code}")
    # Get the contact information for the employee
    employee = gal.AddressEntries.GetFirst()
    while employee:
        print(f"Checking employee {employee.Name} with type {employee.Type}")
        if employee.Type == "EX":
            exchange_user = employee.GetExchangeUser()
            if exchange_user is not None:
                manager = exchange_user.GetExchangeUserManager()
                if manager is not None:
                    manager_user = manager.GetExchangeUser()
                    if manager_user is not None and manager_user.EmployeeID == employee_code:
                        # Add the employee to the direct reports list
                        direct_reports.append(employee)
                        print(f"Found direct report {employee.Name}")
                        # Recursively traverse the direct reports of the employee
                        traverse_org_hierarchy(exchange_user.EmployeeID, direct_reports)
        employee = gal.AddressEntries.GetNext()
        if employee is None:
            break

# Get all the direct reports till the leaf layer for a particular employee code
employee_code = "12345" # Replace with the employee code you want to query
direct_reports = []
traverse_org_hierarchy(employee_code, direct_reports)

# Print the name and email address of each direct report
for report in direct_reports:
    print(report.Name)
    print(report.Address)
    
# Print the number of direct reports found
print(f"Total direct reports found: {len(direct_reports)}")
