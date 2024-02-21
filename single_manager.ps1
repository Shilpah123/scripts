# Install the AzureAD module if you haven't already
# Install-Module -Name AzureAD

# Import the AzureAD module
Import-Module AzureAD

# Define manager's display name or UPN (User Principal Name)
$managerName = "Add email ID here"

# Connect to Azure AD
Connect-AzureAD

# Retrieve the manager's user object
$manager = Get-AzureADUser -Filter "userPrincipalName eq '$managerName'"

if ($manager -ne $null) {
    # Get direct reports for the manager
    $directReports = Get-AzureADUserDirectReport -ObjectId $manager.ObjectId

    if ($directReports.Count -gt 0) {
        # Initialize Excel
        $excel = New-Object -ComObject Excel.Application
        $workbook = $excel.Workbooks.Add()
        $sheet = $workbook.Worksheets.Item(1)
		$sheetName = $manager.GivenName -replace '[\\\/\:\?\*\[\]]', '_'
        Write-Host "Generated Sheet Name: $sheetName"
        $sheet.Name = $sheetName + "'s Direct Reports"
        $row = 2

        # Add headers
        $sheet.Cells.Item(1,1) = "Name"
        $sheet.Cells.Item(1,2) = "UserPrincipalName"
        $sheet.Cells.Item(1,3) = "Department"
        $sheet.Cells.Item(1,4) = "JobTitle"
        $sheet.Cells.Item(1,5) = "Country"
		$sheet.Cells.Item(1,6) = "City"
		
		#Set header row to yellow color
		$headerRange = $sheet.Range("A1", "F1")
        $headerRange.Interior.Color = 65535  # Yellow color code
        
        # Retrieve and print employee information
        foreach ($employee in $directReports) {
            $employeeDetails = Get-AzureADUser -ObjectId $employee.ObjectId
            $sheet.Cells.Item($row,1) = $employeeDetails.DisplayName
            $sheet.Cells.Item($row,2) = $employeeDetails.UserPrincipalName
            $sheet.Cells.Item($row,3) = $employeeDetails.Department
            $sheet.Cells.Item($row,4) = $employeeDetails.JobTitle
            $sheet.Cells.Item($row,5) = $employeeDetails.Country
			$sheet.Cells.Item($row,6) = $employeeDetails.City
            $row++
        }

        # Save Excel file
        $excel.Visible = $true
        $workbook.SaveAs("employee.xlsx")
        $excel.Quit()
        Write-Host "Employee information has been exported to employee.xlsx"
    } else {
        Write-Host "No direct reports found for the manager."
    }
} else {
    Write-Host "Manager not found."
}
