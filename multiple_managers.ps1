# Install the AzureAD module if you haven't already
# Install-Module -Name AzureAD

# Import the AzureAD module
Import-Module AzureAD

# Define an array of manager's display names or UPNs (User Principal Names)
$managerNames = @("abc@gmail.com", "xyz@gmail.com")

# Connect to Azure AD
Connect-AzureAD

# Initialize Excel
$excel = New-Object -ComObject Excel.Application
$workbook = $excel.Workbooks.Add()

# Iterate over each manager
foreach ($managerName in $managerNames) {
    # Retrieve the manager's user object
    $manager = Get-AzureADUser -Filter "userPrincipalName eq '$managerName'"

    if ($manager -ne $null) {
        # Get direct reports for the manager
        $directReports = Get-AzureADUserDirectReport -ObjectId $manager.ObjectId

        if ($directReports.Count -gt 0) {
            # Add a new worksheet/tab for the manager
            $sheet = $workbook.Worksheets.Add()
			$sheetName = $manager.GivenName -replace '[\\\/\:\?\*\[\]]', '_'
            Write-Host "Generated Sheet Name: $sheetName"
            $sheet.Name = $sheetName + "'s Direct Reports"

			#$sheetName = $manager.DisplayName -replace '[\\\/\:\?\*\[\]]', '_'
            #$sheet.Name = $sheetName + "'s Direct Reports"

            #$sheet.Name = $manager.Name + "Direct Reports"
            $row = 2

            # Add headers
            $sheet.Cells.Item(1,1) = "Name"
            $sheet.Cells.Item(1,2) = "UserPrincipalName"
            $sheet.Cells.Item(1,3) = "Department"
            $sheet.Cells.Item(1,4) = "JobTitle"
            $sheet.Cells.Item(1,5) = "Country"
			$sheet.Cells.Item(1,6) = "City"

            # Set header row to yellow color
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
        } else {
            Write-Host "No direct reports found for $($manager.DisplayName)."
        }
    } else {
        Write-Host "Manager $managerName not found."
    }
}

# Save Excel file
$excel.Visible = $true
$workbook.SaveAs("consolidated.xlsx")
$excel.Quit()
Write-Host "Employee information has been exported to consolidated.xlsx"
