# New User Account Process
# Written by J Spitz
# Updated 6-20-19 Added new status_id to match site updates
# Finished 9-10

# To be used with SQL Server and Office 365
# If servers are migrated to Azure, seach for HSG-SERVER in this file to find where servers are referenced

# If support scripts cannot be reached it will null data rows and terminate.


# **********************************************************************************
#   Hyperparameters
# **********************************************************************************
#region

# Default Password
$DefaultPassword = ConvertTo-SecureString "" -AsPlainText -Force

# Default Department
$DefaultDepartment = ""

# Location of creation functions
$creationDirectory = ""

# Date/Time for description updates
$time = [datetime]::Now

#endregion  

# **********************************************************************************
#   Log File Creation
# **********************************************************************************
#region Logs
# Specify Log File
$LogFile = "\Logs\$(Get-Date -Format s | ForEach-Object {$_ -replace ":", "."}).log"
#endregion

# **********************************************************************************
#   Email Parameters
# **********************************************************************************
#region Email templates and sender address

# Sender address for automated emails
$systemsEmailAddress = ""

# Subject header for emails
$emailSubject = "User Account Request Confirmation"

# Email Bodies for Account Submitter and Manager

    # Single user requested
    [string]$emailBodySUR = Get-Content -Path "\Creation Email Templates\acctReqEmailSingleUser.txt"

    # Multiple Users requested
    [string]$emailBodyMUR = Get-Content -Path "\Creation Email Templates\acctReqEmailMultipleUsers.txt"

    # Single user for supervisor
    [string]$emailBodyMSU = Get-Content -Path "\Creation\Creation Email Templates\acctManEmailSingleUser.txt"

    # Multiple Users for supervisor
    [string]$emailBodyMMU = Get-Content -Path "\Creation\Creation Email Templates\acctManEmailMultUser.txt"

#endregion    

# **********************************************************************************
#   Modules, Server, Exclusion Settings
# **********************************************************************************
#region Setup
Import-Module ActiveDirectory
Import-Module SqlServer

# Connection Strings
# SQL Server Connection String (Uses SQL Server Authentication)
$sqlServer = ""
$sqlDatabase = "useraccounts"
$sqlUsername = ''
$sqlPassword = ''

# Profile Server
$ProfileServerDir = ""

# WiscMail Connection String
$emailServer = ""
$emailPort = ""
#$smtp = New-Object System.Net.Mail.SmtpClient($emailserver, $emailPort);

# Connect to Resources
# No Username or Password 
#$global:mailCon = Connect-Email -Server $emailServer -Port $emailPort -Username $emailSenderAcct

# Load Service Account inclusion list
$serviceAccountPositionsFile = $creationDirectory + "ServiceAccountInclusions.txt"

# Load Group Exception List
#endregion

# **********************************************************************************
#   Load Additonal Powershell Functions
# **********************************************************************************
#region Custom PS imports

#endregion

# **********************************************************************************
#   Important SQL Schema Translations
# **********************************************************************************
#region Schema Notes

# Status_ID:
# 0 = Request Error (Generic)
# 1 = Submitted for Creation Approval on the website
# 2 = Submitted for Deletion on the website
# 3 = Creation Request Approved
# 4 = (Creation/Deletion) Request Denied
# 5 = Account Created -- Status for Active Accounts
# 6 = Deletion Request Approved
# 7 = Account Deleted
# 8 = Account Disable Request for Summer
# 9 = Deletion Holding Cell -- Used to hold accounts until End of Appointment has passed so accounts aren't deleted early
# 10 = Account Renable Request -> Sets back to '5'
# 11 = Disabled for Summer
#
# Department_ID:
# 0 = Not Used
# 1 = Administration
# 2 = Dining and Culinary Services
# 3 = Information Technology
# 4 = Residence Halls Facilities
# 5 = Residence Life
# 6 = University Apartments
# 7 = Business Services
# 8 = Communications and Residential Programs
# 9 = Human Resources and Payroll

#endregion

# **********************************************************************************
#   Account Functions -- Default Password: 'Housing!' -- Location: $DefaultPassword
# **********************************************************************************
#region Queries and default OU

# Queries User Account Database 'accounts' table, outputs data rows 
#region Account status SQL queries
$newCreationRequests = Invoke-Sqlcmd -Username $sqlUsername -Password $sqlPassword -ServerInstance $sqlServer -Database $sqlDatabase -OutputAs DataTables -Query "SELECT        accounts.accounttype_id, accounts.status_id, accounts.unit_id, accounts.building_id, accounts.department_id, accounts.supervisor, accounts.first_name, accounts.middle_initial, accounts.last_name, 
accounts.campusid, accounts.copyfrom, accounts.netid, accounts.title, accounts.address1, accounts.address2, accounts.phone_number, accounts.fax_number, accounts.mobile_number, accounts.start_date, 
accounts.end_date, accounts.entered_by, accounts.user_agreement, accounts.access_email, accounts.access_computers, accounts.notes, accounts.access_uas, accounts.access_abs, accounts.access_kronos, 
accounts.access_mrs, accounts.access_crs, accounts.access_chris, accounts.access_pvl, accounts.access_sea, accounts.access_warehouse, accounts.access_isis, accounts.access_3270, 
accounts.access_wisdm, accounts.access_blackboard, accounts.access_cbord, accounts.access_TMA, accounts.kiosk_training, accounts.it_orientation, accounts.kiosk_training_translator
FROM            accounts INNER JOIN
buildings ON accounts.building_id = buildings.ID INNER JOIN
departments ON accounts.department_id = departments.ID INNER JOIN
units ON accounts.unit_id = units.ID
WHERE        (accounts.status_id = 3)"

$newDeletionRequests = Invoke-Sqlcmd -Username $sqlUsername -Password $sqlPassword -ServerInstance $sqlServer -Database $sqlDatabase -OutputAs DataTables -Query "SELECT        accounts.accounttype_id, accounts.status_id, accounts.unit_id, accounts.building_id, accounts.department_id, accounts.supervisor, accounts.first_name, accounts.middle_initial, accounts.last_name, 
accounts.campusid, accounts.copyfrom, accounts.netid, accounts.title, accounts.address1, accounts.address2, accounts.phone_number, accounts.fax_number, accounts.mobile_number, accounts.start_date, 
accounts.end_date, accounts.entered_by, accounts.user_agreement, accounts.access_email, accounts.access_computers, accounts.notes, accounts.access_uas, accounts.access_abs, accounts.access_kronos, 
accounts.access_mrs, accounts.access_crs, accounts.access_chris, accounts.access_pvl, accounts.access_sea, accounts.access_warehouse, accounts.access_isis, accounts.access_3270, 
accounts.access_wisdm, accounts.access_blackboard, accounts.access_cbord, accounts.access_TMA, accounts.kiosk_training, accounts.it_orientation, accounts.kiosk_training_translator
FROM            accounts INNER JOIN
buildings ON accounts.building_id = buildings.ID INNER JOIN
departments ON accounts.department_id = departments.ID INNER JOIN
units ON accounts.unit_id = units.ID
WHERE        (accounts.status_id = 6)"

$newSummerDisableRequests = Invoke-Sqlcmd -Username $sqlUsername -Password $sqlPassword -ServerInstance $sqlServer -Database $sqlDatabase -OutputAs DataTables -Query "SELECT        accounts.accounttype_id, accounts.status_id, accounts.unit_id, accounts.building_id, accounts.department_id, accounts.supervisor, accounts.first_name, accounts.middle_initial, accounts.last_name, 
accounts.campusid, accounts.copyfrom, accounts.netid, accounts.title, accounts.address1, accounts.address2, accounts.phone_number, accounts.fax_number, accounts.mobile_number, accounts.start_date, 
accounts.end_date, accounts.entered_by, accounts.user_agreement, accounts.access_email, accounts.access_computers, accounts.notes, accounts.access_uas, accounts.access_abs, accounts.access_kronos, 
accounts.access_mrs, accounts.access_crs, accounts.access_chris, accounts.access_pvl, accounts.access_sea, accounts.access_warehouse, accounts.access_isis, accounts.access_3270, 
accounts.access_wisdm, accounts.access_blackboard, accounts.access_cbord, accounts.access_TMA, accounts.kiosk_training, accounts.it_orientation, accounts.kiosk_training_translator
FROM            accounts INNER JOIN
buildings ON accounts.building_id = buildings.ID INNER JOIN
departments ON accounts.department_id = departments.ID INNER JOIN
units ON accounts.unit_id = units.ID
WHERE        (accounts.status_id = 8)"

$newSummerEnableRequests = Invoke-Sqlcmd -Username $sqlUsername -Password $sqlPassword -ServerInstance $sqlServer -Database $sqlDatabase -OutputAs DataTables -Query "SELECT        accounts.accounttype_id, accounts.status_id, accounts.unit_id, accounts.building_id, accounts.department_id, accounts.supervisor, accounts.first_name, accounts.middle_initial, accounts.last_name, 
accounts.campusid, accounts.copyfrom, accounts.netid, accounts.title, accounts.address1, accounts.address2, accounts.phone_number, accounts.fax_number, accounts.mobile_number, accounts.start_date, 
accounts.end_date, accounts.entered_by, accounts.user_agreement, accounts.access_email, accounts.access_computers, accounts.notes, accounts.access_uas, accounts.access_abs, accounts.access_kronos, 
accounts.access_mrs, accounts.access_crs, accounts.access_chris, accounts.access_pvl, accounts.access_sea, accounts.access_warehouse, accounts.access_isis, accounts.access_3270, 
accounts.access_wisdm, accounts.access_blackboard, accounts.access_cbord, accounts.access_TMA, accounts.kiosk_training, accounts.it_orientation, accounts.kiosk_training_translator
FROM            accounts INNER JOIN
buildings ON accounts.building_id = buildings.ID INNER JOIN
departments ON accounts.department_id = departments.ID INNER JOIN
units ON accounts.unit_id = units.ID
WHERE        (accounts.status_id = 10)"

$EoARequests = Invoke-Sqlcmd -Username $sqlUsername -Password $sqlPassword -ServerInstance $sqlServer -Database $sqlDatabase -OutputAs DataTables -Query "SELECT        accounts.accounttype_id, accounts.status_id, accounts.unit_id, accounts.building_id, accounts.department_id, accounts.supervisor, accounts.first_name, accounts.middle_initial, accounts.last_name, 
accounts.campusid, accounts.copyfrom, accounts.netid, accounts.title, accounts.address1, accounts.address2, accounts.phone_number, accounts.fax_number, accounts.mobile_number, accounts.start_date, 
accounts.end_date, accounts.entered_by, accounts.user_agreement, accounts.access_email, accounts.access_computers, accounts.notes, accounts.access_uas, accounts.access_abs, accounts.access_kronos, 
accounts.access_mrs, accounts.access_crs, accounts.access_chris, accounts.access_pvl, accounts.access_sea, accounts.access_warehouse, accounts.access_isis, accounts.access_3270, 
accounts.access_wisdm, accounts.access_blackboard, accounts.access_cbord, accounts.access_TMA, accounts.kiosk_training, accounts.it_orientation, accounts.kiosk_training_translator
FROM            accounts INNER JOIN
buildings ON accounts.building_id = buildings.ID INNER JOIN
departments ON accounts.department_id = departments.ID INNER JOIN
units ON accounts.unit_id = units.ID
WHERE        (accounts.status_id = 10)"
#endregion

# Queries for the Department names
$Departments = Invoke-Sqlcmd -ServerInstance $sqlServer -Username $sqlUsername -Password $sqlPassword -OutputAs DataTables -Database $sqlDataBase -Query "SELECT        departments.name, departments.id
FROM            departments"

# Queries for the Employee Types
$AccountTypes = Invoke-Sqlcmd -ServerInstance $sqlServer -Username $sqlUsername -Password $sqlPassword -Database $sqlDataBase -OutputAs DataTables -Query "SELECT id, description 
FROM            accounttypes"

# Root OU for AD organization
$RootOU = ''

# Disable OU
$disableOU = ""

# Deletion OU
$deletionOU = ""

# Summer disable prefix
$disableSummer = ''

# Ends script if there are not new accounts
if($null -eq $newCreationRequests.Columns -and $null -eq $newDeletionRequests -and $null -eq $newSummerDisableRequests -and $null -eq $newSummerEnableRequests -and $null -eq $EoARequests){
    Write-Host "There are no new user account requests to process" -ForegroundColor Blue
    Exit
}
#endregion

# **********************************************************************************
#  Account preprocessing
# **********************************************************************************
#region DataTable Preprocessing
# Create new columns for AD parameter fields
    # These additions only work when the container is made of datatables

# Added If statementto avoid null-value method errors
if($null -ne $newCreationRequests.Columns){
    $newCreationRequests.Columns.Add("CN", [system.type]"string")| Out-Null
    $newCreationRequests.Columns.Add("UserPrincpalName", [system.type]"string")| Out-Null
    $newCreationRequests.Columns.Add("SamAccountName", [system.type]"string")| Out-Null
    $newCreationRequests.Columns.Add("ProfilePath", [system.type]"string")| Out-Null
    $newCreationRequests.Columns.Add("OU", [system.type]"string")| Out-Null
    $newCreationRequests.Columns.Add("DisplayName", [system.type]"string")| Out-Null
    $newCreationRequests.Columns.Add("UserPrincipalName", [system.type]"string")| Out-Null
    $newCreationRequests.Columns.Add("Office", [system.type]"string")| Out-Null
    $newCreationRequests.Columns.Add("ManagerNetID", [system.type]"string")| Out-Null
    $newCreationRequests.Columns.Add("DepartmentAsString", [system.type]"string")| Out-Null
    $newCreationRequests.Columns.Add("AccountTypeAsString", [system.type]"string")| Out-Null


    # Variables to hold row counts
    $newAccountCount = 0

    # Count rows for iterations
    foreach($row in $newCreationRequests){
        $newAccountCount = $newAccountCount + 1
    }

    # Update progress
    Write-Host "Checking for New User Requests" -ForegroundColor Yellow
    Write-Log -LogString "[INFO] Checking for New User Requests" -LogFile $LogFile

    # Status Update
    If ($newAccountCount -eq 1) {
        Write-Host "There is 1 new user account request" -ForegroundColor "Green"
        Write-Log -LogString "[INFO] There is 1 new user account request" -LogFile $LogFile
    }
    Else {
        Write-Host "There are $newAccountCount new user account requests" -ForegroundColor "Green"
        Write-Log -LogString "[INFO] There are $newAccountCount new user account requests" -LogFile $LogFile
    }

    # Iterates through table of new user requests and preps data columns for creation
    foreach ($DataRow in $newCreationRequests) {
        # Normalize User's Name
        # First Name
        $DataRow.first_name = $DataRow.first_name.trimstart(" ").trimend(" ")
        $DataRow.first_name = $DataRow.first_name.ToLowerInvariant()
        $DataRow.first_name = $DataRow.first_name.Insert(0, ($DataRow.first_name.ToUpper())[0])
        $DataRow.first_name = $DataRow.first_name.Remove(1, 1)
        $DataRow.first_name = $DataRow.first_name.trimEnd("`"")

        # Last Name
        $DataRow.last_name = $DataRow.last_name.trimstart(" ").trimend(" ")
        $DataRow.last_name = $DataRow.last_name.ToLowerInvariant()
        $DataRow.last_name = $DataRow.last_name.Insert(0, ($DataRow.last_name.ToUpper())[0]) 
        $DataRow.last_name = $DataRow.last_name.Remove(1, 1)
        $DataRow.last_name = $DataRow.last_name.trimEnd("`"")


        # Sets DisplayName to have names in reverse order
        $DataRow.DisplayName = $DataRow.last_name + ", " + $DataRow.first_name
        $DataRow.CN = $DataRow.first_name + " " + $DataRow.last_name

        # Status update to user and log file
        Write-Host "Building account for: " -ForegroundColor "Green" -NoNewLine
        Write-Host $DataRow.CN -ForegroundColor White
        Write-Log -LogString "[INFO] Building account for: $($DataRow.CN)" -LogFile $LogFile

        # Get Department Name for the new users
        foreach($row in $Departments){
            if ($row.id -eq $DataRow.department_id) {
                $DataRow.DepartmentAsString = $Departments.name
                break
            }
        }

        # Get Employee Type for new users
        foreach($row in $AccountTypes){
            if ($row.id -eq $DataRow.accounttype_id) {
                $DataRow.AccountTypeAsString = $AccountTypes.description
                break
            }
        }

        # Set new table columns values for AD parameters
        $DataRow.SamAccountName = $DataRow.NetID
        $DataRow.UserPrincipalName = $DataRow.NetID + "@housing.wisc.edu"
        $DataRow.Office = $DataRow.DepartmentAsString
        $DataRow.OU = "OU=User Accounts,OU=" + $DataRow.DepartmentAsString + $RootOU
        $DataRow.ProfilePath = $ProfileServerDir + $userNetID
        
        # Single variable for supervisor from the datarow
        $DRmanager = $DataRow.supervisor

        # Get Manager in AD and store that user's NetID as a string in the datarow  
        $DataRow.ManagerNetID = Get-ADUser -f{name -like $DRmanager} -Properties SamAccountName | Select-Object -ExpandProperty SamAccountName

        # Changes space with an underscore from Office365 email creation
        $DataRow.last_name = $DataRow.last_name -replace ' ', '_'
        $DataRow.first_name = $DataRow.first_name -replace ' ', '_'
    }
}
#endRegion

# **********************************************************************************
#  Duplicate account handling 
# **********************************************************************************
#region Duplicate User Handling
# Checks for Duplicate Users; If found, adjusts relevent properties and permissions to treat as an internal move
foreach($DataRow in $newCreationRequests){

    # Store NetID in local variable
    $userNetID = $DataRow.netid

    # Search AD for a user with the submitted NetID
    $existingADUser = Get-ADUser -Filter{SamAccountName -eq $userNetID} -ErrorAction SilentlyContinue

    # If there is a result change the permissions and office information of the user
    if ($existingADUser) {

        # Alert for the duplicate
        Write-Host "Duplicate user account found: " $DataRow.SamAccountName -ForegroundColor Blue
        Write-Log -LogString "[WARNING] Duplicate user account found: $($DataRow.SamAccountName)" -LogFile $LogFile

        # Gets table data for new employee info in AD
        $newOffice = $DataRow.Office
        $userNewOU ='OU=User Accounts,OU=' + $newOffice + ',OU=Departments,DC=housing,DC=wisc,DC=edu'
        $newPermissionReference = Get-ADUser -Identity $DataRow.copyfrom
        $newTitle = $DataRow.title
        $managerChange = $DataRow.ManagerNetID
        $dn = Get-ADUser -Identity $netID -Properties DistinguishedName | Select-Object -ExpandProperty DistinguishedName
        
        # Clears old permissions and copies permissions from the copy-from staff
        Get-ADUser $existingADUser -Properties memberof | Select-Object -ExpandProperty memberof | Remove-ADGroupMember -Members $existingADUser -Confirm:$false -Verbose:$false -ErrorAction SilentlyContinue
        Get-ADUser $newPermissionReference -Properties memberof | Select-Object -ExpandProperty memberof | Add-ADGroupMember -Members $existingADUser -PassThru -Confirm:$false -Verbose:$false -ErrorAction SilentlyContinue

        # Update User AD information
        Set-ADUser $existingADUser -Office $newOffice -Manager $managerChange -Description $newTitle -Confirm:$false -Verbose:$false

        # Move User Account to new OU
        Move-ADObject -Identity $dn -TargetPath $userNewOU

        # Creates the string for the Primary Group name
        $primGroup = "Primary - " + $DataRow.office

        # Add user to primary group; Try-Catch used to surpress expected errors
        try{Add-ADGroupMember -Identity $primGroup -Members $existingADUser -PassThru -Verbose:$false -ErrorAction SilentlyContinue} catch{}

        # Set New Primary Group for User
        $group = get-adgroup $primGroup

        # Check if user needs a service account
        foreach($position in Get-Content $serviceAccountPositionsFile){
            if($DataRow.copyfrom -eq $position){

                # Convert datatable entries to single variables
                $firstName = $DataRow.first_name
                $lastName = $DataRow.last_name
                $description = $DataRow.description
                $netID = $DataRow.netid
                $managerNetID = $DataRow.ManagerNetID
                ##TEST -> Change Function names when completed
                # Imported custom function to make service account
                NewStudentAccount -firstName "$firstName" -lastName "$lastName" -description "$description" -netID $netID -managerNetID "$netID"
            }
        }

        # variable to hold SID
        $groupSid = $group.SID

        # Update SQL Database to change this user to Created
        Invoke-Sqlcmd -ServerInstance $sqlServer -Database $sqlDataBase -Query "UPDATE accounts SET status_id = 5 WHERE (status_id = 3 AND accounts.netid = '$userNetID')"

        # Update Datatable status_id
        $DataRow.status_id = 5

        # Status Update
        Write-Host "User postion moved and removed from new request list" -ForegroundColor Blue

        # Add User to 'Position Changed' collection for email purposes
        $InternalMoveUsers = @()
        $InternalMoveUsers += $existingADUser
    }
}
#endRegion

# **********************************************************************************
#  Account creation 
# **********************************************************************************
#region User Account Creation Process
Write-Host "Starting user account creation process" -ForegroundColor "Green"

# Creates a new table for non-duplicate requests 
$creationRequests = @()
foreach($DataRow in $newCreationRequests){
    if($DataRow.status_id -eq 3){
        $creationRequests += $DataRow
    }
}

foreach ($DataRow in $creationRequests) {

    # Service account variable
    $needsSA = $true

    # Get User's Manager
    $userManager = Get-ADUser -Identity $DataRow.ManagerNetID

    # Add manager NetID to a container for the emailing process
    $newUserManagers = @()
    $newUserManagers += $userManager | Select-Object -ExpandProperty SamAccountName 
    
    # Creates User's name in a format for AD
    $userName = $DataRow.first_name + " " + $DataRow.last_name

    # Gets User Description
    $userDesc = $DataRow.title

    # Stores the user's NetID
    $netID = $DataRow.SamAccountName
    
    # Assembles Housing email address
    $housingEmail = $DataRow.first_name + "." + $DataRow.last_name + ''
    
    # Assemble user's msolUID
    $user_mssolUID = $netID

    # Create the user with data from the table row
    New-ADUser -Name $userName -AccountPassword $DefaultPassword -ChangePasswordAtLogon $true -Department $DefaultDepartment -Description $userDesc -DisplayName $DataRow.DisplayName -EmailAddress $housingEmail -EmployeeID $netID -Enabled $true -GivenName $userName -Manager $userManager -Office $DataRow.Office -OfficePhone $DataRow.phone_number -Path $DataRow.OU -PostalCode "53703" -ProfilePath $ProfileServerDir -SamAccountName $netID -State "WI" -Surname $DataRow.last_name -Title $DataRow.title -UserPrincipalName $DataRow.UserPrincipalName -ErrorAction SilentlyContinue
    
    # Add values to custom AD fields 
    Set-ADUser -Identity $netID -Add @{msolUID = $user_mssolUID}
    Set-ADUser -Identity $netID -Add @{EmployeeType = $DataRow.AccountTypeAsString}

    # Additional AD property changes
    Set-ADUser -Identity $netID -Add @{msolUID = $user_mssolUID}

    # Set Variables for Office 365 Creation
    $userFirst = $DataRow.first_name
    $userLast = $DataRow.last_name
    $managerNetID = $DataRow.ManagerNetID
   
    # Check if user needs a service account
    foreach($position in Get-Content $serviceAccountPositionsFile){
        if($DataRow.copyfrom -eq $position){
            ##TEST -> Change Function names when completed
            # Imported custom function to make service account
            NewStudentAccount -firstName "$userFirst" -lastName "$userLast" -description "$userDesc" -netID "$netID" -managerNetID "$managerNetID"
            
            # Change MSSolUID to reflect the service account creation
            $user_mssolUID = $user_mssolUID + "_housing"
            Set-ADUser -Identity $netID -Replace @{msolUID = $user_mssolUID}
            
            # Boolean variable so email creation can happen out of this loop
            $needsSA = $true
        }
    }
    
    if($needsSA -eq $false) {

        # Create Alias for staff if a service account is not needed
        NewStaffAccount -firstName "$userFirst" -lastName "$userLast" -netID "$netID" -managerNetID "$managerNetID"
    }

    # Set proxy addresses in AD
    Set-ADUser -Identity $netID -Add @{proxyAddresses = "SMTP:$housingEmail" }
    Set-ADUser -Identity $netID -Add @{proxyAddresses = "sip:$housingEmail" }

    # Get reference user's name
    $copyFromUser = Get-ADUser -Identity $DataRow.copyfrom -Properties Description | Select-Object -ExpandProperty Description

    # Copy Permissions from reference staff
    Get-ADUser -Identity $DataRow.copyfrom -Properties memberof | Select-Object -ExpandProperty memberof | Add-ADGroupMember -Members $netID -PassThru -ErrorAction SilentlyContinue | Select-Object -Property SamAccountName

    # Status Updates
    Write-Host "Copying permissions from: $copyFromUser to: $userName" -ForegroundColor "Green"
    Write-Log -LogString "[INFO] Copying group memberships from: $copyFromUser to: $userName" -LogFile $LogFile

    # Creates the string for the Primary Group name
    $primGroup = "Primary - " + $DataRow.office

    # Add user to primary group
    Add-ADGroupMember -Identity $primGroup -Members $netID -PassThru -ErrorAction SilentlyContinue

    # Set New Primary Group for User
    $group = get-adgroup $primGroup

    # variable to hold SID
    $groupSid = $group.sid

    # Gets AD group SID to change primary group
    [int]$groupID = $groupSid.Value.Substring($groupSid.Value.LastIndexOf("-")+1)

    # Set users primary group
    Set-ADUser -identity $netID -Replace @{primaryGroupID="$GroupID"}

    # Update SQL Database to change this user to Created
    Invoke-Sqlcmd -ServerInstance $sqlServer -Database $sqlDataBase -Query "UPDATE accounts SET status_id = 5 WHERE (status_id = 3 AND accounts.netid = '$netID')"

}
#endRegion

# **********************************************************************************
#  Email Updates for Account Creation
# **********************************************************************************
#region Email Functions

#region Creator Email Function

# Create a table of the requestor users
$userRequestors = @()
foreach($DataRow in $newCreationRequests){
    $userRequestor = $DataRow.entered_by
    $userRequestors += Get-ADUser -Identity $userRequestor -Properties CN, SamAccountName, EmailAddress
}

# Handle duplicate entries of the users
$userRequestors = $userRequestors | Select-Object -Unique

# Container for created users
$requestedUsers = @()

# Iterate through the requestor staff to find their requests
foreach($accountRequestor in $userRequestors){

    # Get requestor NetID
    $requestorNetID = $accountRequestor | Select-Object -ExpandProperty SamAccountName

    # Get requestor's name
    $requestorName = Get-ADUser -Identity $requestorNetID -Properties CN | Select-Object -ExpandProperty CN

    # Get requestor email
    $requestorEmail = Get-ADUser -Identity $requestorNetID -Properties EmailAddress | Select-Object -ExpandProperty EmailAddress

    # Create a table of users from the same user
    foreach($DataRow in $newCreationRequests){

        # Store variables from datatables
        $netID = $DataRow.NetID
        $enteredBy = $DataRow.entered_by

        # Matches the users to the requestor
        if($requestorNetID -eq $enteredBy){

            # Store in a temp variable then add to requested user collection
            $requestedUser = Get-ADUser -Identity $netID -Properties CN, EmailAddress
            $requestedUsers += $requestedUser
        }
    }

    # Chooses email templates based on how many account requests
    if ($requestedUsers.Length -eq 1) {

        # Create additional accesses table
        foreach($DataRow in $newCreationRequests){
                
            # Replace 'user' with requestor name
            $emailBodySUR = $emailBodySUR.Replace("[user]", $requestorName)

            # Replace 'reqUser' with new staff
            $emailBodySUR = $emailBodySUR.Replace("[reqUser]", $DataRow.CN)

            #Send Email
            Send-MailMessage -From $systemsEmailAddress -to $requestorEmail -Subject $emailSubject -Body $emailBodySUR -SmtpServer $emailServer -port $emailPort -BodyAsHtml -DeliveryNotificationOption OnFailure
        }
    }

    # Multiple user requests
    ElseIf ($requestedUsers.Length -gt 1) {

        # Temp user variable
        $newUsersTemp = ""

        foreach($DataRow in $newCreationRequests){

            if($DataRow.entered_by -eq $requestorNetID){

                # Adds name to array with a newline character for email formating
                $newUsersTemp += $DataRow.CN + ",`n"
            }
        }

        # Replace 'user' with requestor name
        $emailBodyMUR = $emailBodyMUR.Replace("[user]", $requestorName)

        # Replace 'reqUser' with new staff
        $emailBodyMUR = $emailBodyMUR.Replace("[reqUsers]", $newUsersTemp)

        #Send Email
        Send-MailMessage -From $systemsEmailAddress -to $requestorEmail -Subject $emailSubject -Body $emailBodyMUR -SmtpServer $emailServer -port $emailPort -BodyAsHtml -DeliveryNotificationOption OnFailure

    }

    # Error Handling
    else{
        Write-Output $requestedUsers.Length
        #Write-Log "Unknown email error, email notification not sent to requestor"
        Write-Host "Unknown email error, email notification not sent to requestor"
    }
}
#endregion

# Sends email to listed supervisor if it different than the submitter
#region Manager Email

# Container for managed users
$supervisedUsers = @()

# Container for supervisors
$supervisorsToEmail = @()

foreach($DataRow in $newCreationRequests){

    # Storage variables
    $managerNetID = $DataRow.ManagerNetID
    $enteredBy = $DataRow.entered_by

    # Creates a list of supervisors that need to be emailed
    if ($managerNetID -ne $enteredBy ) {

        $enteredBy = $managerNetID
        $supervisorToEmail = Get-ADUser -Identity $enteredBy -Properties SamAccountName, EmailAddress
        $supervisorsToEmail += $supervisorToEmail
    }
}

# Handle duplicate entries of the users
$supervisorsToEmail = $supervisorsToEmail | Select-Object -Unique

# Handles if there are no differences between supervisor and requestor
if ($null -ne $supervisorsToEmail) {
    
    # Gathers the new users under this supervisor
    foreach($supervisor in $supervisorsToEmail){

        # Create a table of users from the same user
        foreach($DataRow in $newCreationRequests){

            # Get supervisor NetID from table
            $supervisorNetID = $DataRow.ManagerNetID

            # Store AD object SamAccount
            $tempSupervisorSam = Get-ADUser -Identity $supervisor -Properties SamAccountName | Select-Object -ExpandProperty SamAccountName

            # Matches the users to the requestor
            if($supervisorNetID -eq $tempSupervisorSam){

                # Get users
                $supervisedUsers += Get-ADUser -Identity $DataRow.netid

                # Assemble name string for email template
                $supervisedUsersString = ""
                $supervisedUsersString += $DataRow.CN + ",`n"
            }
        }
        
        # Get Supervisor Email
        $supervisorEmail = Get-ADUser -Identity $supervisor -Properties EmailAddress | Select-Object -ExpandProperty EmailAddress

        # Storage variable
        $supervisorName = $DataRow.supervisor

        if($supervisedUsers.Length -eq 1){

            # Replace 'user' with requestor name
            $emailBodyMSU = $emailBodyMSU.Replace("[user]", $supervisorName)

            # Replace 'subordinateUser' with new staff
            $emailBodyMSU = $emailBodyMSU.Replace("[subordinateUser]", $supervisedUsersString)

            #Send Email
            Send-MailMessage -From $systemsEmailAddress -to $supervisorEmail -Subject $emailSubject -Body $emailBodyMSU -SmtpServer $emailServer -port $emailPort -BodyAsHtml -DeliveryNotificationOption OnFailure
        }

        ElseIf($supervisedUsers.Length -gt 1){

            Write-Output $supervisorEmail
            # Replace 'user' with requestor name
            $emailBodyMMU = $emailBodyMMU.Replace("[user]", $supervisorName)

            # Replace 'subordinateUser' with new staff
            $emailBodyMMU = $emailBodyMMU.Replace("[subordinateUsers]", $supervisedUsersString)

            #Send Email
            Send-MailMessage -From $systemsEmailAddress -to $supervisorEmail -Subject $emailSubject -Body $emailBodyMMU -SmtpServer $emailServer -port $emailPort -BodyAsHtml -DeliveryNotificationOption OnFailure
        }

        # Error Handling
        else {
            #Write-Log "Unknown email error, email notification not sent to supervisor"
            Write-Host "Unknown email error, email notification not sent to supervisor"
        }
    }
}

#endRegion manager emailing


#endRegion Email functions

# **********************************************************************************
#  Account Deletion
# **********************************************************************************
#region

# If there are new deletion requests
If ($null -ne $newDeletionRequests) {

    # Iterator
    $delNum = ($newDeletionRequests.Length)
    
    # Write out a status message
    If ($delNum -eq 1) {
        Write-Host "There is 1 account to delete" -ForegroundColor "Green"
        Write-Log -LogString "[INFO] There is one account to delete" -LogFile $LogFile

    }
    Else {
        Write-Host "There are $delNum accounts to delete" -ForegroundColor "Green"
        Write-Log -LogString "[INFO] There are $delNum accounts to delete" -LogFile $LogFile
    }   

    Write-Host "Processing deleted Users" -ForegroundColor "Yellow"
    foreach($DataRow in $newDeletionRequests){
        
        # Storage variables
        $netID = $DataRow.NetID
        $dn = Get-ADUser -Identity $netID -Properties DistinguishedName | Select-Object -ExpandProperty DistinguishedName

        # Find and disable user account
        $user = Get-ADUser -Identity $netID -Properties Description | Set-ADUser -Enabled $false -Description "Disabled by script $time" | out-null
        Start-Sleep -s 2 

        # Move account to deletion holding OU 
        Move-ADObject -Identity $dn -TargetPath $disableOU | out-null

        # Update SQL Database to change this deletion holding
        Invoke-Sqlcmd -ServerInstance $sqlServer -Database $sqlDataBase -Query "UPDATE accounts SET status_id = 9 WHERE (status_id = 6 AND accounts.netid = '$netID')"
    }
    Write-Host "User accounts in disabled and in holding" -ForegroundColor "Green"
}
#endregion

# **********************************************************************************
#   Accounts to be disabled for the summer
# **********************************************************************************
#region

# If there are new deletion requests
If ($null -ne $newSummerDisableRequests) {

    # Iterator
    $disNum = ($newSummerDisableRequests.Length)
    
    # Write out a status message
    If ($disNum -eq 1) {
        Write-Host "There is 1 account to disable" -ForegroundColor "Green"
        Write-Log -LogString "[INFO] There is one account to disable" -LogFile $LogFile

    }
    Else {
        Write-Host "There are $disNum accounts to disable" -ForegroundColor "Green"
        Write-Log -LogString "[INFO] There are $disNum accounts to disable" -LogFile $LogFile
    }   

    Write-Host "Processing users leaving for the summer" -ForegroundColor "Yellow"

    foreach($DataRow in $newSummerDisableRequests){

        # Storage variables
        $netID = $DataRow.NetID

        if(($user = Get-ADUser -Identity $netID)){
            
            # Gets user's department
            $userOU = Get-ADUser $user -Properties CanonicalName | Select-Object -ExpandProperty CanonicalName

            # Find user account and store departmental OU for account moving
            $user = Get-ADUser -Identity $netID -Properties Description, CanonicalName | Select-Object -ExpandProperty CanonicalName
            $parentOU, $userSubOU, $userDepartment, $userAcctLocation, $userFullName = $userOU.split('/')
           
            # Storage variables
            $userOU = 'OU=' + $userDepartment + ''
            $deptDisableOU = $disableSummer + $userOU
            $dn = Get-ADUser -Identity $netID -Properties DistinguishedName | Select-Object -ExpandProperty DistinguishedName

            # Move account to deletion holding OU 
            Move-ADObject -Identity $dn -TargetPath $deptDisableOU | out-null

            # Update SQL Database to change this deletion holding
            Invoke-Sqlcmd -ServerInstance $sqlServer -Database $sqlDataBase -Query "UPDATE accounts SET status_id = 11 WHERE (status_id = 8 AND accounts.netid = '$netID')"
        
            # Disable the user
            Write-Host "`nDisabling User" -ForegroundColor Blue
            Get-ADUser -Identity $netID -Properties Description | Set-ADUser -Enabled $false -Description "Disabled for Summer on $time"
        }
    }
    Write-Host "User accounts moved to summer disable OUs" -ForegroundColor "Green"
}

#endregion

# **********************************************************************************
#   Check for accounts to be re-enabled after the summer
# **********************************************************************************
#region

# If there are new deletion requests
If ($null -ne $newSummerEnableRequests) {

    # Iterator
    $enableNum = $newSummerEnableRequests.Length

    # Give a status update
    Write-Host "Checking for accounts to be Re-Enabled for summer..." -ForegroundColor "Green"
    Write-Log -LogString "[INFO] Checking for accounts to be Re-Enabled for summer" -LogFile $LogFile

    # Write out a status message
    If ($enableNum -eq 1) {
        Write-Host "There is 1 account to Re-Enable for summer" -ForegroundColor "Green"
        Write-Log -LogString "[INFO] There is one account to Re-Enable for summer" -LogFile $LogFile

    }
    Else {
        Write-Host "There are $reEnableNum accounts to Re-Enable for summer" -ForegroundColor "Green"
        Write-Log -LogString "[INFO] There are $reEnableNum accounts to Re-Enable for summer" -LogFile $LogFile
    }

    foreach($DataRow in $newSummerEnableRequests){
         
        # Get NetID and see if user still exists
        $netID = $DataRow.netid        
        if(($user = Get-ADUser -Identity $netID)){

            # Update message
            Write-Host "Account found in Active Directory: " $chkRow.SamAccountName -ForegroundColor "Green"
            Write-Log -LogString "[INFO] Account found in Active Directory: $($chkRow.SamAccountName)" -LogFile $LogFile

            # Finds the users OU path
            $userOU = Get-ADUser -Identity $netID -Properties CanonicalName | Select-Object -ExpandProperty CanonicalName
            $parentOU, $userSubOU, $userDepartment, $userAcctLocation, $userFullName = $userOU.split('/')
            $reEnablePath = 'OU=User Accounts,OU=' + $userDepartment + ''

            # Storage variable
            $dn = Get-ADUser -Identity $netID -Properties DistinguishedName | Select-Object -ExpandProperty DistinguishedName
            $description = $DataRow.title

            # Move user
            Move-ADObject -Identity $dn -TargetPath $reEnablePath

            # Reenable the user
            Write-Host "`nRe-Enabling User" -ForegroundColor Blue
            Get-ADUser -Identity $netID -Properties Description | Set-ADUser -Enabled $true -Description $description | out-null

            # Update database
            Invoke-Sqlcmd -ServerInstance $sqlServer -Database $sqlDataBase -Query "UPDATE accounts SET status_id = 5 WHERE (status_id = 10 AND accounts.netid = '$netID')"
        }
    }

}
#endregion

# **********************************************************************************
#   Check Deletion Holding Cell
# **********************************************************************************
#region Description
# This function is to bridge the Deletion function with the User Deletion Request on
# the user account website. This function sets accounts to expire then sets them to
# delete when the time is in the past. This function is needed to prevent premature
# deletions.
#endregion
# **********************************************************************************
#region

# Process requests
foreach($DataRow in $EoARequests){

    # Get user
    $netID = $DataRow.netid
    $user = Get-ADUser -Identity $netID | Out-Null

    # If the user exists
    If($user){

        # Holding variable
        $EoADate = $DataRow.EndDate

        # Set account expiration
        Set-ADAccountExpiration -Identity $user -DateTime $EoADate

        # Check if accounts have been disabled for a month
        If($time -gt $EoADate){

            # Move user and update database
            Move-ADObject $user -TargetPath $deletionOU
            Invoke-Sqlcmd -ServerInstance $sqlServer -Database $sqlDataBase -Query "UPDATE accounts SET status_id = 7 WHERE (status_id = 9 AND accounts.netid = '$netID')"

        }

    }
}

#endregion