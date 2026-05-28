# ================================
# ACTIVE DIRECTORY USER CREATOR
# WITH PASSWORD EXPIRATION EXCEL REPORT
# ================================

Import-Module ActiveDirectory

# Install ImportExcel once if needed:
# Install-Module ImportExcel -Scope CurrentUser

Import-Module ImportExcel

# ---------- CONFIGURATION ---------- #

$PASSWORD_FOR_USERS = "Testlabs1!"
$NUMBER_OF_ACCOUNTS_TO_CREATE = 50

$PASSWORD_CHANGE_DAYS = 45

$START_DATE_MIN = Get-Date "2026-01-01"
$START_DATE_MAX = Get-Date "2026-05-01"

$EXCEL_PATH = "C:\PasswordChangeReport.xlsx"

$OU_PATH = "OU=_EMPLOYEES," + ([ADSI]"").distinguishedName

# ---------- RANDOM NAME GENERATOR ---------- #

Function Generate-Random-Name {
    $consonants = @('b','c','d','f','g','h','j','k','l','m','n','p','q','r','s','t','v','w','x','z')
    $vowels = @('a','e','i','o','u','y')
    $nameLength = Get-Random -Minimum 3 -Maximum 7
    $name = ""

    for ($count = 0; $count -lt $nameLength; $count++) {
        if ($count % 2 -eq 0) {
            $name += $consonants[(Get-Random -Minimum 0 -Maximum $consonants.Count)]
        }
        else {
            $name += $vowels[(Get-Random -Minimum 0 -Maximum $vowels.Count)]
        }
    }

    return $name
}

# ---------- RANDOM START DATE ---------- #

Function Get-Random-StartDate {
    $daysBetween = ($START_DATE_MAX - $START_DATE_MIN).Days

    return $START_DATE_MIN.AddDays(
        (Get-Random -Minimum 0 -Maximum $daysBetween)
    )
}

# ---------- CREATE USERS ---------- #

Write-Host ""
Write-Host "Creating Active Directory Users..." -ForegroundColor Yellow
Write-Host ""

$count = 1

while ($count -le $NUMBER_OF_ACCOUNTS_TO_CREATE) {

    $firstName = Generate-Random-Name
    $lastName = Generate-Random-Name
    $username = "$firstName.$lastName"

    $password = ConvertTo-SecureString $PASSWORD_FOR_USERS -AsPlainText -Force

    $startDate = Get-Random-StartDate
    $startDateText = $startDate.ToString("yyyy-MM-dd")

    Write-Host "Creating User: $username | Start Date: $startDateText" -ForegroundColor Cyan

    New-ADUser `
        -Name $username `
        -GivenName $firstName `
        -Surname $lastName `
        -DisplayName $username `
        -SamAccountName $username `
        -UserPrincipalName "$username@testlab.com" `
        -EmployeeID $username `
        -Description "StartDate=$startDateText" `
        -AccountPassword $password `
        -Enabled $true `
        -PasswordNeverExpires $false `
        -ChangePasswordAtLogon $false `
        -Path $OU_PATH

    $count++
}

Write-Host ""
Write-Host "User creation complete." -ForegroundColor Green
Write-Host ""

# ---------- PASSWORD EXPIRATION CHECK ---------- #

Write-Host "Scanning users for password expiration..." -ForegroundColor Yellow

$today = Get-Date

$users = Get-ADUser `
    -SearchBase $OU_PATH `
    -Filter * `
    -Properties Description, PasswordNeverExpires

$report = foreach ($user in $users) {

    if ($user.Description -match "StartDate=(\d{4}-\d{2}-\d{2})") {

        $startDate = Get-Date $matches[1]
        $passwordDueDate = $startDate.AddDays($PASSWORD_CHANGE_DAYS)
        $daysRemaining = ($passwordDueDate - $today).Days

        if ($today -ge $passwordDueDate) {
            $status = "PASSWORD CHANGE REQUIRED"

            Set-ADUser `
                $user.SamAccountName `
                -ChangePasswordAtLogon $true
        }
        else {
            $status = "Not Due Yet"
        }

        [PSCustomObject]@{
            Username = $user.SamAccountName
            StartDate = $startDate.ToString("yyyy-MM-dd")
            PasswordChangeDue = $passwordDueDate.ToString("yyyy-MM-dd")
            DaysRemaining = $daysRemaining
            Status = $status
        }
    }
}

# ---------- DISPLAY DETAILED TABLE ---------- #

Write-Host ""
Write-Host "Password Expiration Report" -ForegroundColor Yellow

$report |
    Sort-Object PasswordChangeDue |
    Format-Table Username, StartDate, PasswordChangeDue, DaysRemaining, Status -AutoSize

# ---------- EXPORT TO EXCEL ---------- #

$report |
    Sort-Object PasswordChangeDue |
    Export-Excel `
        -Path $EXCEL_PATH `
        -WorksheetName "Password Report" `
        -AutoSize `
        -BoldTopRow `
        -FreezeTopRow `
        -TableName "PasswordChangeReport" `
        -ClearSheet

Write-Host ""
Write-Host "Excel password change report saved to $EXCEL_PATH" -ForegroundColor Green