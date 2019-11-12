$Users = Import-CSV C:\Users\Karel.ESC\Downloads\aduserimport.csv

foreach ($User in $Users)
{
    $ADUserParams = @{
        SamAccountName   = $User.SamAccountName
        UserPrincipalName    = $User.UserPrincipalName
        DisplayName         = $User.DisplayName
        Surname             = $User.Surname
        GivenName = $User.GivenName
        Name            = $User.Name
        EmailAddress    = $User.EmailAddress
    }
    Set-ADUser @ADUserParams
}

Test