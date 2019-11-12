# Import Active Directory Module 
try 
{ 
    import-module activedirectory  
} 
catch 
{ 
    write-Output "Error while importing Active Directly module, please install AD Powershell modules first before running this script." 
} 
 
# Get AD Users and export results to CSV 
Get-ADUser -Filter '*' |  select SamAccountName, UserPrincipalName, DisplayName, GivenName, Surname, Name, EmailAddress | Export-Csv c:\temp\ADUsers.csv