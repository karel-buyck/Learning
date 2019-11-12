Get-ADUser -filter * -Properties * | Select name, userprincipalname, department, title, Enabled | Export-csv c:\temp\aduserinventory.csv

Test
