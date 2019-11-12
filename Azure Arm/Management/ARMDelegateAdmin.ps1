Set-Location -Path "\\files\infrastructuur\PowerShell\AzureDelegatedAccess"

Connect-AzAccount 

New-AzDeployment -Name "ESCDelegatedAdmin" `
-Location "West Europe" `
-TemplateFile "\\files\infrastructuur\PowerShell\AzureDelegatedAccess\delegatedResourceManagement.json" `
-TemplateParameterFile "\\files\infrastructuur\PowerShell\AzureDelegatedAccess\armdelegateadminparameter.json" `
-Verbose

DisConnect-AzAccount

