New-AzResourceGroupDeployment -ResourceGroupName DeleteMe -TemplateFile  "C:\Users\buyckka\OneDrive - Delaware\Documents\Learning\Azure Arm\Storage Accounts\Deployment\template.json"

##########################################

New-AzResourceGroupDeployment -Name ExampleDeployment -ResourceGroupName DeleteMe -TemplateFile "C:\Users\buyckka\OneDrive - Delaware\Documents\Learning\Azure Arm\Storage Accounts\Deployment\template.json" -TemplateParameterFile "C:\Users\buyckka\OneDrive - Delaware\Documents\Learning\Azure Arm\Storage Accounts\Deployment\parameters.json"