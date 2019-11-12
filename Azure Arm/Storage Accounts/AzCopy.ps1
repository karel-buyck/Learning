#tenant id 17b35a1d-057c-4ac5-a15a-08758f7a7064 = Infra Playground
#subscription id  a6e10cb0-79e5-4b68-af13-d17fc5f7505a = Infra Playground

######################################################################################################################################

#Er is geen manier om een Azure Storage Blob aan te spreken via Windows Explorer, daarom vertrekken we van een UNC pad.
#De source heeft een UNC pad en wordt gebruikt om de noodzakelijk files op te laden in de Azure Storage Account
#we gebruiken onderstaande script om de verbinding te maken met Azure Files
#Een developer zal dit moeten uitvoeren op zijn lokale computer om de files op te laden in Azure Files
#de gebruiker in kwestie kan steeds Azure Storage Explorer gebruiken, maar we gaan ervan uit dat dit niet zo is.


cmdkey /add:demooutlookaddin.file.core.windows.net /user:Azure\demooutlookaddin /pass:3UYF20E+mhO5nPXeaf14YTm1xZr+71Zh5lcetiFCkcFpF1oJhTvkqkS8kQ11AEeHEWxW0s8KYN1x0H0ZpxVlfg==
net use \\demooutlookaddin.file.core.windows.net\outlookaddin-source /persistent:No

#security gebeurd momenteel op vlak van de token "pass:..."
#Dit dient nadien te gebeuren met Azure AD Authentication (voor Delaware personel, AB2B of Customer management)

sleep -Seconds 10

#PopUp voor user die de source folder selecteert

Add-Type -AssemblyName System.Windows.Forms
$FolderBrowser = New-Object System.Windows.Forms.FolderBrowserDialog
[void]$FolderBrowser.ShowDialog()
$FolderBrowser.SelectedPath

#na het uitvoeren van het script kan de gebruiker via Windows Explorer de nodige files oploaden


$source=$FolderBrowser.SelectedPath 
$target="\\demooutlookaddin.file.core.windows.net\outlookaddin-source\delaware FSM Connect"

Copy-Item $source -Destination $target -Recurse -force


######################################################################################################################################

#https://www.thomasmaurer.ch/2019/05/how-to-install-azcopy-for-azure-storage/

#Download AzCopy
Invoke-WebRequest -Uri "https://aka.ms/downloadazcopy-v10-windows" -OutFile AzCopy.zip -UseBasicParsing
 
#Curl.exe option (Windows 10 Spring 2018 Update (or later))
curl.exe -L -o AzCopy.zip https://aka.ms/downloadazcopy-v10-windows
 
#Expand Archive
Expand-Archive ./AzCopy.zip ./AzCopy -Force

#Create C:\Azure
New-Item -ItemType directory -Path C:\Azure\AzCopy -force

#Move AzCopy to the destination you want to store it
Get-ChildItem ./AzCopy/*/azcopy.exe | Move-Item -Destination "C:\Azure\AzCopy\AzCopy.exe" -Force
 
#Add your AzCopy path to the Windows environment PATH (C:\Users\thmaure\AzCopy in this example), e.g., using PowerShell:
$userenv = [System.Environment]::GetEnvironmentVariable("Path", "User")
[System.Environment]::SetEnvironmentVariable("PATH", $userenv + ";C:\Azure\AzCopy", "User")

######################################################################################################################################

Set-PSRepository -Name 'PSGallery' -InstallationPolicy Trusted

function Load-Module ($m) {

    # If module is imported say that and do nothing
    if (Get-Module | Where-Object {$_.Name -eq $m}) {
        write-host "Module $m is already imported."
    }
    else {

        # If module is not imported, but available on disk then import
        if (Get-Module -ListAvailable | Where-Object {$_.Name -eq $m}) {
            Import-Module $m -Verbose
        }
        else {

            # If module is not imported, not available on disk, but is in online gallery then install and import
            if (Find-Module -Name $m | Where-Object {$_.Name -eq $m}) {
                Install-Module -Name $m -Force -Verbose -Scope CurrentUser
                Import-Module $m -Verbose
            }
            else {

                # If module is not imported, not available and not in online gallery then abort
                write-host "Module $m not imported, not available and not in online gallery, exiting."
                EXIT 1
            }
        }
    }
}

Load-Module "Az"

#Install-Module -Name Az -AllowClobber -Scope AllUsers
#uninstall-module az -AllVersions
#Import-Module Az

######################################################################################################################################

Connect-AzAccount -Tenant "17b35a1d-057c-4ac5-a15a-08758f7a7064" -Subscription "a6e10cb0-79e5-4b68-af13-d17fc5f7505a"

cd C:\Azure\AzCopy

#Momenteel nog een issue met de log location
#.\azcopy env set AZCOPY_LOG_LOCATION=c:\temp
#$env:AZCOPY_LOG_LOCATION=c:\temp

#sync Azure Files met Azure Storage Container "outlookaddindev"
.\azcopy sync "\\demooutlookaddin.file.core.windows.net\outlookaddin-source\delaware FSM Connect" "https://demooutlookaddin.blob.core.windows.net/outlookaddindev/delaware FSM Connect?sv=2019-02-02&ss=bfqt&srt=sco&sp=rwdlacup&se=2020-11-07T20:25:13Z&st=2019-11-07T12:25:13Z&spr=https&sig=AMhR3lc2DH4%2B0EKsuXNag5z5lbSl%2FplZ%2FjVbj2Bu1%2FM%3D"--recursive=true --delete-destination

#sync Azure Storage Container "outlookaddindev" met Azure Storage Container "outlookaddintest"
#accesskeys moet nog beter gedefinieerd worden

#.\azcopy sync "https://demooutlookaddin.blob.core.windows.net/outlookaddindev/delaware FSM Connect?sv=2019-02-02&ss=bfqt&srt=sco&sp=rwdlacup&#se=2020-11-07T20:25:13Z&st=2019-11-07T12:25:13Z&spr=https&sig=AMhR3lc2DH4%2B0EKsuXNag5z5lbSl%2FplZ%2FjVbj2Bu1%2FM%3D" "https://#demooutlookaddin.blob.core.windows.net/outlookaddintest/delaware FSM Connect?sv=2019-02-02&ss=bfqt&srt=sco&sp=rwdlacup&se=2020-11-07T20:25:13Z&#st=2019-11-07T12:25:13Z&spr=https&sig=AMhR3lc2DH4%2B0EKsuXNag5z5lbSl%2FplZ%2FjVbj2Bu1%2FM%3D"--recursive=true --delete-destination  

#sync Azure Storage Container "outlookaddintest" met Azure Storage Container "outlookaddinlive"
#.\azcopy sync "https://demooutlookaddin.blob.core.windows.net/outlookaddintest/delaware FSM Connect?sv=2019-02-02&ss=bfqt&srt=sco&sp=rwdlacup&#se=2020-11-07T20:25:13Z&st=2019-11-07T12:25:13Z&spr=https&sig=AMhR3lc2DH4%2B0EKsuXNag5z5lbSl%2FplZ%2FjVbj2Bu1%2FM%3D" "https://#demooutlookaddin.blob.core.windows.net/outlookaddinlive/delaware FSM Connect?sv=2019-02-02&ss=bfqt&srt=sco&sp=rwdlacup&se=2020-11-07T20:25:13Z&#st=2019-11-07T12:25:13Z&spr=https&sig=AMhR3lc2DH4%2B0EKsuXNag5z5lbSl%2FplZ%2FjVbj2Bu1%2FM%3D"--recursive=true  --delete-destination 

net use \\demooutlookaddin.file.core.windows.net\outlookaddin-source /del