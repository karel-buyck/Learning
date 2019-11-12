<#

    .SYNOPSIS
    Send you email with overall statistics about computer objects in your environment and break down per O.S Version and service pack



    #--------------
    # Script Info
    #--------------

    # Script Name                :         Get-ADComputerVersionReport
    # Script Version             :         1.0
    # Author                     :         Ammar Hasayen
    # Blog                       :         http://ammarhasayen.wordpress.com
    

     .DESCRIPTION


        - The script will send you a nice email with a summary information about all computers it detects in your AD
        - By default the script will search all your AD. You can adjust this behaviour by uncommenting this line #-SearchBase "DC=Contoso,DC=COM" , and scope it as you like
        - The script will generate a csv file at the script running directory with extensive information
        - The email will include two charts, one for Server versions and one for workstation versions.
        - No switches are needed to run the script. Download and run !

        --------------
         Script Requirement
        --------------
        - The script needs only an account with read permissions in your AD
        - The script assumes that the machine from which it is run, has the Active Directory PowerShell Module


    .EXAMPLE
    Run the script, no switches needed
    .\Get-ADComputerVersionReport.PS1
#>

#--------------
# Script Customization
#--------------

#Set the SMTP Server setting and info
$SMTP_from = "noreply@contoso.com"
$SMTP_to   = "admin@contoso.com"
$SMTP_Host = "smtp.contoso.com"


#.................Screen header Info

write-host 
write-host 
write-host 
write-host "--------------------------" 
write-host "Script Info" -foreground Green
write-host "--------------------------"
write-host
write-host " Script Name :   Get-ADComputerVersionReport"  -foreground Cyan
write-host " Author      :   Ammar Hasayen"  -foreground Cyan
write-host " Version     :   1.0"  -foreground Cyan
write-host " Description :   Send you email with overall statistics about computer objects in your environment and break down per O.S version and service pack"  -foreground Cyan
write-host
write-host "--------------------------" 
write-host "Script Release Notes" -foreground Green
write-host "--------------------------"
write-host
write-host " -  Download the script and run it without any switches. You need only to make sure the machine contains Active Directory PowerShell Module."
write-host " -  Don't Worry :) The script WILL NOT perform any change or SET operations in your environment...It will only use GET and read operations."
write-host " -  Don't forget to open the script and modify your SMTP settings under Script Customization Section !"
write-host " -  ALWAYS CHECK FOR NEWER VERSION @ http://ammarhasayen.wordpress.com."  -foreground Yellow
write-host


sleep 2

write-host "--------------------------" 
write-host "Script Start" -foreground Green
write-host "--------------------------"
Write-Host
write-host





$CurrentScriptLocation = Get-Location
write-host
write-host " 1. Creating Files" -foreground "magenta"
write-host
write-host "....... Script working directory for generating files is: $($CurrentScriptLocation)"
write-host


#.................Preperation Tasks




$date = [DateTime]::Today.AddDays(-90)

[array]$All_Computers                 = @()
[array]$All_Computers_Custom          = @()
[array]$All_Servers                   = @()
[array]$All_Workstations              = @()
[array]$All_Servers_Aggregates        = @()
[array]$All_Workstations_Aggregates   = @()


Try{
New-Item `
 -ItemType file `
   -Path "$($CurrentScriptLocation)" `
   -Name ADComputerReport.htm `
    -Force -ea stop
	$filename = "$($CurrentScriptLocation)" +"\ADComputerReport.htm" 
   

   

New-Item `
 -ItemType file `
   -Path "$($CurrentScriptLocation)" `
   -Name ADComputerCSV.csv `
    -Force -ea stop
	$filename2 = "$($CurrentScriptLocation)" +"\ADComputerCSV.csv"

}

Catch{
    write-host
    write-host " OPS!!!! Access is denied to write the output files (Charts , CSV files) on the current script directory $($CurrentScriptLocation)...Exiting " -foreground "Red"
    write-host
    Exit


    }




#.................Functions

# START :Code Taken from Scripting Guy Blog#
function Get-MyModule 
{ 
Param([string]$name) 
if(-not(Get-Module -name $name)) 
{ 
if(Get-Module -ListAvailable | 
Where-Object { $_.name -eq $name }) 
{ 
Import-Module -Name $name 
$true 
} #end if module available then import 
else { $false } #module not available 
} # end if not module 
else { $true } #module already loaded 
} 
 # END :Code Taken from Scripting Guy Blog#


 
Function sendEmail 
{ param($from,$to,$subject,$smtphost,$htmlFileName) 

$msg =  new-object Net.Mail.MailMessage
$smtp = new-object Net.Mail.SmtpClient($smtphost)
$msg.From = $from
$msg.To.Add($to)
$msg.Subject = $subject
$msg.Body = Get-Content $htmlFileName 
$msg.isBodyhtml = $true 



if($All_ServersCount) {$chart1= "$($CurrentScriptLocation)" +"\Server OS Version.jpeg"
$att1= new-object Net.Mail.Attachment($chart1)
$msg.Attachments.Add($att1)}


if($All_WorkstationsCount){$chart2= "$($CurrentScriptLocation)" +"\Workstation OS Version.jpeg"
$att2 = new-object Net.Mail.Attachment($chart2)
$msg.Attachments.Add($att2)}



$smtp.Send($msg)
if($All_ServersCount){$att1.Dispose()}
if($All_WorkstationsCount){$att2.Dispose()}


} 


    
function _getOU 
{
    param ($dn)
    # Give me Distinguished name and i will give you Parent OU
    $OU = $dn.substring(($dn.IndexOf(',') +1 ))
    $OU

}


FUNCTION  _generateHTML
{
 

        $Output="<html>
    <body>
    <font size=""1"" face=""Arial,sans-serif"">
    <h1 align=""center"">AD Computer Inventory Report</h1>
    <h3 align=""center"">Generated $((Get-Date).ToString())</h3>
    </font>
    <table border=""0"" cellpadding=""3"" style=""font-size:8pt;font-family:Arial,sans-serif"">
    <tr bgcolor=""#8E1275"">
    <th><font color=""#ffffff"">Total Computers</font></th>
    <th colspan=""$($ServerOSArray.count+1)""><font color=""#ffffff"">Servers:</font></th>
    <th colspan=""$($WorkstationOSArray.Count+1)""><font color=""#ffffff"">Workstations</font></th>
    <th><font color=""#ffffff"">Machines with emtpy OS info</font></th></tr>


    <tr bgcolor=""#E46ACB"">"
    # Show Column Headings based on the Exchange versions we have
    $Output+="<th></th>"

    $Output+="<th><font color=""#ffffff"">Total Servers</font></th>"
    $ServerOSArray   | %{$Output+="<th><font color=""#ffffff"">$($_)</font></th>"}

    $Output+="<th><font color=""#ffffff"">Total Wrokstations</font></th>"
    $WorkstationOSArray| %{$Output+="<th><font color=""#ffffff"">$($_)</font></th>"}
    $Output+= "<th></th>"

    $Output+="<tr>"

    $Output+="<tr align=""center"" bgcolor=""#dddddd"">"
    $Output+="<td > $($All_Computers_Count)</td>" 

    $Output+="<th bgcolor=""#FFA500"">$($ALL_ServersCount)</th>"
    $ServerCountArray| %{$Output+="<td>$($_)</td>" }

    $Output+="<th bgcolor=""#FFA500"">$($ALL_WorkstationsCount)</th>"
    $WorkstationCountArray| %{$Output+="<td>$($_)</td>" }
     $Output+= "<th>$($ALL_Unknown)</th>"

    $Output+="</tr><tr><tr></table><br>"




    $output 
}


function _drawServers
{

    

    [void][Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms") 
    [void][Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms.DataVisualization") 
 
    # create chart object 
    $Chart = New-object System.Windows.Forms.DataVisualization.Charting.Chart 
    $Chart.Width          = 800 
    $Chart.Height         = 400
    $Chart.Left           = 900 
    $Chart.Top            = 80
    $CHART.BackColor      = "red" 
 
    # create a chartarea to draw on and add to chart 
    $ChartArea = New-Object System.Windows.Forms.DataVisualization.Charting.ChartArea 
    $Chart.ChartAreas.Add($ChartArea)
    $chartarea.Area3DStyle.Enable3D =$true

    $chartarea.backcolor =  "White"
 
 
    $titlefont      = new-object system.drawing.font("ARIAL",12,[system.drawing.fontstyle]::bold)
    $title          = new-Object System.Windows.Forms.DataVisualization.Charting.title
    $chart.titles.add($title)
    $chart.titles[0].text             = "Server O.S Versions ( $($ALL_Servers.Count) Servers ) "
    $chart.titles[0].font             = $titlefont
    $chart.titles[0].forecolor        = "Red"
 

    #add data to chart 
    
 
    [void]$Chart.Series.Add("Data") 
     $Chart.Series["Data"].Points.DataBindXY( $ServerOSArray, $ServerCountArray)
     $Chart.Series["Data"].IsvalueShownAsLabel=$true 

 
    # add title and axes labels 
 
 
    $ChartArea.AxisX.Title      = "Server O.S Version  " 
    $chartArea.AxisY.Title      = " Count"
    $ChartArea.AxisX.Interval   = 1
 
  
     #Find point with max/min values and change their colour 
    $maxValuePoint = $Chart.Series["Data"].Points.FindMaxByValue() 
    $maxValuePoint.Color = [System.Drawing.Color]::pink 
 
    $minValuePoint = $Chart.Series["Data"].Points.FindMinByValue() 
    $minValuePoint.Color = [System.Drawing.Color]::yellow
 
 
    # change chart area colour 
    $Chart.BackColor = "WHITE"
    $chart.Series["Data"]["DrawingStyle"] = "Emboss"
 
 
    # save chart to file 
    $Chart.SaveImage("$($CurrentScriptLocation)" +"\Server OS Version.jpeg", "JPEG")



}

function _drawWorkstations
{

    

    [void][Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms") 
    [void][Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms.DataVisualization") 
 
    # create chart object 
    $Chart = New-object System.Windows.Forms.DataVisualization.Charting.Chart 
    $Chart.Width          = 800 
    $Chart.Height         = 400
    $Chart.Left           = 900 
    $Chart.Top            = 80
    $CHART.BackColor      = "red" 
 
    # create a chartarea to draw on and add to chart 
    $ChartArea = New-Object System.Windows.Forms.DataVisualization.Charting.ChartArea 
    $Chart.ChartAreas.Add($ChartArea)
    $chartarea.Area3DStyle.Enable3D =$true

    $chartarea.backcolor =  "White"
 
 
    $titlefont      = new-object system.drawing.font("ARIAL",12,[system.drawing.fontstyle]::bold)
    $title          = new-Object System.Windows.Forms.DataVisualization.Charting.title
    $chart.titles.add($title)
    $chart.titles[0].text             = "Server O.S Versions ( $($ALL_Workstations.Count) Workstation ) "
    $chart.titles[0].font             = $titlefont
    $chart.titles[0].forecolor        = "Red"
 

    #add data to chart 
    
 
    [void]$Chart.Series.Add("Data") 
     $Chart.Series["Data"].Points.DataBindXY( $WorkstationOSArray,  $WorkstationCountArray)
     $Chart.Series["Data"].IsvalueShownAsLabel=$true 

 
    # add title and axes labels 
 
 
    $ChartArea.AxisX.Title      = " Workstation O.S Version  " 
    $chartArea.AxisY.Title      = " Count"
    $ChartArea.AxisX.Interval   = 1
 
  
     #Find point with max/min values and change their colour 
    $maxValuePoint = $Chart.Series["Data"].Points.FindMaxByValue() 
    $maxValuePoint.Color = [System.Drawing.Color]::pink 
 
    $minValuePoint = $Chart.Series["Data"].Points.FindMinByValue() 
    $minValuePoint.Color = [System.Drawing.Color]::yellow
 
 
    # change chart area colour 
    $Chart.BackColor = "WHITE"
     $chart.Series["Data"]["DrawingStyle"] = "Emboss"
 
 
    # save chart to file 
    $Chart.SaveImage("$($CurrentScriptLocation)" +"\Workstation OS Version.jpeg", "JPEG")



}


#.................Code

if (!(Get-MyModule "ActiveDirectory"))

{ 
write-host
write-host " Active Directory Module not installed on this machine... Exiting" -foreground "RED"
write-host
Exit
}

Import-module activedirectory 

write-host
write-host " 2. Collecting Computer Objects data from your domain (Get-ADComputer)" -foreground "magenta"
write-host "...we are searching whole directory. Open the script and search for a line that start with (-SearchBase), uncomment it and enter a filter if you like" -foreground "Yellow"
write-host
$All_Computers = Get-ADComputer `
                -Filter {(PwdLastSet -gt $date) -and (enabled -eq $True)} `
                  -Property *  `
                   #      -SearchBase "DC=contoso,DC=COM" 
 [int]$All_Computers_Count = $All_Computers.count


write-host ".....$($All_Computers_Count) computers are detected." 


 
if($All_Computers.count -gt 0)
 {
        
                write-host
                write-host " 3. Aggregating data..." -foreground "magenta"
                write-host
              
               # Creating array of computer with custom properties so to obtain the OU property
                foreach ($computer in $All_Computers )
                        {

                          if(($computer.OperatingSystem))
                          {$computerOSClean = ($computer.OperatingSystem).Replace("Windows","Win").Replace("Enterprise","Ent").Replace("Professional","PRO").Replace("Standard","STD").Replace("Server","SRV")  }
                          else {$computerOSClean = $computer.OperatingSystem}
                          if($computer.OperatingSystemServicePack)
                          {$computerSPClean = $computer.OperatingSystemServicePack.Replace("Service Pack","SP")}
                          else {$computerSPClean = $computer.OperatingSystemServicePack}
                          
                          
                           
                         $OSExtra = $computerOSClean + " " + $computerSPClean 
                          $info = New-Object -TypeName psobject 
                           $info | Add-Member -MemberType NoteProperty -Name Name                               -Value $computer.Name
                            $info | Add-Member -MemberType NoteProperty -Name OperatingSystem                   -Value $computer.OperatingSystem
                             $info | Add-Member -MemberType NoteProperty -Name OperatingSystemServicePack       -Value $computer.OperatingSystemServicePack
                              $info | Add-Member -MemberType NoteProperty -Name OperatingSystemVersion          -Value $computer.OperatingSystemVersion
                               $info | Add-Member -MemberType NoteProperty -Name whenCreated                    -Value $computer.whenCreated
                                $info | Add-Member -MemberType NoteProperty -Name pwdLastSet                    -Value $computer.pwdLastSet
                                 $info | Add-Member -MemberType NoteProperty -Name OU                           -Value (_getOU($computer.DistinguishedName))
                                   $info | Add-Member -MemberType NoteProperty -Name OSExtra                    -Value $OSExtra


                           $All_Computers_Custom  += $info 
                        }




                # Divide computers to servers and workstations
                $All_Servers         = $All_Computers_Custom  |
                                            where {$_.OperatingSystem -like "*server*"}
                $All_ServersCount = $All_Servers.Count

                # This will exclude computers with Empty Operating Systems
                $All_Workstations    = $All_Computers_Custom  |
                                            where {$_.OperatingSystem -notlike "*server*" -and $_.OperatingSystem -like "*Windows*" }
                [int]$All_WorkstationsCount = $All_Workstations.Count

                [int]$ALL_Unknown =  $All_Computers_Count - ( $All_ServersCount +$All_WorkstationsCount)



               # Aggregate ALL_Servers information
               [array] $ServerOSArray      = @()
               [array] $ServerCountArray   = @()

               $All_Servers_Aggregates = $All_Servers |Group-Object -Property OSExtra 
                    foreach ($Aggregate in $All_Servers_Aggregates)

                            {
                             $ServerOSArray     += $Aggregate.Name
                             $ServerCountArray  += $Aggregate.Count
                            }

                # Aggregate ALL_Workstations information
               [array] $WorkstationOSArray      = @()
               [array] $WorkstationCountArray   = @()

               $All_Workstation_Aggregates = $All_Workstations |Group-Object -Property OSExtra 
                    foreach ($Aggregate in $All_Workstation_Aggregates)

                            {
                             $WorkstationOSArray     += $Aggregate.Name
                             $WorkstationCountArray  += $Aggregate.Count
                            } 

                write-host
                write-host " 4. Create HTML, Charts and CSV files..." -foreground "magenta"
                write-host
               if($All_ServersCount){_drawServers}
               if($All_WorkstationsCount){_drawWorkstations}
               $OutputHTML = _generateHTML
               Add-Content $filename $OutputHTML
               #Export all collected data to CSV file
                $All_Computers_Custom | Select-Object Name,OperatingSystem,OperatingSystemServicePack,OperatingSystemVersion,whenCreated,OU| Export-CSV $filename2 -NoTypeInformation -Encoding UTF8
 
               



                write-host
                write-host " 5. Sending email to $($SMTP_to)" -foreground "magenta"
                write-host
                SendEmail $SMTP_from $SMTP_to "AD_Computer_OS_version_report _$(Get-Date -f 'yyyy-MM-dd')" $SMTP_Host $fileName

                write-host
                write-host " Hey, Check the CSV Created here $($filename2) that contains AD Computer Inventory..." -foreground "Yellow"
                write-host


                
 }

 else
 {
                write-host
                write-host " WARNING : NO Computers are detected... Existing" -foreground "red"
                write-host
                
               
 }


#--------------
# Script Ending Info
#--------------

#Displaying info




Write-Host
Sleep 2
Write-Host
write-host "--------------------------" 
write-host "Script Ends" -foreground Green
write-host "--------------------------"
write-host 
write-host "Send your feedback at http://ammarhasayen.wordpress.com"  


#--------------
# Script END
#--------------  
   
   









                          