
<# TO-DO:

COMPLETE Standard-Filename FUNCTION
CHANGE ALL WRITE-HOST TO WRITE-OUTPUT
LEARN WRITE-VERBOSE
CHECK TO SEE IF OUTLOOK IS IN USE, kill process if it is before running code


https://www.jsnover.com/blog/2013/12/07/write-host-considered-harmful/
https://gist.github.com/davefunkel/415a4a09165b8a6027a297085bf812c5 # follow template to clean-up

#>


<# Global Configuration #> 

# to be tweaked


Function Standard-Filename ($targetLocation)
{
<#*******************************************************************************
purpose: Standardize filename for all output files as YYYY_MM(-1) by interating 
         through selected folder structure and recursively append file prefix to them

 Modifications
 Date           Name                    Revision Notes
 ----------     --------------------    --------------------------
 08-DEC-2019    William Hu              Initial version
*******************************************************************************#>

    $filePath = $targetLocation
   
    Set-location $filePath

    #setup new file prefix
    $startDate,$endDate = Get-prevMonthDates
    $year = $startDate.substring(6,4)
    $month = $startDate.Substring(3,2)

    foreach ($f in Get-ChildItem -File -Recurse)
    {   
        $currentFileNamePath = $f.fullname
        #write-host $currentFileNamePath
        $newFileName =  "$year`_$month`_" + $f.name
        #write-host $newFileName

        Rename-Item -path $currentFileNamePath -NewName $newFileName -Force     

<#

        if (-not( test-path $currentFileNamePath)) #need to expand scope to include if new file also exists
        {   
            write-host "Main_std_filename: file not exists - processing"
            Rename-Item -path $currentFileNamePath -NewName $newFileName -Force           
        }
        else
        {
            write-host "Main_std_filename: file exists already"
        }
#>
    }      

} #end function

Function Get-prevMonthDates ($Mode)
{
<#*******************************************************************************
purpose: Output 2 date values to be used in the sqlQuery, DMT reports

 Mode Flag: if not null it'll grab current run month's begin/end date (used to run Get-Outlook_reports)

 Output: startOfprevMonth (string)
         endOfprevMonth (string)

 Modifications
 Date           Name                    Revision Notes
 ----------     --------------------    --------------------------
 02-JAN-2020     William Hu              Account for month of January, added optional flag to address incorrect date range for get-outlook_report
 08-FEB-2019     William Hu              Initial version
*******************************************************************************#>

$date = Get-Date    

IF ($Mode -ne $null) 
{
    #write-host "flag is not null set month and year for Get-Outlook_reports"
    $month = $date.Month
    $year  = $date.year
}
ELSE
{
    #write-host "flag is null"

    IF ($date.Month -eq 1) #account for when $month is 0 - this will only happen for December report ran in january
    {
        $month = 12
        $year = $date.Year - 1
    }
    ELSE
    {
        $month = $date.Month -1 # for previous month
        $year = $date.Year
    }    
}

#write-host 'before creating datetime obj' $month ,$year

    # create a new DateTime object set to the first day of a given month and year
    $startOfPevMonth_f = Get-Date -Year $year -Month $month -Day 1 -Hour 0 -Minute 0 -Second 0 -Millisecond 0 
    # add a month and subtract the smallest possible time unit
    $endOfPrevMonth_f = ($startOfPevMonth_f).AddMonths(1).AddTicks(-1)

    # apply format to match CR 
    $startofPrevMonth = $startOfPevMonth_f.ToString("dd/MM/yyyy")
    $endOfPrevMonth = $endOfPrevMonth_f.ToString("dd/MM/yyyy")

RETURN $startofPrevMonth, $endOfPrevMonth
}

Function Get-CVC_prevMonthDates
{
<#*******************************************************************************
purpose: Output 2 date values to be use to drive the cvc crystal reports

 Output: cvc_startOfprevMonth
         cvc_endOfprevMonth

 Modifications
 Date           Name                    Revision Notes
 ----------     --------------------    --------------------------
 02-JAN-2020     William Hu              Account for month of January
 13-AUG-2019     William Hu              Initial version
*******************************************************************************#>

$date = Get-Date    
$year = $date.Year

if ($date.Month -eq 1) #account for when $month is 0 - this will only happen for December report ran in january
{
    $month = 12
    $year = $date.Year - 1
}
else 
{
    $month = $date.Month -1 # for previous month
}
    
# create a new DateTime object set to the first day of a given month and year
$startOfPevMonth_f = Get-Date -Year $year -Month $month -Day 1 -Hour 0 -Minute 0 -Second 0 -Millisecond 0 
# add a month and subtract the smallest possible time unit
$endOfPrevMonth_f = ($startOfPevMonth_f).AddMonths(1).AddTicks(-1)

# apply format to match CR 
$cvc_startofPrevMonth = $startOfPevMonth_f.ToString("dd/MM/yyyy")
$cvc_endOfPrevMonth = $endOfPrevMonth_f.ToString("dd/MM/yyyy")

return $cvc_startofPrevMonth, $cvc_endOfPrevMonth
} 

Function upload-SPFile ($targetDir, $DestinationDir,$AutomationAccount, $AutomationCred) 
{

 <#*******************************************************************************
 reference:

https://social.msdn.microsoft.com/Forums/security/en-US/d838443e-5ecf-4de6-a768-d053422d35ef/sharepoint-online-powershell-command-executequery-doesnt-run-in-the-server?forum=sharepointdevelopment
https://stackoverflow.com/questions/43510867/how-to-connect-to-sharepoint-2013-on-premise-using-csom-in-powershell
https://www.toddklindt.com/blog/Lists/Posts/Post.aspx?ID=487

#adding auto credential 
http://duffney.io/AddCredentialsToPowerShellFunctions

 Purpose: Upload files to iScheduler reporting sharepoint site
 Required: installation of sharepointclientcomponents
 
 Modifications


 Date           Author          Description                     
 ---------------------------------------------------------
 23-DEC-2019   William Hu       Initial version
 *******************************************************************************#>

Add-Type -Path 'C:\Program Files\Common Files\Microsoft Shared\Web Server Extensions\15\ISAPI\Microsoft.SharePoint.Client.dll'
Add-Type -Path 'c:\Program Files\Common Files\microsoft shared\Web Server Extensions\15\ISAPI\Microsoft.SharePoint.Client.Runtime.dll'

$destination = $DestinationDir
$files = get-childitem $targetDir

#write-host "Upload-SPFile $files" $files

$password = Get-Content $AutomationCred | ConvertTo-SecureString  
$useTemp = Get-Content $AutomationAccount | ConvertTo-SecureString
$user = (New-Object System.Management.Automation.PSCredential 'N/A', $useTemp).GetNetworkCredential().Password


$credentials = New-Object System.Management.Automation.PSCredential ($user, $password)

$webclient = New-Object System.Net.WebClient 
$webclient.Credentials = $credentials 

    foreach ($file IN $files) 
    {
        $webclient.UploadFile($destination + "/" + $File.Name, "PUT", $File.FullName)
    }

    
    #eg func call: upload-SPFile 'C:\testSP.txt' 'http://SHAREPOINT.ORG/BLAHBLAHDIR/BLAHYEAR' $AutomationAcctDir $AutomationCredDir

} #end func
                                       
                                    
                                    #Main()
 <#*******************************************************************************
 Purpose: Main() to call all functions

 TO-DO: 
  
 1. Make all process run in parallel to save time
 2. grab the emailing portion from dmt3.ps1
 3. utlize PSScriptRoot $a = $PSScriptRoot  #write-host "$a\test.ps1"
 4. add event logger

 Modifications
 Date           Author          Description                     
 ---------------------------------------------------------
 22-Oct-2020    William Hu      Officially Implemented Sharepoint upload
 03-JAN-2020    William Hu      Exterminate any instance of Outlook prior to dmt func call
 16-DEC-2019    William Hu      Implemented Security to drive SQL server related functions
 05-DEC-2019    William Hu      Refactored - tweaked file naming / output location to be centralized
 08-JUL-2019    William Hu      Refactored 
 *******************************************************************************#>

 $beginTime = (Get-Date).DateTime

#CREDENTIAL SETUP
#Credentials to drive functions - using .password property as both user/name are encrypted
$CredUserDir =  #<TO-DO>WHERE YOUR USERNAME.SECURESTRING IS STORED<TO-DO>
$CredpswdDir =  #<TO-DO>WHERE YOUR PW.SECURESTRING IS STORED<TO-DO>
# for uploading to sharepoint and email to users
$AutomationAcctDir =  #<TO-DO>WHERE YOUR USERNAME.SECURESTRING IS STORED<TO-DO>
$AutomationCredDir =  #<TO-DO>WHERE YOUR USERNAME.SECURESTRING IS STORED<TO-DO>

#PROCESSING CONTENT SETUP
$outlook_repts_Array = 70,72,73,89,103 #for Get-Outlook_reports 

# 70  $dmtOutputFolder   # IssueTraks - OPEN
# 72  $dmtOutputFolder   # IssueTraks - NEW
# 73  $dmtOutputFolder   # IssueTraks - CLOSED
# 89  $dmtOutputFolder   # IssueTraks - ACTIVE INITIATIVES
# 103 $dmtOutputFolder   # IssueTraks - FHA LAST MONTH

$SQL_rpt_array = @("PRVAUDIT","LOGREVIEW")

#FOLDER SETUP

 #Script Location
 Set-Location  #<TO-DO>DIR WHERE ALL UR SCRIPT ARE STORED<TO-DO>

 #load functions
. .\ps_sqlsrv_func3.ps1
. .\ps_dmt_func_rev.ps1
. .\ps_crystal_func2.ps1

#intial run folder setup (create PS_outputs folder and the sub-folder 'MONTHLY')
$MasterFilepath = 'C:\PS_outputs\TEST\' #'C:\PS_outputs\' 

IF (Test-Path -Path $MasterFilepath) 
{
    write-host "MAIN: $MasterFilepath Exists"
}
ELSE
{
    write-host "MAIN: $MasterFilepath Path Not Exists - Creating folders"
    New-Item -ItemType Directory -force -path "$MasterFilepath" | Set-Location
    New-Item -ItemType Directory -force -path ((Get-location).ToString() + '\MONTHLY\')
}
# goes into the MONTHLY dir
#write-host 'MAIN: setting location'
Set-location -path "$MasterFilepath\MONTHLY\"

# Within the MONTHLY folder - YYYY_MM output file folder 
#add standard sub-folders (APP/DMT/CRY)
$startDate,$endDate = Get-prevMonthDates
$outputFolder = $startDate.substring(6,4) + '_' + $startDate.Substring(3,2)

IF (Test-Path -Path ((get-location).ToString() + "\" + $outputFolder)) 
{
    write-host MAIN: output folder of ((get-location).ToString() + "\" + $outputFolder) exists
}
ELSE
{
    write-host "MAIN: $outputFolder path not exists - created"
    New-Item -ItemType Directory -force -path ((get-location).ToString() + "\" + $outputFolder) | Set-location
    New-Item -ItemType Directory -force -Name "APP"
    New-Item -ItemType Directory -force -Name "DMT"
    New-Item -ItemType Directory -force -Name "CRY"   
    New-Item -ItemType Directory -force -Name "CVC" 
}
# END FOLDER SETUP

#Setup path to be used by functions below
$appOutputFolder = "$MasterFilepath" +"MONTHLY\$outputFolder\APP"
$dmtOutputFolder = "$MasterFilepath" +"MONTHLY\$outputFolder\DMT\"
$cryOutputfolder = "$MasterFilepath" +"MONTHLY\$outputFolder\CRY"
$cvcOutputfolder = "$MasterFilepath" +"MONTHLY\$outputFolder\CVC"

#dmt_func calls
Get-Process | Where-Object {$_.name -match "Outlook"} | Stop-process -Force # kill existing resources to prevent 'RPC Server Error'
foreach ($outlook_report in $outlook_repts_Array)
{   
    Get-Outlook_reports $outlook_report $dmtOutputFolder
}

#sqlsrv_func calls
foreach ($sql_report in $SQL_rpt_array)
{
    Get-SQLReports $sql_report $appOutputFolder $CredUserDir $CredpswdDir
}

#crystal_func calls
Set-Location C:\PS_outputs\PS_Outlook\CRY\monthly_crystals\ #set to where the crystal_repos are stored
$reports = get-childitem -Exclude 'archive' -Name

foreach ($report in $reports)
{   
    write-host "MAIN:cr name: $report"
    Get-CrystalReports $report $cryOutputfolder ExcelRecord $startDate $endDate $CredUserDir $CredpswdDir    
}

#Rename all output files according to naming standard
$targetFolder = "$MasterFilepath" +"MONTHLY\$outputFolder\"
#write-host "targetfolder: " $targetFolder
Standard-Filename $targetFolder


#Upload to SP2013
$folders = Get-ChildItem $targetFolder

foreach ($folder in $folders)
{   

    cd $folder

    $uploadDestination = "not specified"  
    #change SP upload path base on folder
    switch -regex ($folder)
         {   
           'APP'  { $uploadDestination  = '#<TO-DO>ADD EXACT SHAREPOINT SITE DIR LOCATION<TO-DO>'; break}
           #'CRY'  { $uploadDestination  = '############################################################################################CRY destination'; break}
           #'CVC'  { $uploadDestination  = '############################################################################################CVC destination'; break}
           'DMT'  { $uploadDestination  = '#<TO-DO>ADD EXACT SHAREPOINT SITE DIR LOCATION<TO-DO>'; break}
         }
        
    $uploadtargets = Get-ChildItem -recurse | %{$_.FullName}
    #write-host "upload targets" $uploadtargets

    
    foreach ($uploadtarget in $uploadtargets)
    {
        #write-host 'target: '      $uploadtarget 
        #write-host 'destination: ' $uploadDestination    

        upload-SPFile $uploadtarget $uploadDestination $AutomationAcctDir $AutomationCredDir
    }


    cd.. #reset for-loop pointer so it goes each of the 4 folders to grab the full file path
} #foreach 



#Run-Stats
$endTime = (get-Date).DateTime
$totalTime = New-TimeSpan -Start $beginTime -End $endTime
write-host "MAIN: total run time: $totalTime"



<######################################### to be used later ################################################################

Function Send-Email($subject,$recipients,$filePath)
{
#*******************************************************************************
#Purpose: Send email
#
#TO-DO: limitation of .to property as its a string not an array need to switch to using 
#        send-mailmessage
#
# Modifications
# Date                   Author                                 Description                                                                                    
#------------------------------------------------------------------------------------
#26-FEB-2019             William Hu                             Initial version
#*******************************************************************************#
     
    #populate from the array
    $recipients = @($recipients)  
    $outlookSent = New-Object -com Outlook.Application

    foreach ($recipient in $recipients) 
    {    
        $mail = $outlookSent.CreateItem(0)
        $mail.importance = 1   # 0(low)-2(high)
        $mail.subject = $subject 
        $mail.body = "Auto-generated - Do Not Reply"
        $mail.to = "$recipient"
        
        $outFiles = Get-ChildItem $filePath -exclude 'Archive' | Select-Object -ExpandProperty fullName
        write-host $outFile

            #grab all item in the folder and add to the recipient
            foreach ($outFile in $outFiles)
            {
                $mail.Attachments.Add($outFile)       
            }
        
        $mail.Send()
     }
    $outlookSent.Quit()
    Stop-Process -Name "OUTLOOK"  #optional - quit() works just as well
}#end function Send-Email

#>