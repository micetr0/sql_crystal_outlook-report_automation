 <#*******************************************************************************
 Purpose: Drives crystal report output to excel for reporting purpose

 
 Modifications
 Date			Author			Description						
 ---------------------------------------------------------
 16-Dec-2019   William Hu       Implemented secureString usage
 08-Dec-2019   William Hu       Standardized output via main_monthly path
 18-APR-2019   William Hu	    Initial version
 *******************************************************************************#>


function Get-CrystalReports($CRreportLoad, $CRreportOutputdir,$CRreportOutputType, $CRParmeter1, $CRParmeter2, $query_userDir, $query_pswdDir) 
{

 <#*******************************************************************************
 Purpose: Invoke call to run crystal reports and output to specified file format

 Output is defined by the exporttoDisk function (ExcelRecord / PortableDocFormat)
 
 Modifications
 Date			Author			Description						
 ---------------------------------------------------------
 16-Dec-2019   William Hu       Implemented secureString usage
 08-Dec-2019   William Hu       Removed filename locally - filename changed via ps_main_monthly
 18-APR-2019   William Hu	    Initial version
 *******************************************************************************#>

[reflection.assembly]::LoadWithPartialName('CrystalDecisions.Shared') |out-null
[reflection.assembly]::LoadWithPartialName('CrystalDecisions.CrystalReports.Engine') |Out-Null

#Add-Type -AssemblyName 'CrystalDecisions.Shared'
#Add-Type -AssemblyName 'CrystalDecisions.CrystalReports.Engine'

 #Crystal reports source repo
 $reportRepo_filepath  = 'C:\PS_outputs\PS_Outlook\CRY\monthly_crystals\' 
 #file output location
 $output_filepath =  $CRreportOutputdir

  <#debug info#>
 #write-output $CRreportLoad, $CRreportOutputdir,$CRreportOutputType, $CRParmeter1, $CRParmeter2, $query_userDir, $query_pswdDir
 #write-host "Get-CrystalReport: CRreportOutputdir: $CRreportOutputdir"
 #write-host "Get-CrystalReport: output_filepath: $output_filepath"
 #write-host "Get-CrystalReport: report name: $CRreportLoad"
 #write-host "Get-CrystalReport: parm1: $CRParmeter1"
 #write-host "Get-CrystalReport: parm2: $CRParmeter2"


 #Security Credential Setup to hide both username / password 
 $userTemp = Get-Content $query_userDir | ConvertTo-SecureString  
 $user = (New-Object System.Management.Automation.PSCredential 'N/A', $userTemp).GetNetworkCredential().Password  
 $pswdTemp = Get-Content $query_pswdDir | ConvertTo-SecureString  
 $pswd = (New-Object System.Management.Automation.PSCredential 'N/A', $pswdTemp).GetNetworkCredential().Password


if (Test-Path -Path $output_filepath) 
{
    #write-host "Get-CrystalReport: $output_filepath exists"
}
else
{
    #write-host "Get-CrystalReport: $output_filepath path not exists - created"
    New-Item -ItemType Directory -force -path "$output_filepath"
}
   
 #convert to datetime format for Crystal
 $fromDate      =  [datetime]::parseexact($CRParmeter1, 'dd/MM/yyyy', $null) 
 $toDate        =  [datetime]::parseexact($CRParmeter2, 'dd/MM/yyyy', $null)

 switch -regex ($CRreportOutputType)
     {   
       'ExcelRecord'        { $docExtension  = '.xls'; break}
       'PortableDocFormat'  { $docExtension  = '.pdf'; break}
     }

$report = New-Object CrystalDecisions.CrystalReports.Engine.ReportDocument

$report.Load("$reportRepo_filepath\$CRreportLoad")
#write-host "Get-CrystalReport: parm count: " ($report.ParameterFields.Count) 

$report.SetDatabaseLogon($user,$pswd)

    #parameter check since some reports dont need any to run
    if ($report.ParameterFields.Count -eq 0) 
    {
        #write-host "Get-CrystalReport: 0 count - no param"
    }
    else
    {
        #values from and to for all of the current reports
        $report.SetParameterValue("FromDate",$fromDate) 
        $report.SetParameterValue("ToDate",$toDate)
    }

#setting up filename and output path
$reportOutputPath = "$output_filepath\$CRreportLoad$docExtension"

#write-host "Get-CrystalReport: report output $reportOutputPath"

$report.ExportToDisk($CRreportOutputType,$reportOutputPath)
$report.close()
} #end main

