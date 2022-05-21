<#*******************************************************************************
purpose: Extract dmt report from outlook
   
 Dependency: Outlook inbox 'DMT Reports' 
             Outlook forwarding rule for incoming DMT mails
             Hard-Coded report ID to report name (rename function)
 
 Limitation: Atm only handles for case of one report return per month

 TO-DO: tweak Extract function to return only one report at all time

 Modifications
 Date           Name                    Revision Notes
 ----------     --------------------    --------------------------
 2019-Jul-08    William Hu              Initial version
*******************************************************************************#>

Function Extract-Attachment ($reportCode, $dmtOutputFolder)
 {
  <#*******************************************************************************
 Purpose: Grab DMT report from an designated inbox and 
          save them out to designated folder for further processing

          There should only be one month return per reportcode
 
 Dependency: Get-prevMonthDates function (this is covered as the main_monthly will run it)

 Parem: 
    IN  - accepts string reportCode, location
    OUT - returns the full file name path for renaming function
    
 Modifications
 Date			Author			Description						
 ---------------------------------------------------------
 10-Dec-2019 William Hu     Add return so filename can be used by the rename function after
 09-Dec-2019 William Hu     Reduce load time by adding date filter & refactor filter foreach loop
 27-FEb-2019 William Hu     Organized into a function
 25-Feb-2019 William Hu	    Initial version
 *******************************************************************************#> 

    $filePath = $dmtOutputFolder   #("C:\PS_outputs\PS_Outlook\DMT\" + $reportCode +"\") 
    #write-host "dmt_func_extract: filepath is $filePath"
    
    #for setting up -match param
    $reportCodeCombined =  '_' + $reportCode + '_'

    Add-type -assembly "Microsoft.Office.Interop.Outlook" | out-null
    $olFolders = "Microsoft.Office.Interop.Outlook.olDefaultFolders" -as [type]
    $outlook = new-object -comobject outlook.application
    $namespace = $outlook.GetNameSpace("MAPI") 
    $folder = $namespace.getDefaultFolder($olFolders::olFolderInBox).Folders.Item("DMT Reports")
 
    $sBeginDate, $sEndDate = Get-prevMonthDates 'dmt'

    #write-output 'test' $sBeginDate , $sEndDate
    #convert string typed date back to date type with the correct format to use in the if statement
    $dtBeginDate = [datetime]::ParseExact($sBeginDate,'dd/MM/yyyy',$null).ToString('MM/dd/yyyy')
    $dtEndDate = [datetime]::ParseExact($sEndDate,'dd/MM/yyyy',$null).ToString('MM/dd/yyyy')
 
    foreach ($item in $folder.items)
    {
        If ($item.ReceivedTime -ge $dtBeginDate -and $item.ReceivedTime -le $dtEndDate) 
        {                        
            foreach($attach in $item.Attachments)
            {        
                if($attach.filename.contains('xls') -and $attach.filename -match ($reportCodeCombined) )
                {   
                    $outfilePathCheck = $filePath + $attach.filename
                    #write-host "dmt_func_extract: $outfilePathCheck"
                   
                    if (-not( test-path $outfilePathCheck))
                    {   
                        #write-host "dmt_func_extract: file not exists - processing"
                        $attach.SaveAsFile($filePath + $attach.FileName) 
                        $outputedFileName = $attach.FileName                
                    }
                    else
                    {
                        #write-host "dmt_func_extract: file exists already"
                        $outputedFileName = $attach.FileName 
                    }        
                } #end if not exists
            }
                        
        }# end if date range

    } #foreach item

    #setting up path for next operation    
    Set-Location -Path $filePath
  return $outputedFileName

 }#end Function

 Function Change-FileName ($filenametoChange)
{
 <#*******************************************************************************
 Purpose: looks for DMT report code to parse the correct filename 

 NOTE: rep_code are hard-coded
 
 Modifications
 Date			Author			Description						
 ---------------------------------------------------------
 09-Jul-2019	William Hu	    Initial version
 *******************************************************************************#>     

    #write-host 'dmt_func_rename: file name to be changed' $fileNametoChange   
    $oldFileName = $fileNametoChange   

     switch -regex ($fileNametoChange)
     {   
       '_70_'  { $filenametoChange  = 'IssueTraks - OPEN.xls'; break}
       '_72_'  { $filenametoChange  = 'IssueTraks - NEW.xls'; break}
       '_73_'  { $filenametoChange  = 'IssueTraks - CLOSED.xls'; break}
       '_89_'  { $filenametoChange  = 'ACTIVE INITIATIVES Summary.xls'; break}
       '_103_' { $filenametoChange  = 'FHA_last_month.xls'; break}
       '_91_'  { $filenametoChange  = 'IssueTraks - TEST.xls'; break}  #testing purpose only
       '_90_'  { $filenametoChange  = 'IssueTraks-TEST2.xls'; break} #testing purpose only
     }

     #write-host "dmt_func_rename: new file name is $filenametoChange"
     $reNamepath = (get-location).ToString() + "\" + $oldFileName
     #write-host "renamepath is" $reNamepath   
     
     $outfilePathCheck = (get-location).ToString() + "\" + $filenametoChange
     #write-host "dmt_func_rename: $outfilePathCheck"
                   
     if (-not( test-path $outfilePathCheck))
     {   
        #write-host "dmt_func_rename: file not exists - processing"
        Rename-Item -path $reNamePath -NewName $filenametoChange -Force           
     }
     else
     {
        #write-host "dmt_func_rename: file exists already"
     }  

} #end of func   

 
Function Get-Outlook_reports($reportCode,$dmtOutputFolder)
{
                                    #Main()
 <#*******************************************************************************
 Purpose: Main() to call all functions

 Modifications
 Date			Author			Description						
 ---------------------------------------------------------
 10-DEC-2019 William Hu     Refactored 
 25-Feb-2019 William Hu	    Initial version
 *******************************************************************************#>

#write-host "Get-outlook_report: reportCode $reportCode dmtOutputFolder $dmtOutputFolder"
$outPutFileName = Extract-Attachment $reportCode $dmtOutputFolder
#write-host "Get-outlook_report: out file name from extract-attachment function is $outPutFileName"
Change-FileName ($outPutFileName)

} #end main




#Get-Outlook_reports 72 'C:\TEST\DMT'