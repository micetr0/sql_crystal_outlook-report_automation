<#*******************************************************************************
purpose: Output report via SQL source from iSCheduler

 Input: Report Code 'PRVAUDIT', 'LOGREVIEW', 'TEABR'    
 Output:  Individual report in excel format after call        

 Dependency: Get-PrevMonthDates from Main

 Modifications
 Date           Name                    Revision Notes
 ----------     --------------------    --------------------------
 08-Jul-2019    William Hu              Initial version
*******************************************************************************#>

Function Get-query_OutFileName ($queryFlag, $monthStart, $monthEnd)
{
 <#*******************************************************************************
 Purpose: Select Query and outfile name
 
 #output: sqlquery - to be used by the connection string
          outFile  - filename to be passed for naming output
          
 Modifications
 Date           Author          Description                     
 ---------------------------------------------------------
 2019-Jul-08    William Hu      Added TEABR SQL placeholder
 2019-03-07     William Hu      Added loginReview SQL
 2019-03-05     William Hu      Initial version
 *******************************************************************************#> 

 #write-host 'Get-Query: inside function' $monthStart $monthEnd

 $formatMonthStart = "$monthStart 12:00:00"
 $formatMonthEnd = "$monthEnd 11:59:59"
 
 #write-host 'Get-Query: post function before switch' $formatMonthStart $formatMonthEnd

     switch -regex ($queryFlag)
     {
     #privilege audit report
     'PRVAUDIT' { 
                    $SQLquery  = "
                                 SELECT AUDITTIME AS `"DATE`", UserName AS `"USER NAME`",eventType AS `"TYPE NAME`",actionName AS `"ACTION`", CHANGEDDETAIL AS `"DETAIL INFO`" 
                                 FROM dbo.vwaudits 
                                 WHERE username = `'whu`' AND AUDITTIME >= '$formatMonthStart' AND AUDITTIME < '$formatMonthEnd' ORDER BY audittime DESC
                                 "
                                 
                    $outFile   = "PrivelegeAudit"
                    ;break
                }

     #Login Review report
     'LOGREVIEW' {
                    $SQLquery = "
                                SELECT hreg.HealthRegion,hubs.HUBNAME,u.USERNAME,(u.FIRSTNAME +' '+ u.LASTNAME) as `"Name`", h.LoginDate, CASE h.IsLoginSuccess WHEN '1' THEN 'succeed' ELSE 'failed' END as `"RESULT`", h.LogoutDate,h.IP
                                FROM tblUserLoginHistory h 
                                INNER JOIN tblUsers u  ON u.USERID = h.UserId  AND u.ACTIVE = 1 
                                INNER JOIN tblHubs hubs  ON hubs.HUBID = u.SiteID  
                                INNER JOIN tblHealthRegions hreg ON hreg.HealthRegionID = hubs.HealthRegionID 
                                WHERE h.LoginDate >= '$formatMonthStart' and h.LoginDate < '$formatMonthEnd' 
                                ORDER BY h.LoginDate DESC
                                "

                    $outFile = "LoginReview"
                    ;break
                 }

     #TEABR monthly report
     'TEABR' {
                    $SQLquery = "
                                SELECT hreg.HealthRegion,hubs.HUBNAME,u.USERNAME,(u.FIRSTNAME +' '+ u.LASTNAME) as `"Name`", h.LoginDate, CASE h.IsLoginSuccess WHEN '1' THEN 'succeed' ELSE 'failed' END as `"RESULT`", h.LogoutDate,h.IP
                                FROM tblUserLoginHistory h 
                                INNER JOIN tblUsers u  ON u.USERID = h.UserId  AND u.ACTIVE = 1 
                                INNER JOIN tblHubs hubs  ON hubs.HUBID = u.SiteID  
                                INNER JOIN tblHealthRegions hreg ON hreg.HealthRegionID = hubs.HealthRegionID 
                                WHERE h.LoginDate >= '$formatMonthStart' and h.LoginDate < '$formatMonthEnd' 
                                ORDER BY h.LoginDate DESC
                                "

                    $outFile = "TEABR"
                    ;break
            }
    }


return $SQLquery, $outFile
}

Function Get-SQLReports ($query_OutFileName, $query_outputDir, $query_userDir, $query_pswdDir)
{
    <#*******************************************************************************
     Purpose: Query PROD SQL database 
 
     Modifications
     Date           Author          Description                     
     ---------------------------------------------------------
     16-DEC-2019  William Hu      Implemented secureString usage
     07-JUNE-2019 William Hu      Refactored for performance
     07-FEB-2019  William Hu      Initial version
    *******************************************************************************#>
    $outputFileLocation = $query_outputDir 
    $dataSource =  #INPUT DB IP AND PORT - ie.10.1.10.11,5555

    $pswdTemp = Get-Content $query_pswdDir | ConvertTo-SecureString  
    $pswd = (New-Object System.Management.Automation.PSCredential 'N/A', $pswdTemp).GetNetworkCredential().Password
    $userTemp = Get-Content $query_userDir | ConvertTo-SecureString  
    $user = (New-Object System.Management.Automation.PSCredential 'N/A', $userTemp).GetNetworkCredential().Password  
       
    $database = #SPECIFY THE DB NAME
    $sqlConnectionString = "Data Source=$dataSource; Network Library=DBMSSOCN; Database =$database; User ID=$user; Password=$pswd; Integrated Security=False;"

    $monthStart, $monthEnd = Get-prevMonthDates
    #convert string format so it works with SQL server 2008 by first convert string back to datetime then output again in different time format
    $monthStart = [Datetime]::ParseExact($monthStart,"dd/MM/yyyy",$null)
    $monthStart = $monthStart.ToString("yyyy-MM-dd")
    $monthEnd = [Datetime]::ParseExact($monthEnd,"dd/MM/yyyy",$null)
    $monthEnd = $monthEnd.ToString("yyyy-MM-dd")

    $sqlquery, $outFile  = Get-query_OutFileName $query_OutFileName $monthStart $monthEnd    #removed #$outFile_YYYY_MM

    #write-host 'Get-SQLReports: outputFile' $outFile
    #write-host 'Get-SQLReports: sql query' $sqlquery

    $sqlConnection = New-Object System.Data.SqlClient.SqlConnection($sqlConnectionString)
    $sqlCommand = New-Object System.Data.SqlClient.SqlCommand ($sqlQuery,$sqlConnection)
    $sqlConnection.Open()
    $adapter = new-object System.Data.SqlClient.SqlDataAdapter $sqlCommand
    $dataset = New-Object System.Data.DataSet
    [void] $adapter.Fill($dataset) #cast to void to avoid having rows added output to console
    $DataSetTable = $Dataset.Tables["Table"]
    $sqlConnection.Close()

    # excel Obj Setup #
    $xlsObj = New-Object -ComObject Excel.Application;
    $xlsObj.Visible = 0;
    $xlsWb = $xlsobj.Workbooks.Add();
    $xlsSh = $xlsWb.Worksheets.item(1);

    # build the Excel column heading:
    [Array] $getColumnNames = $DataSetTable.Columns | Select ColumnName;

    # build column header:
    [Int] $RowHeader = 1;
    foreach ($ColH in $getColumnNames)
    {
    $xlsSh.Cells.item(1, $RowHeader).font.bold = $true;
    $xlsSh.Cells.item(1, $RowHeader) = $ColH.ColumnName;
    $RowHeader++;
    };

    # adding the data start in row 2 column 1:
    [Int] $rowData = 2;
    [Int] $colData = 1;

    foreach ($rec in $DataSetTable.Rows)
    {
        foreach ($Coln in $getColumnNames)
        {
            # next line convert cell to be text only:
            $xlsSh.Cells.NumberFormat = "@";

            # populating columns:
            $xlsSh.Cells.Item($rowData, $colData) = `
            $rec.$($Coln.ColumnName).ToString();
            $ColData++;
        };
    $rowData++; $ColData = 1;
    };


    # adjusting columns in the Excel sheet:
    $xlsRng = $xlsSH.usedRange;
    $xlsRng.EntireColumn.AutoFit();

    # saving Excel file - if the file exist do delete then save
    $xlsFile = "$outputFileLocation\$outFile.xlsx"; #$outFile_YYYY_MM revised output
    #write-host 'Get-SQLReports: xlsFile is: ' $xlsFile

    if (Test-Path $xlsFile)
    {
    #write-host 'Get-SQLReports: if file exists'
    Remove-Item $xlsFile -Confirm:$false
    $xlsObj.ActiveWorkbook.SaveAs($xlsFile);
    }
    else
    {
    #write-host 'Get-SQLReports: if file doesnt exists'
    $xlsObj.ActiveWorkbook.SaveAs($xlsFile);
    };

    # quit Excel and terminate Excel Application process:
    $xlsObj.Quit(); (Get-Process Excel*) | foreach ($_) { $_.kill() };

} #end func


#Test calls:

#$a = Get-prevMonthDates
#Get-SQLReports LOGREVIEW $a 
