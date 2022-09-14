reportName= Parameter("P_Report_Name")
repPath=parameter("p_Report_Path_In")

call Tableau_Buffer()
ReportPath= fnDownload_CRL_Report(reportName)
vFilePath=repPath&reportName&".csv"
vFilePathWithXls=repPath&reportName&".xls"
call Tableau_Buffer()
call Tableau_Buffer()
Wait(60)
'Wait(60)
'Rep_Out_Path= fn_Convert_CSV_TO_Xlsx(vFilePath)
'Rep_Out_Path= fn_Convert_CSV_TO_Xlsx(vFilePathWithXls)
parameter("p_Report_Path_Out")=vFilePathWithXls
parameter("p_Report_Name_Out")=reportName
