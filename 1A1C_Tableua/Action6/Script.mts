
strreportname = Parameter("P_Report_Name")

'Function to Navigate to the specified report
'Call fn_UserDefined_Reports_Navig(strreportname)

'setting the property for the report
'Test_Object("ele_UDR_Report_Name").SetTOProperty "innertext",reportName

'waiting for the report to load
Call Tableau_Loading_Report_Buffer()
Wait(60)

If Test_Object("img_UDR_Filter_icon").exist(5) Then
	Click_Object("img_UDR_Filter_icon")
	wait 3
End If

call Enter_Value_In_Edit_Field("txt_UDR_search_report",strreportname,"No")
    
If Test_Object("ele_UDR_report_seacrIcon").exist(5) Then
	Click_Object("ele_UDR_report_seacrIcon")
	wait 3
End If    

If Test_Object("lnk_TransactionDetailLock").exist(5) Then
	Click_Object("lnk_TransactionDetailLock")
	wait 3
End If    

If  Test_Object("lnk_Dashboard_report").exist(5)Then
	Test_Object("lnk_Dashboard_report").Highlight
	Test_Object("lnk_Dashboard_report").click
	Wait 5
End If
	Call Tableau_Loading_Report_Buffer()

'Function to Navigate to the specified report
Function fn_UserDefined_Reports_Navig(reportname)
  
If Test_Object("img_UDR_Filter_icon").exist(5) then
    'Test_Object("img_UDR_Filter_icon").highlight
	Test_Object("img_UDR_Filter_icon").click    
End If
        
call Enter_Value_In_Edit_Field("txt_UDR_search_report",reportname,"No")
    
Test_Object("ele_UDR_report_seacrIcon").click
      
Call SetTo_Click_Object("lnk_TransactionDetailLock", "name", strreportname)
Wait 5

	Call Tableau_Loading_Report_Buffer()
	Wait(60)	
      
End Function
