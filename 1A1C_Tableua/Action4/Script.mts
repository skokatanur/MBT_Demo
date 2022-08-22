'Navigate to the Link where the User defined report exists
If Test_Object("lnk_UDR_User_Defined_Reports").exist(5) Then
	Click_Object("lnk_UDR_User_Defined_Reports")
	wait 3
End If



'Function to Navigate to the specified report
Call fn_UserDefined_Reports_Navig(reportName)

'setting the property for the report
'Test_Object("ele_UDR_Report_Name").SetTOProperty "innertext",reportName

'waiting for the report to load
Call Tableau_Loading_Report_Buffer()
Wait(60)
'call SetTo_Object("ele_UDR_Report_Name", "innertext", reportName)
'setting the property for the report
'''Test_Object("ele_UDR_Report_Name").SetTOProperty "innertext",reportName
Wait(30)

'Check if the user is navigated to the report 
'ele_Assignment_Report
If Test_Object("ele_UDR_Report_Name").exist(90) Then
	 Call LogResult_And_CaptureImage("pg_User_Defined_Reports","User Navigated to the " & reportName& "Report Successfully","PASS","User Navigated to the " & reportName& "Report Successfully")
				
Else 
	Call LogResult_And_CaptureImage("pg_User_Defined_Reports","User Failed to Navigate to the " & reportName& "Report Successfully","FAIL","User Failed to Navigate to the " & reportName& "Report")
End If



'Function to Navigate to the specified report
Function fn_UserDefined_Reports_Navig(strreportname)

'If Test_Object("img_ViewMode_UDR").exist(5) then
'            'Test_Object("img_ViewMode_UDR").highlight
'            Test_Object("img_ViewMode_UDR").click
'            wait 2
'            Test_Object("ele_List_UDR").click
'            
'        End If
'Test_Object("lnk_UDR_User_Defined_Reports").click 
Wait 5
  
If Test_Object("img_UDR_Filter_icon").exist(5) then
    'Test_Object("img_UDR_Filter_icon").highlight
	Test_Object("img_UDR_Filter_icon").click    
End If
        
call Enter_Value_In_Edit_Field("txt_UDR_search_report",strreportname,"No")
        
Test_Object("ele_UDR_report_seacrIcon").click
      
Call SetTo_Click_Object("lnk_TransactionDetailLock", "name", strreportname)
Wait 5
'       If strreportname= "RepAssignmentReport-SIT_15643197438630" or strreportname="Rep Assignment Report" or strreportname="Rep Assignment Report - Manager FY20" Then 'FY20
If strreportname= "RepAssignmentReport-SIT_15643197438630" or strreportname="Rep Assignment Report" or strreportname="Rep Assignment Report FY22 - Manager" Then 	'FY21
	Wait 10
	Test_Object("lnk_Rep Assignment Report").Highlight
	Test_Object("lnk_Rep Assignment Report").click
'ElseIf strreportname ="Compensation Statement Unlock" Then  
ElseIf strreportname = "Compensation Statement FY21" Then  
	wait 2
'    added by suprita Apr16,2021
	If Test_Object("lnk_Dashboard2_report").exist(5) then
		Test_Object("lnk_Dashboard2_report").click
	End If
	
	Call Tableau_Loading_Report_Buffer()
	Wait(60)	
	'Check if the user is navigated to the report 
	If Test_Object("ele_UDR_Report2_Name").exist(90) Then
		 Call LogResult_And_CaptureImage("pg_User_Defined_Reports","User Navigated to the " & reportName& "Report Successfully","PASS","User Navigated to the " & reportName& "Report Successfully")
	Else 
		Call LogResult_And_CaptureImage("pg_User_Defined_Reports","User Failed to Navigate to the " & reportName& "Report Successfully","FAIL","User Failed to Navigate to the " & reportName& "Report")
	End If
	Exit Function
End  If

wait 10
If Test_Object("lnk_Dashboard_report").exist(5) then
	Test_Object("lnk_Dashboard_report").click
End If      
      
End Function
