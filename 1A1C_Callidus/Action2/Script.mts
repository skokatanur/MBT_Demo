
strMappingFileName = Parameter("P_Mapping_File_Name_In")
strMappingSheetName = Parameter("P_Mapping_Sheet_Name_In")
sheet1=split(strMappingSheetName,":")
strMappingSheetName=sheet1(0)
SheetName=parameter("P_Output_Sheet_Name_In")
strSrcColumn = Parameter("P_Source_column_Name")
strBseColumn = Parameter("p_Callidus_column_Name")
strOutputSheetName=parameter("P_Output_Sheet_Name_In")
StrCalenderFileName="CalendarData.xlsx" 
StrCalenderSheetName="Period"

'

'Environment.value("PROJECT_FOLDER_PATH")="C:\FAST_Test_Automation\"
strpath="C:\FAST_Test_Automation\"
''''	'Tableau Detail report path
'	call click_object("ele_Apps")
'	call click_object("lnk_Commissions")
'	If Test_object("ele_Comm_Apps").exist(10) Then
'		
'	End If

'Create Output Sheet
set objSheet = create_1A1C_Output_File(strOutputSheetName,SheetName)
set objDictionary_Mapping = get_Source_Base_Mapping_field_Names(strMappingFileName,strMappingSheetName, strSrcColumn, strBseColumn) 

'Call Add_Source_Headers_In_Output_File(objDictionary_Mapping, strSrcColumn)
Call Add_Source_Headers_In_Output_File_MIPR_Report(strMappingSheetName,objDictionary_Mapping, strSrcColumn)
Participant_Columns_Start = objSheet.usedrange.columns.count
str_participant_count=Participant_Columns_Start

'location of the Downloded file
strFolderLocation= strpath&"Test_Data\Callidus\Callidus_To_CRL"
strfilename=parameter("P_Tableau_Input_File_In")

'Renaming the sheet.
Set objExcel1 = CreateObject("Excel.Application")
objExcel1.Visible = True
objExcel1.DisplayAlerts = False
Set objWorkbook1 = objExcel1.Workbooks.Open (strfilename)'(strFolderLocation&"\"&strfilename)
Set ws=objWorkbook1.Sheets(strOutputSheetName)
ws.Activate
'Getting the row count excel
TotalNumRows=ws.UsedRange.Rows.Count
'Getting the column count excel
Column_Count = ws.UsedRange.Columns.Count
	'Adding headed names
	Set objDictionary_TableauHeaderData = CreateObject("Scripting.Dictionary")
    objDictionary_TableauHeaderData.RemoveAll
    'To get the Tablueau Header data
    For loop1 = 2 To 2 Step 1
        For loop2 = 1 To Column_Count Step 1
            TcolumnHeader=ws.cells(1,loop2).value
            'TcolumnValue=ObjExcelSheet1.cells(loop1,loop2).value
            TcolumnValue=TcolumnValue+1
            TcolumnHeader=cstr(TcolumnHeader)
            On error resume next
            objDictionary_TableauHeaderData.Add TcolumnHeader, TcolumnValue
        Next    
    next
    str_Tablueau_ColumnNames=objDictionary_TableauHeaderData.Keys
    'Get the column no for "Period Start Date
    str_Tablueau_ColumnNumber=objDictionary_TableauHeaderData.Item("Period Start Date")
    

    '''***************************Create dictionary for period start date values ENDED ***********************************************
    '''New code to match the period in source(tableau)
        '''***************************Create dictionary for period start date values******************************************************
'''    '''Commenting on 13th May 2020 as added a new parameter to validate for a specific period
	Set odictperiod=createobject("Scripting.Dictionary")
	odictperiod.RemoveAll
	For i = 2 To TotalNumRows Step 1
	If odictperiod.Exists(ws.cells(i,str_Tablueau_ColumnNumber).value)<> True Then
	    odictperiod.Add ws.cells(i,str_Tablueau_ColumnNumber).value, i&"period"
	End If
	Next
	period=odictperiod.Keys
	
	For Iterator1 = 0 To UBOUND(period) Step 1                
		strperiod1=period(Iterator1)
		stryear=year(strperiod1)
		str_year_tobe_validated=year(parameter("P_Emp_Validation_For_Period"))
		If stryear=str_year_tobe_validated Then
			strmonth=month(strperiod1)
			str_month_tobe_validated=month(parameter("P_Emp_Validation_For_Period"))
			If strmonth=str_month_tobe_validated Then
				strperiod1="#"+strperiod1+"#"
				''msgbox strperiod1
				Exit for
			End If
		End If
	Next
	

    
    
''''    'Create dictionary for EmpID and select distinc emp id from the tabalu result 
'''''    stremp_col_num=objDictionary_TableauHeaderData.Item("Employee ID")
'''''    Set odictempid=createobject("Scripting.Dictionary")
'''''    odictempid.RemoveAll
'''''    For i = 2 To TotalNumRows Step 1
'''''        If odictempid.Exists(ws.cells(i,stremp_col_num).value)<> True Then
'''''            odictempid.Add ws.cells(i,stremp_col_num).value, i&"ID"
'''''        End If
'''''    Next
'''''
	''strperiod1="#"&parameter("P_Emp_Validation_For_Period")&"#"
	empid=Split(Parameter("P_Num_of_Employee"),"||")


'
''''empid=odictempid.Keys
''''objExcel1.close
''''objExcel1.quit
'''
'''Test cdoe


'''Test Code
'''********************************************************VALIDATION of POSITION DATA IN CALLIDUS '********************************************************
'''Navigate to postions
call fn_Callidus_Sales_Commission_Reports_Navigation("Positions")

'-----------set advance filter criteria
intSourceTable_RowLine=3
strfilterflag=0
For intempid = 0 To ubound(empid) Step 1

    empid1=empid(intempid)
  set objconnection_Excel = createobject("ADODB.connection")
    'strSqlQuery_SourceTable= fnCreateSheetName(objExcel1,strfilename,strFolderLocation,strOutputSheetName,SheetName,objDictionary_Mapping,strperiod)
    strSqlQuery_SourceTable=fnSqlQuery(objExcel1,strfilename,strFolderLocation,strOutputSheetName,SheetName,objDictionary_Mapping,strperiod1,empid1)
    
    'Set ObjRecordSet_SourceTable = execute_SQL_Query1(strSqlQuery_SourceTable)
     Set ObjRecordSet_SourceTable=execute_SQL_Query2(strSqlQuery_SourceTable,objconnection_Excel)
        Do Until ObjRecordSet_SourceTable.EOF
            'Fetching Field Names and Values for Advance filter
            strParticipant = fn_Callidus_Emp_Len(ObjRecordSet_SourceTable("Employee ID").value)
            strperiod2=ObjRecordSet_SourceTable("Period Start Date").value
            strFieldNames = "Participant"
            strFieldValues = strParticipant
            
            
            If strfilterflag=0 Then
                'Advance Filter
     
                Call fn_Callidus_Sales_Commissions_Set_Advanced_Filter(strFieldNames,strFieldValues)
                strfilterflag=1
            End If
            'Changing the dateformat
            strPeriod2 = MonthName(Month(strperiod2))&" "&Year(strperiod2)
            strSourceTablePeriod=MonthName(Month(ObjRecordSet_SourceTable("Period Start Date").value))& " "&Year(ObjRecordSet_SourceTable("Period Start Date").value)
            
             If cstr(strPeriod2)=cstr(strSourceTablePeriod) Then
            
                Call fn_Callidus_Sales_Commission_Set_Default_Period(strperiod2)
            End If
                        
            Call fn_GetCallidus_Applicationvalues("tbl_Position_Summary", objDictionary_Mapping,ObjRecordSet_SourceTable)
            
            ObjRecordSet_SourceTable.MoveNext
            
            intSourceTable_RowLine = intSourceTable_RowLine+1
            'To validate 10 records and then exit the Do loop
            If intempid>=ubound(empid) Then
                strposFlag=1
                Exit do
                
            End If
        wait 5
    Loop
    'To validate 10 records and then exit the For loop
        If strposFlag=1 Then
            call close_Excel_DB_Connection()
            Exit For
        End If
        call close_Excel_DB_Connection()
        strfilterflag=0
Next
''''--------------------------------------------Navigate to Participant: -----------------------------------------------------------------

call fn_Callidus_Sales_Commission_Reports_Navigation("Participants")
strMappingSheetName=sheet1(1)
'strSrcColumn = Parameter("P_Source_column_Name")
'strBseColumn = Parameter("P_base_column_Name")
strOutputSheetName=parameter("P_Output_Sheet_Name_In")
SheetName="MIPR_Report"
set objDictionary_Mapping = get_Source_Base_Mapping_field_Names(strMappingFileName,strMappingSheetName, strSrcColumn, strBseColumn) 
'Call Add_Source_Headers_In_Output_File_Participants(objDictionary_Mapping, strSrcColumn)

Call Add_Source_Headers_In_Output_File_MIPR_Report(strMappingSheetName,objDictionary_Mapping, strSrcColumn)
Deposits_Columns_Start=objSheet.usedrange.columns.count
intSourceTable_RowLine=3
strfilterflagparticipant=0
For intempid = 0 To ubound(empid) Step 1

    empid1=empid(intempid)
    
    set objconnection_Excel = createobject("ADODB.connection")
    'strSqlQuery_SourceTable= fnCreateSheetName(objExcel1,strfilename,strFolderLocation,strOutputSheetName,SheetName,objDictionary_Mapping,strperiod)
    strSqlQuery_SourceTable=fnSqlQuery(objExcel1,strfilename,strFolderLocation,strOutputSheetName,SheetName,objDictionary_Mapping,strperiod1,empid1)
    'Set ObjRecordSet_SourceTable = execute_SQL_Query1(strSqlQuery_SourceTable)
     Set ObjRecordSet_SourceTable=execute_SQL_Query2(strSqlQuery_SourceTable,objconnection_Excel)
        Do Until ObjRecordSet_SourceTable.EOF
            strperiod2=ObjRecordSet_SourceTable("Period Start Date").value
            strParticipant=fn_Callidus_Emp_Len(ObjRecordSet_SourceTable("Employee ID").value)
            'Filter Names and values
            strFieldNames = "Employee ID"
            strFieldValues = strParticipant             
                         
            If strfilterflagparticipant=0 Then
            
                'Advance Filter Function
                Call fn_Callidus_Sales_Commissions_Set_Advanced_Filter(strFieldNames,strFieldValues)
                strfilterflagparticipant=1
            End If
            '
            strPeriod2 = MonthName(Month(strperiod2))&" "&Year(strperiod2)
            strSourceTablePeriod=MonthName(Month(ObjRecordSet_SourceTable("Period Start Date").value))& " " &Year(ObjRecordSet_SourceTable("Period Start Date").value)
          
	        If cstr(strPeriod2)=cstr(strSourceTablePeriod) Then
	              'Default Period function
	              Call fn_Callidus_Sales_Commission_Set_Default_Period(strperiod2)
	        End If
            
            'validation the application values with the tableau report for Participant table    
            Call fn_GetCallidus_Applicationvalues("tbl_Participant_Summary", objDictionary_Mapping,ObjRecordSet_SourceTable)
            
            ObjRecordSet_SourceTable.MoveNext
            
            intSourceTable_RowLine = intSourceTable_RowLine+1
            
            'To validate 10 records and then exit the Do loop
            If intempid>=ubound(empid) Then
                strposFlag=2
                Exit do
                
            End If
        wait 5
    Loop
     strfilterflagparticipant=0
    'To validate 10 records and then exit the For loop
    
        If strposFlag=2 Then
            call close_Excel_DB_Connection()
            Exit For
        End If
        call close_Excel_DB_Connection()
       
Next

'''--------------------------------------------Navigate to Deposits: -----------------------------------------------------------------
Call fn_Callidus_Sales_Commission_Reports_Navigation("Calculations||Deposits")
''strMappingFileName = Parameter("P_Tableau_Input_File_In")
''strMappingSheetName = "Credits"'Parameter("P_Mapping_Sheet_Name_In")
strMappingSheetName=sheet1(2)
'strSrcColumn = Parameter("P_Source_column_Name")
'strBseColumn = Parameter("P_base_column_Name")
strOutputSheetName=parameter("P_Output_Sheet_Name_In")
SheetName="MIPR_Report"
set objDictionary_Mapping = get_Source_Base_Mapping_field_Names(strMappingFileName,strMappingSheetName, strSrcColumn, strBseColumn) 
'Call Add_Source_Headers_In_Output_File_Credits(objDictionary_Mapping, strSrcColumn)

Call Add_Source_Headers_In_Output_File_MIPR_Report(strMappingSheetName,objDictionary_Mapping, strSrcColumn)
Incentives_Columns_Start=objSheet.usedrange.columns.count
intSourceTable_RowLine=3

For intempid = 0 To ubound(empid) Step 1

    empid1=empid(intempid)
    
    set objconnection_Excel = createobject("ADODB.connection")
    'strSqlQuery_SourceTable= fnCreateSheetName(objExcel1,strfilename,strFolderLocation,strOutputSheetName,SheetName,objDictionary_Mapping,strperiod)
    strSqlQuery_SourceTable=fnSqlQuery(objExcel1,strfilename,strFolderLocation,strOutputSheetName,SheetName,objDictionary_Mapping,strperiod1,empid1)
    'Set ObjRecordSet_SourceTable = execute_SQL_Query1(strSqlQuery_SourceTable)
      Set ObjRecordSet_SourceTable=execute_SQL_Query2(strSqlQuery_SourceTable,objconnection_Excel)   
        Do Until ObjRecordSet_SourceTable.EOF
            'Function to check EMP ID is 8 digits and add '0' if it is less than 8 digits
            strParticipant=fn_Callidus_Emp_Len(ObjRecordSet_SourceTable("Employee ID").value)
            strMetricNum=ObjRecordSet_SourceTable("Metric Name").value
            strmetricname=split(strMetricNum, "-")
            strfilter=split(trim(strmetricname(0)), " ")
            strMetricNum=strfilter(1)&strfilter(0)
            'strMetricName="DRO_"&strMetricNum&"_MIPR_Hold"
			'strMetricName="DR_MIPR_Hold_"&strMetricNum
			 strMetricName="MIPR_Hold"
            'filter names and values
            strFieldNames = "Participant,Name" 
            strFieldValues = strParticipant&","&strMetricName
            strperiod2=ObjRecordSet_SourceTable("Period Start Date").value
            If strfilterflagcredit=0 Then
                'Advance filter 
                Call fn_Callidus_Sales_Commissions_Set_Advanced_Filter(strFieldNames,strFieldValues)
                strfilterflagcredit=1
            End If
            'Changing the Date format
            strPeriod2 = MonthName(Month(strperiod2))&" "&Year(strperiod2)
            strSourceTablePeriod=MonthName(Month(ObjRecordSet_SourceTable("Period Start Date").value))& " " &Year(ObjRecordSet_SourceTable("Period Start Date").value)
            
            If cstr(strPeriod2)=cstr(strSourceTablePeriod) Then
                'Default Period function
                Call fn_Callidus_Sales_Commission_Set_Default_Period(strperiod2)
                
            End If
            
            'validation the application values with the tableau report for Deposits table        
            Call fn_GetCallidus_Applicationvalues("tbl_Deposit_Summary", objDictionary_Mapping,ObjRecordSet_SourceTable)
            
            ObjRecordSet_SourceTable.MoveNext
            'intperiod=intperiod+1
            intSourceTable_RowLine = intSourceTable_RowLine+1
            
            'To validate 10 records and then exit the Do loop
            If intempid>=ubound(empid) Then
                strposFlag=3
                Exit do
                
            End If
        wait 5
    Loop
    'To validate 10 records and then exit the For loop
        If strposFlag=3 Then
            call close_Excel_DB_Connection()
            Exit For
        End If
    'Close the DB connection    
    call close_Excel_DB_Connection()
	strfilterflagcredit = 0    
Next

'***************New code for Position Groups

Call fn_Callidus_Sales_Commission_Reports_Navigation("Global Values||Position Groups")
strMappingSheetName=sheet1(3)
strOutputSheetName=parameter("P_Output_Sheet_Name_In")
SheetName="MIPR_Report"
set objDictionary_Mapping = get_Source_Base_Mapping_field_Names(strMappingFileName,strMappingSheetName, strSrcColumn, strBseColumn) 


Call Add_Source_Headers_In_Output_File_MIPR_Report(strMappingSheetName,objDictionary_Mapping, strSrcColumn)
intSourceTable_RowLine=3

For intempid = 0 To ubound(empid) Step 1

    empid1=empid(intempid)
    
    set objconnection_Excel = createobject("ADODB.connection")
    'strSqlQuery_SourceTable= fnCreateSheetName(objExcel1,strfilename,strFolderLocation,strOutputSheetName,SheetName,objDictionary_Mapping,strperiod)
    strSqlQuery_SourceTable=fnSqlQuery(objExcel1,strfilename,strFolderLocation,strOutputSheetName,SheetName,objDictionary_Mapping,strperiod1,empid1)

    Set ObjRecordSet_SourceTable = execute_SQL_Query1(strSqlQuery_SourceTable)
    
        Do Until ObjRecordSet_SourceTable.EOF
            strParticipant=fn_Callidus_Emp_Len(ObjRecordSet_SourceTable("Employee ID").value)
            strperiod2=ObjRecordSet_SourceTable("Period Start Date").value
            'New filter values as discussed with Daniela
            Str_Position_Groups=ObjRecordSet_SourceTable("Position Group").value
            
            'strMetricName="IRO_"&strMetricNum&"_Commission_Reporting"
            'strMetricName="IRO_"&strMetricNum&"_MIPR_Hold"
            strFieldNames = "Name"            'this is the order that needs to be followed
            strFieldValues = Str_Position_Groups
            
            If strfilterflagincentive=0  Then
                'Setting Advance Filter
                Call fn_Callidus_Sales_Commissions_Set_Advanced_Filter(strFieldNames,strFieldValues)
                strfilterflagincentive=1
            End If
            
            'Changing the Date format
            strPeriod2 = MonthName(Month(strperiod2))&" "&Year(strperiod2)
            strSourceTablePeriod=MonthName(Month(ObjRecordSet_SourceTable("Period Start Date").value))& " " &Year(ObjRecordSet_SourceTable("Period Start Date").value)
            
             If cstr(strPeriod2)=cstr(strSourceTablePeriod) Then
                'Default Period function
                Call fn_Callidus_Sales_Commission_Set_Default_Period(strperiod2)
                
            End If
            
            'validation the application values with the tableau report for Incentive table        
            Call fn_GetCallidus_Applicationvalues("tbl_Position_Group", objDictionary_Mapping,ObjRecordSet_SourceTable)
            
            ObjRecordSet_SourceTable.MoveNext
            'intperiod=intperiod+1
            intSourceTable_RowLine = intSourceTable_RowLine+1
            
            'To validate 10 records and then exit the Do loop
            If intempid>=ubound(empid) Then
                strposFlag=4
                Exit do
                
            End If
        wait 5
    Loop
    'To validate 10 records and then exit the For loop
        If strposFlag=4 Then
            call close_Excel_DB_Connection()
            Exit For
        End If
    'Close DB connection    
    call close_Excel_DB_Connection()
    strfilterflagincentive=0
Next

''''************************* New code for calendar

Call fn_Callidus_Sales_Commission_Reports_Navigation("Global Values||Calendars")
strMappingSheetName=sheet1(4)
strOutputSheetName="MIPR_Report"'parameter("P_Output_Sheet_Name_In")
SheetName="MIPR_Report"
set objDictionary_Mapping = get_Source_Base_Mapping_field_Names(strMappingFileName,strMappingSheetName, strSrcColumn, strBseColumn)
'Call Add_Source_Headers_In_Output_File_Transaction_Report(strMappingSheetName,objDictionary_Mapping, strSrcColumn)
i=0
str_No_Of_Records=parameter("P_Num_of_Employee")
Call Add_Source_Headers_In_Output_File_MIPR_Report(strMappingSheetName,objDictionary_Mapping, strSrcColumn)
intSourceTable_RowLine=3
ValFlag=0
    StrCalenderName=strpath&"Test_Data\Callidus\Callidus_To_CRL\"&StrCalenderFileName
    
            If Test_Object("lnk_MF_Calender").Exist(10)Then
                Test_Object("lnk_MF_Calender").highlight
                Test_Object("lnk_MF_Calender").click

                If Test_Object("ele_Download_Calender").Exist(10)Then
                            Test_Object("ele_Download_Calender").highlight
                            Test_Object("ele_Download_Calender").click
                            Test_Object("ele_Download_selected_Calendars").highlight
                            Test_Object("ele_Download_selected_Calendars").click
                            
                     If Test_Object("txt_Calendar_Filename").exist(60) Then
                          call Enter_Value_In_Edit_Field("txt_Calendar_Filename", StrCalenderName,"No")
                          Click_Object("btn_Calendar_Save")
                          If Test_Object("btn_calendar_Replace_Yes").exist Then
                          	Test_Object("btn_calendar_Replace_Yes").click
                          End If
                        Wait 5
                        StrCalenderFile="CalendarData.xls"'parameter("P_CalenderFile_In")
                         'Call fn_Convert_Xlsx_TO_Xls(StrCalenderName,strFolderLocation,StrCalenderFile)
                        StrCalendarxls= fn_Convert_Xlsx_TO_Xls(StrCalenderName) 
                    End If
                 End If
             End If
    
 
'''' For intperiod=0 To ubound(objDictionary) Step 1
For intempid = 0 To ubound(empid) Step 1

    empid1=empid(intempid)
    
   ' strperiod1="1/10/2019"
    set objconnection_Excel = createobject("ADODB.connection")
   
If Icounter = 0 Then
	strSqlQuery_SourceTable=fnSqlQuery(objExcel1,strfilename,strFolderLocation,strOutputSheetName,SheetName,objDictionary_Mapping,strperiod1,empid1)
 	Set ObjRecordSet_SourceTable = execute_SQL_Query1(strSqlQuery_SourceTable)
 	Icounter=Icounter+1
 	else
 	
 	stramount = ObjRecordSet_SourceTable.fields("Amount Held").value
 	strSqlQuery_SourceTable = fnSqlQuery1(objExcel1,strfilename,strFolderLocation,strOutputSheetName,SheetName,objDictionary_Mapping,strperiod1,empid1)
    Set ObjRecordSet_SourceTable = execute_SQL_Query2(strSqlQuery_SourceTable,objconnection_Excel)  
 	
End If 	
' 	
 
 		set objconnection_Excel1 = createobject("ADODB.connection")       
				
				StrCalendarxls="C:\FAST_Test_Automation\Test_Data\Callidus\Callidus_To_CRL\CalendarData.xls"
				
				objconnection_Excel1.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source="&StrCalendarxls&";Extended Properties= ""Excel 8.0;HDR=Yes;IMEX=1"""
				objconnection_Excel1.Open
			    strcounter = 0
	
 	Do Until ObjRecordSet_SourceTable.EOF'strSqlQuery_SourceTable.EOF
			
				strPeriod = ObjRecordSet_SourceTable.fields("Period Start Date").value  
				strPeriod = MonthName(Month(strPeriod))&" "&Year(strPeriod)
				strPeriodYear=Year(strPeriod)
				StrPeriodType="month"
				
				strcounter = strcounter + 1
				
					
				
				strSQLQuery ="select * from [" & StrCalenderSheetName & "$] Where [Name]='"&strPeriod&"'" 'and [Short Name]='"&strPeriodYear&"'
				'strSQLQuery ="select * from [" & StrCalenderSheetName & "$] Where [Period Type]='"&StrPeriodType&"'  and [Name]='"&strPeriod&"'" 'and [Short Name]='"&strPeriodYear&"'
				'Set objRecord_Excel1 = execute_SQL_Query1(strSQLQuery,objconnection_Excel) 
				Set objRecord_Excel1 = objconnection_Excel1.Execute(strSQLQuery)
				    If objRecord_Excel1.EOF = true Then
				        Reporter.ReportEvent micFail, "Execute SQL Query "&sqlQuery, "No result found for the sql query "&sqlQuery
				        
				    Else
				        objRecord_Excel1.MoveFirst    
				        
				    End If
			    
			   
			       StrShortName = objRecord_Excel1.Fields("Short Name").value    
			       StrName = objRecord_Excel1.Fields("Name").value
			       stryear1=split(StrShortName, " ")
			       stryear=stryear1(1)
			       strColValue_SourceTable=strPeriodYear
			    If cstr(StrName)=cstr(strPeriod) Then
				        
				        objSheet.cells(intSourceTable_RowLine,i+20).Value = StrName    
				        objSheet.cells(intSourceTable_RowLine,i+20).Interior.ColorIndex = 4
				        Reporter.ReportEvent micPass, "TABLEAU APPLICATION VALUE MATCHES WITH CALLIDUS APPLICATION","TABLEAU APPLICATION VALUE FOR Fiscal Year is "&strPeriod&" MATCHES WITH CALLIDUS APPLICATION FOR" &strColValue_Application& " is " &StrName 
						strpass = 0
			                           
				ELSE        
			                    
			
			            objSheet.cells(intSourceTable_RowLine,i+20).Value = StrName    
			            objSheet.cells(intSourceTable_RowLine,i+20).Interior.ColorIndex = 3
			            res = strcounter& ".validation of column value for Fiscal Year is failed. TABLEAU APPLICATION VALUE is "&strColValue_SourceTable&" and CALLIDUS APPLICATION value is "&StrName
			            reporter.ReportEvent micFail, "validating the presence of column value for Fiscal Year ", res
			            objSheet.cells(intSourceTable_RowLine,1).Value = objSheet.cells(intSourceTable_RowLine,1).Value & res & vblf
			            call FN_Update_Status("",res)
			            strcounter = strcounter + 1
			   End If 
			         
			       
			       
			    If (Cint(stryear)=Cint(strColValue_SourceTable)) Then
			
			        objSheet.cells(intSourceTable_RowLine,i+21).Value = stryear    
			        objSheet.cells(intSourceTable_RowLine,i+21).Interior.ColorIndex = 4
			        Reporter.ReportEvent micPass, "TABLEAU APPLICATION VALUE MATCHES WITH CALLIDUS APPLICATION","TABLEAU APPLICATION VALUE FOR Fiscal Year is "&strColValue_SourceTable&" MATCHES WITH CALLIDUS APPLICATION FOR" &strColValue_Application& " is " &StrShortName 
					strpass = 0
			                
			   ELSE        
			
			        objSheet.cells(intSourceTable_RowLine,i+21).Value = stryear    
			        objSheet.cells(intSourceTable_RowLine,i+21).Interior.ColorIndex = 3
			        res = strcounter& ".validation of column value for Fiscal Year is failed. TABLEAU APPLICATION VALUE is "&strColValue_SourceTable&" and CALLIDUS APPLICATION value is "&StrShortName
			        reporter.ReportEvent micFail, "validating the presence of column value for Fiscal Year ", res
			        objSheet.cells(intSourceTable_RowLine,1).Value = objSheet.cells(intSourceTable_RowLine,1).Value & res & vblf
			        call FN_Update_Status("",res)
			        strcounter = strcounter + 1
			     End If   
			
				
			    
			   intSourceTable_RowLine = intSourceTable_RowLine+1
			   ObjRecordSet_SourceTable.movenext
			       If cint(ValFlag)=Ubound(empid) Then
			       		Exit For
			       End If
			       
			       If strcounter = 2 Then
			       	Exit do
			       End If
		ValFlag=ValFlag + 1
       loop
     'Close DB connection    
	  call close_Excel_DB_Connection(objconnection_Excel)
	  call close_Excel_DB_Connection(objconnection_Excel1)
Next


'Save and Close the Output File
Call save_And_Close_Outputfile()



'********************************************************************END of Script***********************************************************************************


'''*****************************************************************Local Functions****************************************************************************************
'Function to get callidus application values and compare them against Tableau report values
'**************************************************************************************************************************************************************************

Function fn_GetCallidus_Applicationvalues(strTableName, objDictionary_Mapping,ObjRecordSet_SourceTable)
    Set objPosSummaryHeader = Test_Object(strTableName)
    Set objPosSummaryDictionary = CreateObject("Scripting.Dictionary")
    wait 5
    RowCount=Test_Object(strTableName).rowcount
    ColCoumnCount=Test_Object(strTableName).columncount(2)
    
    For intColIndex = 1 to ColCoumnCount
        objPosSummaryDictionary. Add objPosSummaryHeader.getcelldata(1,intColIndex+1 ), objPosSummaryHeader.getcelldata(2,intColIndex+1 )
    Next
    
    Select Case strTableName
        Case "tbl_Position_Summary"
                'Call fn_Validations(objPosSummaryDictionary, objDictionary_Mapping,ObjRecordSet_SourceTable)
                Call fn_Validations_New(RowCount,"Positions",objPosSummaryDictionary, objDictionary_Mapping,ObjRecordSet_SourceTable)
        Case "tbl_Participant_Summary" 
                'Call fn_Validations_Participants(objPosSummaryDictionary, objDictionary_Mapping,ObjRecordSet_SourceTable)
                Call fn_Validations_New(RowCount,"Participants",objPosSummaryDictionary, objDictionary_Mapping,ObjRecordSet_SourceTable)
        Case "tbl_Deposit_Summary"
                Call fn_Validations_New(RowCount,"Deposits",objPosSummaryDictionary, objDictionary_Mapping,ObjRecordSet_SourceTable)
                'Call fn_Validations_Credits(Credits,objPosSummaryDictionary, objDictionary_Mapping,ObjRecordSet_SourceTable)
'''        Case "tbl_Incentive_Summary"
'''                Call fn_Validations_New(RowCount,"Incentives",objPosSummaryDictionary, objDictionary_Mapping,ObjRecordSet_SourceTable)
        Case "tbl_Position_Group"
                Call fn_Validations_New(RowCount,"Position Group",objPosSummaryDictionary, objDictionary_Mapping,ObjRecordSet_SourceTable)
        Case "tbl_MF_Fiscal_Calendar"
                Call fn_Validations_New(RowCount,"Period",objPosSummaryDictionary, objDictionary_Mapping,ObjRecordSet_SourceTable)
        
                
    End Select

End Function

'**************************************************************************************************************************************************************************        
'Function to Create sheet at runtime and pass query based on the data from tableau report
'**************************************************************************************************************************************************************************
Function fnCreateSheetName(objExcel1,strfilename,strFolderLocation,strOutputSheetName,SheetName,objDictionary_Mapping,strperiod)
                Set objWorkbook1 = objExcel1.Workbooks.Open (strfilename)
                Set objWorksheet1 = objWorkbook1.Worksheets(1)
                    objWorksheet1.Name = SheetName
                    objWorkbook1.SaveAs (strFolderLocation&"\"&strOutputSheetName&".xls")
                    'objWorkbook1.save
                    objWorkbook1.close
                    objExcel1.Quit
                    objconnection_Excel.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source="&strFolderLocation&"\" &strOutputSheetName&".xls;Extended Properties= ""Excel 8.0;HDR=Yes;IMEX=1"""
                    objconnection_Excel.Open
                    'Execute the Excel query
                    fnCreateSheetName = "select * from ["&SheetName&"$] Where [Period Start Date]=#"&strperiod&"#"
                    
                    
End Function

'**************************************************************************************************************************************************************************'
'Menu and Sub Menu Navigation
'**************************************************************************************************************************************************************************'
Function fn_Callidus_Menu_SubMenu_Navigation(MenuSubmenu)

Menu=split(MenuSubmenu,"|")
If Test_Object("ele_Menu_Expand").exist(10)=True Then
    Browser("br_SAP_Commissions").Page("pg_SAP_Commissions").WebElement("ele_Menu_Expand").Click
End If
    '**************************************************Click on Expand to view all the Links******************************************************************************** 
If Test_Object("ele_Menu_Expand").exist(10)=True Then
    Browser("br_SAP_Commissions").Page("pg_SAP_Commissions").WebElement("ele_Menu_Expand").Click
End If

'**************************************************Click on Expand arrow for a specific Menu******************************************************************************** 
Set objDesc = Description.Create
objDesc("xpath").value = "//*[text()='"&Menu(0)&"']/../span[contains(@class,'cald-sub-nav  sap-icon-slim-arrow')]"
objDesc("visible").value = "true"
wait 2
set eleMenus = Browser("br_SAP_Commissions").Page("pg_SAP_Commissions").ChildObjects(objDesc)
If eleMenus.count>0 Then
    eleMenus(0).click
End If

'**************************************************Click on Sub-Menu with in a specific Menu******************************************************************************** 
Set submenu =Browser("br_SAP_Commissions").Page("pg_SAP_Commissions").Link("lnk_Sub_Menu_New")
submenu.SetTOProperty "name", Menu(1)
'submenu.SetTOProperty "xpath", "//*[contains(@class,'cald-submenu-list-container svelte') and text()='"& Parameter("P_Table_Name_In")&"']"

submenu.SetTOProperty "text", Menu(1)
wait 3
submenu.highlight
submenu.Click
wait 10
'**************************************************Clicking Collapse link for the Menu*************************************************************************************************************** 
Set objDesc = Description.Create
objDesc("xpath").value = "//*[text()='"&Menu(0)&"']/../span[contains(@class,'cald-sub-nav  sap-icon-slim-arrow')]"
objDesc("visible").value = "true"
wait 2
set eleMenus = Browser("br_SAP_Commissions").Page("pg_SAP_Commissions").ChildObjects(objDesc)
If eleMenus.count>0 Then
    eleMenus(0).click
End If
End Function

'**************************************************************************************************************************************************************************
'Add Header infor in the output sheet for MIPR Report
'**************************************************************************************************************************************************************************
Function Add_Source_Headers_In_Output_File_MIPR_Report(MappingSheetName,objDictionary_Mapping, strSourceTable_Name)
    objSheet.cells(1,1).Value = strSourceTable_Name
    intRecCounter = 1
    intSourceTable_RowLine = 2
    
    SourceTableCols_arr = objDictionary_Mapping.Keys
    'add Headers
    Select Case MappingSheetName
        Case "Positions"
                    For i=0 To ubound(SourceTableCols_arr) Step 1
                        If intRecCounter = 1 Then
                            objSheet.cells(intSourceTable_RowLine,1).Value = "REMARKS"  
                            objSheet.Columns("A").ColumnWidth = 80
                            objSheet.cells(intSourceTable_RowLine,i+2).Value = SourceTableCols_arr(i)
                            intRecCounter =intRecCounter+1 
                        else
                        objSheet.cells(intSourceTable_RowLine,i+2).Value = SourceTableCols_arr(i)
                        End If
                    Next
                    intSourceTable_RowLine = intSourceTable_RowLine+1
        Case "Participants"
                    For i=0 To ubound(SourceTableCols_arr) Step 1
                        If intRecCounter = 1 Then
                            objSheet.cells(intSourceTable_RowLine,1).Value = "REMARKS"  
                            objSheet.Columns("A").ColumnWidth = 80
                            objSheet.cells(intSourceTable_RowLine,(str_participant_count+1)).Value = SourceTableCols_arr(i)
                            intRecCounter =intRecCounter+1 
                        else
                            'objSheet.cells(intSourceTable_RowLine,i+10).Value = SourceTableCols_arr(i)
                            objSheet.cells(intSourceTable_RowLine,(str_participant_count+1)).Value = SourceTableCols_arr(i)
                        End If
                        str_participant_count=(str_participant_count+1)
                    Next
                    intSourceTable_RowLine = intSourceTable_RowLine+1
        Case "Deposits"
                    For i=0 To ubound(SourceTableCols_arr) Step 1
                        If intRecCounter = 1 Then
                            objSheet.cells(intSourceTable_RowLine,1).Value = "REMARKS"  
                            objSheet.Columns("A").ColumnWidth = 80
                            objSheet.cells(intSourceTable_RowLine,(Deposits_Columns_Start+1)).Value = SourceTableCols_arr(i)
                            intRecCounter =intRecCounter+1 
                        else
                            'objSheet.cells(intSourceTable_RowLine,i+13).Value = SourceTableCols_arr(i)
                            objSheet.cells(intSourceTable_RowLine,(Deposits_Columns_Start+1)).Value = SourceTableCols_arr(i)
                        End If
                        Deposits_Columns_Start=(Deposits_Columns_Start+1)
                    Next
                    intSourceTable_RowLine = intSourceTable_RowLine+1

         Case "Position Group"
                            For i=0 To ubound(SourceTableCols_arr) Step 1
                        If intRecCounter = 1 Then
                            objSheet.cells(intSourceTable_RowLine,1).Value = "REMARKS"  
                            objSheet.Columns("A").ColumnWidth = 80
                            'objSheet.cells(intSourceTable_RowLine,i+20).Value = SourceTableCols_arr(i)
                            'objSheet.cells(intSourceTable_RowLine,(Incentives_Columns_Start+1)).Value = SourceTableCols_arr(i)
                             objSheet.cells(intSourceTable_RowLine,i+19).Value = SourceTableCols_arr(i)
                            intRecCounter =intRecCounter+1 
                        else
                            'objSheet.cells(intSourceTable_RowLine,i+20).Value = SourceTableCols_arr(i)
                            'objSheet.cells(intSourceTable_RowLine,(Incentives_Columns_Start+1)).Value = SourceTableCols_arr(i)
                             objSheet.cells(intSourceTable_RowLine,i+19).Value = SourceTableCols_arr(i)
                        End If
                        'Incentives_Columns_Start=(Incentives_Columns_Start+1)
                    Next
                    intSourceTable_RowLine = intSourceTable_RowLine+1
        Case "Period"
                            For i=0 To ubound(SourceTableCols_arr) Step 1
                        If intRecCounter = 1 Then
                            objSheet.cells(intSourceTable_RowLine,1).Value = "REMARKS"  
                            objSheet.Columns("A").ColumnWidth = 80
                            objSheet.cells(intSourceTable_RowLine,i+20).Value = SourceTableCols_arr(i)
                            intRecCounter =intRecCounter+1 
                        else
                            objSheet.cells(intSourceTable_RowLine,i+20).Value = SourceTableCols_arr(i)
                        End If
                    Next
                    intSourceTable_RowLine = intSourceTable_RowLine+1
    End Select
    
End Function
'**************************************************************************************************************************************************************************
'Set Advance Filter using Child object method
'**************************************************************************************************************************************************************************
Function set_Filter_Advance_Search_Callidus(strFieldNames,strFieldValues)
    AdvSrhPUBUFlag=false
    Click_Object("ele_Advanced_Search")
    If AdvSrhPUBUFlag = false Then
        If Test_Object("txt_Process_Unit").exist(10)= true Then
            Test_Object("txt_Process_Unit").set "Global"
            Test_Object("ele_PU_Global").click
        End If
        If Test_Object("txt_Business_Unit").exist(10)= true Then
            Test_Object("txt_Business_Unit").set "BU Global"
            Test_Object("ele_BU_Global").click
        End If
        AdvSrhPUBUFlag = true
    End If
    

    fieldNames = split(strFieldNames,",")
    fieldValues = split(strFieldValues,",")
    Set tblObj = Test_Object("tbl_Field_Name_filter")
    toalrow = tblObj.getROProperty("rows")
'    If toalrow > 2 then
'        For j=2 To toalrow-1 Step 1
'            tblObj.ChildItem(j, 4, "WebElement",0).click
'        Next
'    End If
    'xpath="//*[@class='ember-view ember-text-field w45']"
    For i = 0 To Ubound(fieldNames) Step 1
        tblObj.ChildItem(i+2, 1, "WebList",0).select trim(fieldNames(i))
        tblObj.ChildItem(i+2, 2, "WebList",0).select "equals"
        wait 1

        CI_objcount = tblObj.ChildItemCount(i+2, 3, "WebEdit")
        If CI_objcount>0 Then
            If tblObj.ChildItem(i+2, 3, "WebElement",7).exist Then
                If tblObj.ChildItem(i+2, 3, "WebElement",7).getroproperty("visible") = true Then
                    tblObj.ChildItem(i+2, 3, "WebElement",7).click
                    If fieldNames(i)="Value" Then
                        Test_Object("txt_amount_value_search").set trim(fieldValues(i))
                        If tblObj.ChildItem(i+2, 3, "WebElement",8).exist Then
                            If tblObj.ChildItem(i+2, 3, "WebElement",8).getroproperty("visible") = true Then
                                'tblObj.ChildItem(i+2, 3, "WebElement",8).click
                                Test_Object("ele_Amount_Search_dropdown").click
                                Test_Object("txt_Search_Value_Filter").set trim(fieldValues(i+1))
                            End  IF
                        End If
                    else
                        Test_Object("txt_Search_Value_Filter").set trim(fieldValues(i))
                    End If
                    
                    wait 5
                    If Test_Object("lst_Filter_Search_List").Getroproperty("outertext") = "No matches found" Then
                        strfound = 0
                        Exit for
                    Else
                        Test_Object("lst_Filter_Search_List").select trim(fieldValues(i))
                        strfound = 1
                    End If
                else
                    tblObj.ChildItem(i+2, 3, "WebEdit",0).set trim(fieldValues(i))
                End If
            else
                tblObj.ChildItem(i+2, 3, "WebEdit",0).set trim(fieldValues(i))
            End If
'        ElseIf tblObj.ChildItem(i+2, 3, "WebElement",7).exist Then
'            tblObj.ChildItem(i+2, 3, "WebElement",7).click
'            Test_Object("txt_Search_Value_Filter").set trim(fieldValues(i))
'            wait 5
'            If Test_Object("lst_Filter_Search_List").Getroproperty("outertext") = "No matches found" Then
'                strfound = 0
'                Exit for
'            Else
'                Test_Object("lst_Filter_Search_List").select trim(fieldValues(i))
'                strfound = 1
'            End If
        End If
        
        If tblObj.ChildItem(i+2, 4, "WebElement",0).GetROProperty("class") = "sap-icon-less comm-icon" Then
        else
            tblObj.ChildItem(i+2, 4, "WebElement",0).click
            wait 2
        End If
    Next
    
    Click_Object("btn_Apply")
    set_Filter_Advance_Search_Callidus = strfound
End Function        

'**************************************************************************************************************************************************************************
'Set Default Period function
'**************************************************************************************************************************************************************************

Function fn_Callidus_Sales_Commission_Set_Default_Period(strDefaultPeriod)
    strPeriod = MonthName(Month(strDefaultPeriod))&" "&Year(strDefaultPeriod)
    fn_Callidus_Sales_Commission_Set_Default_Period = "FAIL"
    If Click_Object("ele_Defaul_Period_Calendar") <> "PASS" Then
        Exit Function
    End If    
	wait 3
	
	 Browser("br_SAP_Commissions").Page("pg_SAP_Commissions").WebElement("ele_Period").Click
'    If Click_Object("ele_Period") <> "PASS" Then
'        Exit Function
'    End If    
    
    If Enter_Value_In_Edit_Field("txt_Comm_Search_Default_Period", strPeriod, "") <> "PASS" Then
        Exit Function
    End If
    
    If Click_Object("ele_Comm_Default_Period_List") <> "PASS" Then
        Exit Function
    End If    
    
    If Click_Object("btn_Ok") <> "PASS" Then
        Exit Function
       

    End If
    
    Wait 5
    
    fn_Callidus_Sales_Commission_Set_Default_Period = "PASS"
    
End Function

'**************************************************************************************************************************************************************************
'Validation of Position, Participants, Deposits and Incentives values in callidus application

'**************************************************************************************************************************************************************************
Function fn_Validations_New(RowCount,strValidate,objPosSummaryDictionary,objDictionary_Mapping,ObjRecordSet_SourceTable)
'    Set ApplicationTableObject = Test_Object("tbl_Participant_Summary")    
'    Set objDictionary_appTableCol = get_Application_Table_Columns(ApplicationTableObject)
'    
    'intSourceTable_RowLine = 3
    strcounter = 1
    'ObjRecordSet_SourceTable.MoveFirst
    SourceTableCols_arr = objDictionary_Mapping.keys
For i=0 To ubound(SourceTableCols_arr) Step 1
    
    strColHeader_SourceTable = SourceTableCols_arr(i)
    strColValue_SourceTable = ObjRecordSet_SourceTable.Fields(strColHeader_SourceTable)
    strColHeader_Application = objDictionary_Mapping.Item(SourceTableCols_arr(i))
    strColValue_ApplicationValue = objPosSummaryDictionary(strColHeader_Application)
    
    If strColHeader_SourceTable = "Employee Name" Then
        strColHeader_Application1 = objDictionary_Mapping.Item(SourceTableCols_arr(i))
        strColHeader_Application2=split(strColHeader_Application1, "||")
        strColHeader_Application_FirstName=strColHeader_Application2(0)
        strColHeader_Application_LastName=strColHeader_Application2(1)
        strColValue_ApplicationValue1=objPosSummaryDictionary(strColHeader_Application_FirstName)
        strColValue_ApplicationValue2=objPosSummaryDictionary(strColHeader_Application_LastName)
        strColValue_ApplicationValue=strColValue_ApplicationValue1&" "&strColValue_ApplicationValue2
    End If

    
    
    Select Case strValidate
    
    Case "Positions"
        
        If strColHeader_Application = "Manager Position Name" Then
            empno=split(strColValue_ApplicationValue,"-")
            strColValue_ApplicationValue=empno(1)
            'source value formating
            empno1=split(strColValue_SourceTable, "(")
            empno2=split(empno1(1), ")")
            strColValue_SourceTable=empno2(0)
        End If
            If isnull(strColValue_SourceTable) or isnull(strColValue_ApplicationValue) Then
                strColValue_SourceTable=""
                strColValue_ApplicationValue=""
            End If
        objSheet.cells(intSourceTable_RowLine,i+2).Value = strColValue_SourceTable
        If RowCount=0 Then
            
            objSheet.cells(intSourceTable_RowLine,i+2).Interior.ColorIndex = 3
                    res = strcounter& "No Records Returned in the Callidus application"
                    reporter.ReportEvent micFail, "validating the presence of column value for "&strColHeader_SourceTable, res
                    objSheet.cells(intSourceTable_RowLine,1).Value = objSheet.cells(intSourceTable_RowLine,1).Value & res & vblf
                    call FN_Update_Status("",res)
                    strcounter = strcounter + 1
        else
        If cstr(strColValue_SourceTable)= cstr(strColValue_ApplicationValue) Then
                    objSheet.cells(intSourceTable_RowLine,i+2).Interior.ColorIndex = 4
                    Reporter.ReportEvent micPass, "TABLEAU APPLICATION VALUE MATCHES WITH CALLIDUS APPLICATION","TABLEAU APPLICATION VALUE FOR " &strColHeader_SourceTable& " is " &strColValue_SourceTable& "MATCHES WITH CALLIDUS APPLICATION FOR" &strColValue_Application& " is " &strColValue_ApplicationValue 
                    strpass = 0
            
        ELSE
                    If strColValue_SourceTable="" Then
                        strColValue_SourceTable="BLANK"
                    End If
                    If strColValue_ApplicationValue="" Then
                        strColValue_ApplicationValue="BLANK"
                    End If
        
                    objSheet.cells(intSourceTable_RowLine,i+2).Interior.ColorIndex = 3
                    res = strcounter& ".validation of column value for "&strColHeader_SourceTable&" is failed. TABLEAU APPLICATION VALUE is "&strColValue_SourceTable&" and CALLIDUS APPLICATION value is "&strColValue_ApplicationValue
                    reporter.ReportEvent micFail, "validating the presence of column value for "&strColHeader_SourceTable, res
                    objSheet.cells(intSourceTable_RowLine,1).Value = objSheet.cells(intSourceTable_RowLine,1).Value & res & vblf
                    call FN_Update_Status("",res)
                    strcounter = strcounter + 1
            
        End If
    End If
    Case "Participants"
        If strColHeader_Application = "Employee ID" Then
        'source value formating
        
                strlen=len(strColValue_SourceTable)
                DO while(strlen<8)
                    strColValue_SourceTable="0"&strColValue_SourceTable
                    strlen=strlen+1
                    If strlen=8 Then
                        Exit do
                    End If
                Loop
        End If
        
        
        If isnull(strColValue_SourceTable) Then
            strColValue_SourceTable=""
        End If
        
        If isnull(strColValue_ApplicationValue) Then
            strColValue_ApplicationValue=""
        End If
    
            objSheet.cells(intSourceTable_RowLine,i+8).Value = strColValue_SourceTable
            'objSheet.cells(intSourceTable_RowLine,(Participant_Columns_Start+1)).Value = strColValue_SourceTable
            
        If cstr(strColValue_SourceTable)= cstr(strColValue_ApplicationValue) Then
                    'objSheet.cells(intSourceTable_RowLine,(Participant_Columns_Start+1)).Interior.ColorIndex = 4
                    objSheet.cells(intSourceTable_RowLine,i+8).Interior.ColorIndex = 4
                    Reporter.ReportEvent micPass, "TABLEAU APPLICATION VALUE MATCHES WITH CALLIDUS APPLICATION","TABLEAU APPLICATION VALUE FOR " &strColHeader_SourceTable& " is " &strColValue_SourceTable& "MATCHES WITH CALLIDUS APPLICATION FOR" &strColValue_Application& " is " &strColValue_ApplicationValue 
                            strpass = 0
            
        ELSE
        
                    If strColValue_SourceTable="" Then
                        strColValue_SourceTable="BLANK"
                    End If
                    If strColValue_ApplicationValue="" Then
                        strColValue_ApplicationValue="BLANK"
                    End If
                        
                    objSheet.cells(intSourceTable_RowLine,i+8).Interior.ColorIndex = 3
                    '.cells(intSourceTable_RowLine,(Participant_Columns_Start+1)).Interior.ColorIndex = 3
                    res = strcounter& ".validation of column value for "&strColHeader_SourceTable&" is failed. TABLEAU APPLICATION VALUE is "&strColValue_SourceTable&" and CALLIDUS APPLICATION value is "&strColValue_ApplicationValue
                    reporter.ReportEvent micFail, "validating the presence of column value for "&strColHeader_SourceTable, res
                    objSheet.cells(intSourceTable_RowLine,1).Value = objSheet.cells(intSourceTable_RowLine,1).Value & res & vblf
                    call FN_Update_Status("",res)
                    strcounter = strcounter + 1
            
        End If
'        Participant_Columns_Start=(Participant_Columns_Start)+1
'        If Participant_Columns_Start=objSheet.usedrange.columns.count Then
'            Exit Function
'        End If

    
     Case "Deposits"
         If strColHeader_Application = "Manager Position Name" Then
            empno=split(strColValue_ApplicationValue,"-")
            strColValue_ApplicationValue=empno(1)
            'source value formating
            empno1=split(strColValue_SourceTable, "(")
            empno2=split(empno1(1), ")")
            strColValue_SourceTable=empno2(0)
        End If
        
        If strColHeader_Application = "Create Date" Then
            strColValue_SourceTable = MonthName(Month(strColValue_SourceTable))&" "&Year(strColValue_SourceTable)
            If isempty(strColValue_ApplicationValue) Then
                strColValue_ApplicationValue="BLANK"
            else
                strColValue_ApplicationValue = MonthName(Month(strColValue_ApplicationValue))&" "&Year(strColValue_ApplicationValue)
            End If
            
        End If
        
        If strColHeader_Application = "GN1: TIA Earned Percentage" Then
            If instr(strColValue_ApplicationValue, "%")>1 Then
                New_strColValue_ApplicationValue=trim(replace(strColValue_ApplicationValue, "%", ""))
                strColValue_ApplicationValue=round(New_strColValue_ApplicationValue/100,2)
            End If
            
        End If
        
		   If strColHeader_Application = "Period" Then
            New_strColValue_ApplicationValue=split(strColValue_ApplicationValue, " ")
                strColValue_ApplicationValue=New_strColValue_ApplicationValue(1)
          
        End If
		
		   If strColHeader_Application = "Reason Code" Then
            If strColValue_ApplicationValue="" Then
            	strColValue_ApplicationValue="No reason!"
            End If
          
        End If
        
        If strColHeader_Application = "Value" Then
            strColValue_ApplicationValue= round(fn_Callidus_Remove_alpha_special_from_amount(strColValue_ApplicationValue),2)
            strColValue_SourceTable=round(strColValue_SourceTable,2)
        End If
            If isnull(strColValue_SourceTable) and strColValue_ApplicationValue="" Then
                strColValue_SourceTable=""
                strColValue_ApplicationValue=""
            End If
        objSheet.cells(intSourceTable_RowLine,i+11).Value = strColValue_SourceTable
        'objSheet.cells(intSourceTable_RowLine,(Participant_Columns_Start+1)).Value = strColValue_SourceTable
        If cstr(strColValue_SourceTable)= cstr(strColValue_ApplicationValue) Then
                    objSheet.cells(intSourceTable_RowLine,i+11).Interior.ColorIndex = 4
                    'objSheet.cells(intSourceTable_RowLine,(Participant_Columns_Start+1)).Interior.ColorIndex = 4
                    Reporter.ReportEvent micPass, "TABLEAU APPLICATION VALUE MATCHES WITH CALLIDUS APPLICATION","TABLEAU APPLICATION VALUE FOR " &strColHeader_SourceTable& " is " &strColValue_SourceTable& "MATCHES WITH CALLIDUS APPLICATION FOR" &strColValue_Application& " is " &strColValue_ApplicationValue 
                    strpass = 0
            
        ELSE
                    If strColValue_SourceTable="" Then
                        strColValue_SourceTable="BLANK"
                    End If
                    If strColValue_ApplicationValue="" Then
                        strColValue_ApplicationValue="BLANK"
                    End If
                    objSheet.cells(intSourceTable_RowLine,i+11).Interior.ColorIndex = 3
                    'objSheet.cells(intSourceTable_RowLine,(Participant_Columns_Start+1)).Interior.ColorIndex = 3
                    res = strcounter& ".validation of column value for "&strColHeader_SourceTable&" is failed. TABLEAU APPLICATION VALUE is "&strColValue_SourceTable&" and CALLIDUS APPLICATION value is "&strColValue_ApplicationValue
                    reporter.ReportEvent micFail, "validating the presence of column value for "&strColHeader_SourceTable, res
                    objSheet.cells(intSourceTable_RowLine,1).Value = objSheet.cells(intSourceTable_RowLine,1).Value & res & vblf
                    call FN_Update_Status("",res)
                    strcounter = strcounter + 1
            
        End If    

   Case "Position Group"
     
    
        
        If strColHeader_Application = "Value" Then
            strColValue_ApplicationValue= round(fn_Callidus_Remove_alpha_special_from_amount(strColValue_ApplicationValue),2)
            strColValue_SourceTable=round(strColValue_SourceTable,2)
        End If
            If isnull(strColValue_SourceTable) and strColValue_ApplicationValue="" Then
                strColValue_SourceTable=""
                strColValue_ApplicationValue=""
            End If
                   
        If isnull(strColValue_SourceTable) Then
            strColValue_SourceTable=""
        End If
        
        If isnull(strColValue_ApplicationValue) Then
            strColValue_ApplicationValue=""
        End If
        
        objSheet.cells(intSourceTable_RowLine,i+19).Value = strColValue_SourceTable
        'objSheet.cells(intSourceTable_RowLine,(Participant_Columns_Start+1)).Value = strColValue_SourceTable
        If RowCount=0 Then
                    objSheet.cells(intSourceTable_RowLine,i+19).Interior.ColorIndex = 3
                    'objSheet.cells(intSourceTable_RowLine,(Participant_Columns_Start+1)).Interior.ColorIndex = 3
                    res = strcounter& "No Records Returned in the Callidus application"
                    reporter.ReportEvent micFail, "validating the presence of column value for "&strColHeader_SourceTable, res
                    objSheet.cells(intSourceTable_RowLine,1).Value = objSheet.cells(intSourceTable_RowLine,1).Value & res & vblf
                    call FN_Update_Status("",res)
                    strcounter = strcounter + 1
        else
            If cstr(strColValue_SourceTable)= cstr(strColValue_ApplicationValue) Then
                        objSheet.cells(intSourceTable_RowLine,i+19).Interior.ColorIndex = 4
                        'objSheet.cells(intSourceTable_RowLine,(Participant_Columns_Start+1)).Interior.ColorIndex = 4
                        Reporter.ReportEvent     micPass, "TABLEAU APPLICATION VALUE MATCHES WITH CALLIDUS APPLICATION","TABLEAU APPLICATION VALUE FOR " &strColHeader_SourceTable& " is " &strColValue_SourceTable& "MATCHES WITH CALLIDUS APPLICATION FOR" &strColValue_Application& " is " &strColValue_ApplicationValue 
                        strpass = 0
                
            ELSE        
            
            
                        If strColValue_SourceTable="" Then
                        strColValue_SourceTable="BLANK"
                        End If
                        If strColValue_ApplicationValue="" Then
                        strColValue_ApplicationValue="BLANK"
                        End If
                        objSheet.cells(intSourceTable_RowLine,i+19).Interior.ColorIndex = 3
                        'objSheet.cells(intSourceTable_RowLine,(Participant_Columns_Start+1)).Interior.ColorIndex = 3
                        res = strcounter& ".validation of column value for "&strColHeader_SourceTable&" is failed. TABLEAU APPLICATION VALUE is "&strColValue_SourceTable&" and CALLIDUS APPLICATION value is "&strColValue_ApplicationValue
                        reporter.ReportEvent micFail, "validating the presence of column value for "&strColHeader_SourceTable, res
                        objSheet.cells(intSourceTable_RowLine,1).Value = objSheet.cells(intSourceTable_RowLine,1).Value & res & vblf
                        call FN_Update_Status("",res)
                        strcounter = strcounter + 1
                
            End If    
        End If     
        Case "Period"
            objSheet.cells(intSourceTable_RowLine,i+20).Value = strColValue_SourceTable
        If RowCount=0 Then
            objSheet.cells(intSourceTable_RowLine,i+20).Interior.ColorIndex = 3
                    res = strcounter& "No Records Returned in the Callidus application"
                    reporter.ReportEvent micFail, "validating the presence of column value for "&strColHeader_SourceTable, res
                    objSheet.cells(intSourceTable_RowLine,1).Value = objSheet.cells(intSourceTable_RowLine,1).Value & res & vblf
                    call FN_Update_Status("",res)
                    strcounter = strcounter + 1
        else
            If cstr(strColValue_SourceTable)= cstr(strColValue_ApplicationValue) Then
                        objSheet.cells(intSourceTable_RowLine,i+20).Interior.ColorIndex = 4
                        Reporter.ReportEvent micPass, "TABLEAU APPLICATION VALUE MATCHES WITH CALLIDUS APPLICATION","TABLEAU APPLICATION VALUE FOR " &strColHeader_SourceTable& " is " &strColValue_SourceTable& "MATCHES WITH CALLIDUS APPLICATION FOR" &strColValue_Application& " is " &strColValue_ApplicationValue 
                        strpass = 0
                
            ELSE        
            
            
                        If strColValue_SourceTable="" Then
                        strColValue_SourceTable="BLANK"
                        End If
                        If strColValue_ApplicationValue="" Then
                        strColValue_ApplicationValue="BLANK"
                        End If
                    
                        objSheet.cells(intSourceTable_RowLine,i+20).Interior.ColorIndex = 3
                        res = strcounter& ".validation of column value for "&strColHeader_SourceTable&" is failed. TABLEAU APPLICATION VALUE is "&strColValue_SourceTable&" and CALLIDUS APPLICATION value is "&strColValue_ApplicationValue
                        reporter.ReportEvent micFail, "validating the presence of column value for "&strColHeader_SourceTable, res
                        objSheet.cells(intSourceTable_RowLine,1).Value = objSheet.cells(intSourceTable_RowLine,1).Value & res & vblf
                        call FN_Update_Status("",res)
                        strcounter = strcounter + 1
                
            End If    
        End If
        
    End Select
Next
End  Function


Function fn_Callidus_Remove_alpha_special_from_amount(ByVal strvalue)

    If instr(1,strValue, "(") > 0 Then : strFlag = "-" : End If
    strValue = Replace(Replace(strValue,"(", ""), ")","")
    For j = 1 To len(strValue) Step 1
    str=mid(strValue,j,1)
    If asc(str)>64 and asc(str)<123 Then
        count=count+1
    Else 
        count=count+1
    End If
    
    If isnumeric(str) Then
        strnew=right(strValue,(len(strValue)-count+1))
        fn_Callidus_Remove_alpha_special_from_amount= strFlag & strnew
        Exit for
    End If
    
Next
    
    
End Function

'**************************************************************************************************************************************************************************
'Setting advance filter criteria using Test Object Method

'**************************************************************************************************************************************************************************

Function fn_Callidus_Sales_Commissions_Set_Advanced_Filter(strAdvancedFilterFieldNames, strAdvancedFilterFieldValues)
        
    fn_Callidus_Sales_Commissions_Set_Advanced_Filter = "FAIL"
    
    
    
    
    If Click_Object("ele_Advanced_Search") <> "PASS" Then
        Exit Function
    End If
    
    'Added by vijay to remove the existing filter options
    While(Test_Object("ele_Comm_Remove_Filter_Row").exist(2)=true)
        Click_Object("ele_Comm_Remove_Filter_Row")
    Wend
    
    Set objProcessUnit = Test_Object("txt_Process_Unit")
    If objProcessUnit.Exist(10) Then
        If Enter_Value_In_Edit_Field("txt_Process_Unit", "Global", "") <> "PASS" Then
            Exit Function
        End If
            
        If Click_Object("ele_PU_Global") <> "PASS" Then
            Exit Function
        End If    
        
    End If
    
    Set objBusinessUnit = Test_Object("txt_Business_Unit")
    If objBusinessUnit.Exist(5) Then
        If Enter_Value_In_Edit_Field("txt_Business_Unit", "BU Global", "") <> "PASS" Then
            Exit Function
        End If
        
        If Click_Object("ele_BU_Global") <> "PASS" Then
            Exit Function
        End If
    End  IF
    
    Wait 2
    
    If Instr(1, strAdvancedFilterFieldNames, "||") > 0 Then
        strFieldNamesCollection = Split(strAdvancedFilterFieldNames, "||")
        strFieldValuesCollection = Split(strAdvancedFilterFieldValues, "||")
    Else
        strFieldNamesCollection = Split(strAdvancedFilterFieldNames, ",")
        strFieldValuesCollection = Split(strAdvancedFilterFieldValues, ",")
    End If

    For intRowIndex = 0 To Ubound(strFieldNamesCollection) Step 1
        If intRowIndex = 1 Then
            Set objDeleteRowIcon = Test_Object("ele_Comm_Delete_Filter_Row")
            If Not objDeleteRowIcon.Exist(1) Then
                If Click_Object("ele_Comm_Add_Filter_Row") <> "PASS" Then
                    Exit Function
                End If
            End If
        ElseIf intRowIndex > 1 Then
            If Click_Object("ele_Comm_Add_Filter_Row") <> "PASS" Then
                Exit Function
            End If
        End If
        
        If strFieldNamesCollection(intRowIndex) = "Name" Then
        Call fn_Callidus_Filter_Selection(intRowIndex + 1, strFieldNamesCollection(intRowIndex), "contains", strFieldValuesCollection(intRowIndex))	
        else
        Call fn_Callidus_Filter_Selection(intRowIndex + 1, strFieldNamesCollection(intRowIndex), "equals", strFieldValuesCollection(intRowIndex))	
        End If
        
        
    Next

    If Click_Object("btn_Apply") <> "PASS" Then
        Exit Function
    End If
    
    Wait 5
    
    fn_Callidus_Sales_Commissions_Set_Advanced_Filter = "PASS"

End Function

Function fn_Callidus_Filter_Selection(intRowNumber,strFieldName, strComparionCondition, strFilterValue )
'    'added condition  2* 3 
'    If strFieldName="GA2" Then
'        If SetTo_Select_List_Value("lst_Comm_Select_FieldName", "xpath", "(//TABLE[@id='search_conditions']//SELECT)["& (intRowNumber*2)&"]", strFieldName) <> "PASS" Then
'            Exit Function
'        End If
'
'        If SetTo_Select_List_Value("lst_Comm_Select_Comparision_Condition", "xpath", "(//TABLE[@id='search_conditions']//SELECT)["& (intRowNumber*2)+1 &"]", strComparionCondition) <> "PASS" Then
'            Exit Function
'        End If
'    else
        'intRowNumber = intRowNumber + intRowNumber - 1
        If SetTo_Select_List_Value("lst_Comm_Select_FieldName", "xpath", "(//TABLE[@id='search_conditions']//SELECT)["& (intRowNumber*2) - 1 &"]", strFieldName) <> "PASS" Then
            Exit Function
        End If

        If SetTo_Select_List_Value("lst_Comm_Select_Comparision_Condition", "xpath", "(//TABLE[@id='search_conditions']//SELECT)["& (intRowNumber*2) &"]", strComparionCondition) <> "PASS" Then
            Exit Function
        End If
            
    'End If
    
    
    
    If strFieldName = "Employee ID"  Then
            IF SetTo_Enter_Value_In_Edit_Field("txt_Comm_Search_Value", "xpath", "(//TABLE[@id='search_conditions']//input)[" & intRowNumber & "]", strFilterValue) <> "PASS" Then
                Exit Function
            End If
        'code updated by vijay (added elseif condition to accomodate the text field values for Name
        ElseIf strFieldName = "Name" Then
            IF SetTo_Enter_Value_In_Edit_Field("txt_Comm_Search_Value", "xpath", "(//TABLE[@id='search_conditions']//input)[" & (intRowNumber*2) - 1 &"]", strFilterValue) <> "PASS" Then
            Exit Function
        End If
        
        ElseIf strFieldName="GA2" Then
            IF SetTo_Enter_Value_In_Edit_Field("txt_Comm_Search_Value", "xpath", "(//TABLE[@id='search_conditions']//input)[" & (intRowNumber*2) &"]", strFilterValue) <> "PASS" Then
            Exit Function
        End If
        
        ElseIf strFieldName="GB1: Is reporting format" Then
            If SetTo_Select_List_Value("lst_Comm_Select_Comparision_Condition", "xpath", "(//TABLE[@id='search_conditions']//SELECT)["& (intRowNumber*2)+1  &"]", strFilterValue) <> "PASS" Then
            Exit Function
        End If
        
        
            
        
'        ElseIf strFieldName="GN1: Month actual" Then
'            If SetTo_Click_Object("ele_Comm_Search_Value", "xpath", "(//TABLE[@id='search_conditions']//a)[" & (intRowNumber-1) & "]") <> "PASS" Then
'            Exit Function
'        End If
    Else
        If SetTo_Click_Object("ele_Comm_Search_Value", "xpath", "(//TABLE[@id='search_conditions']//a)[" & intRowNumber & "]") <> "PASS" Then
            Exit Function
        End If 
        
        If strFieldName = "Participant" Then 
            call SetTo_Enter_Value_In_Edit_Field("txt_Comm_Search_Value", "xpath", "//ul[@class='select2-results']//following::li[text()='Please enter 1 or more character']/../../div/input|//div[@class='select2-drop select2-display-none select2-with-searchbox select2-drop-active select2-drop-above']/div/input|(//div[@class='select2-drop select2-display-none select2-with-searchbox select2-drop-active'])[6]/div/input", strRequiredFildName)
            
'                If Instr(1, strFilterValue, "+") > 0 Then
'                    strFilterValueTemp = Split(strFilterValue, "+")
'                    strRequiredFildName = strFilterValueTemp(0)
'                    strFilterValue = strFilterValueTemp(1) & " " & strFilterValueTemp(2)
'                Else
'                    strFilterValueTemp = Split(strFilterValue, " ")
'                    strRequiredFildName = strFilterValueTemp(1)
'                End If
                
                call Enter_Value_In_Edit_Field("txt_Comm_Search_Value", strFilterValue, "")
                wait 2
                
                Call SetTo_Click_Object("ele_Comm_Search_Value", "xpath", "(//li[@class='select2-results-dept-0 select2-result select2-result-selectable select2-highlighted']/div)")
                Exit function
                
        Else
            strRequiredFildName = strFilterValue
            'Call SetTo_Click_Object("ele_Comm_Search_Value", "xpath", "//li[@class='select2-results-dept-0 select2-result select2-result-selectable select2-highlighted']/div")
            'Exit function
        End If
        
        
        If strFieldName = "Value" or strFieldName="GN1: Month actual" Then
            call SetTo_Enter_Value_In_Edit_Field("txt_Comm_Search_Value", "xpath", "//ul[@class='select2-results']//following::li[text()='Please enter 1 or more character']/../../div/input|//div[@class='select2-drop select2-display-none select2-with-searchbox select2-drop-active select2-drop-above']/div/input|(//div[@class='select2-drop select2-display-none select2-with-searchbox select2-drop-active'])[6]/div/input", strRequiredFildName)
            If Instr(1, strFilterValue, "+") > 0 Then
                strFilterValueTemp = Split(strFilterValue, "+")
                strRequiredFildName = strFilterValueTemp(0)
                strFilterValue = strFilterValueTemp(1)
            Else
                strFilterValueTemp = Split(strFilterValue, " ")
                strRequiredFildName = strFilterValueTemp(0)
                strFilterValue = strFilterValueTemp(1)
            End If
            
            If Enter_Value_In_Edit_Field("txt_Comm_Filter_FieldName_Value_Of_Value", strRequiredFildName, "") <> "PASS" Then
                Exit Function
            End If
            
            If Enter_Value_In_Edit_Field("txt_Comm_Search_Value", strFilterValue, "") <> "PASS" Then
                Exit Function
            End If
        Else
            
            If Enter_Value_In_Edit_Field("txt_Comm_Search_Value", strRequiredFildName, "") <> "PASS" Then
                Exit Function
            End If
            
'            ''Update by Vijay as the object Index changing - Discussed with Noor
'            IF SetTo_Enter_Value_In_Edit_Field("txt_Comm_Search_Value", "xpath", "(//*[text ()='Please enter 1 or more character'])[" & intRowNumber+1 & "]/../../div/input", strRequiredFildName) <> "PASS" Then
'                Exit Function
'            End If
        End If
    
    
        If SetTo_Select_List_Value("lst_Comm_Select_List_Value", "all items", strFilterValue & ".*", strFilterValue) <> "PASS" Then
            Exit Function
        End If
    End If
        
End Function

'**************************************************************************************************************************************************************************
'Callidus Navigation Function

'**************************************************************************************************************************************************************************

Function fn_Callidus_Sales_Commission_Reports_Navigation(strNavigationPath)
    
    fn_Callidus_Sales_Commission_Reports_Navigation = "FAIL"
    
    Set objPerformanceHome = Test_Object("pg_Sales_Performance_Home")
    Set objSAPCommissions = Test_Object("pg_SAP_Commissions")
    
    Do Until intCounter=60
        intCounter = intCounter + 1
        If  objPerformanceHome.Exist(2) Then
            If Click_Object("ele_Apps") <> "PASS" Then
                Exit Function
            End If
            
            If Click_Object("lnk_Commissions") <> "PASS" Then
                Exit Function
            End If
            Exit Do
        ElseIf objSAPCommissions.Exist(1) Then
            Set objHomeEle = Test_Object("ele_Comm_Home")
            If objHomeEle.Exist(1) Then
                If Click_Object("ele_Comm_Home") <> "PASS" Then
                    Exit Function
                End If
            Else
                If Click_Object("ele_Comm_Apps") <> "PASS" Then
                    Exit Function
                End If
                
                If Click_Object("lnk_Comm_Commissions") <> "PASS" Then
                    Exit Function
                End If
            End If
            Exit Do
        End  IF
    Loop

wait 3
    If Instr(1, strNavigationPath, "||") > 0 Then
        strRequiredNavigation = Split(strNavigationPath, "||")
    ElseIf Instr(1, strNavigationPath, "-->") > 0 Then
        strRequiredNavigation = Split(strNavigationPath, "-->")
    ElseIf Instr(1, strNavigationPath, "->") > 0 Then
        strRequiredNavigation = Split(strNavigationPath, "->")
    End If
    
    Do Until intObjectFoundCounter=90
        intObjectFoundCounter = intObjectFoundCounter + 1
        If isempty(strRequiredNavigation) Then
            strRequiredNavigation=strNavigationPath
            Set objNavigationObject = SetTo_Object("lnk_Comm_Generic_Object", "name",strRequiredNavigation)
            If SetTo_Click_Object("lnk_Comm_Generic_Object", "name", strRequiredNavigation) <> "PASS" Then
                Exit Function
            End If
            fn_Callidus_Sales_Commission_Reports_Navigation = "PASS"
            wait 2
            Exit Function
        End If
        Set objNavigationObject = SetTo_Object("lnk_Comm_Generic_Object", "name",strRequiredNavigation(0))
        If objNavigationObject.Exist(1) Then
            Exit Do
        Else
            Wait 1
        End If
    Loop
    
    For intNavigationIndex = 0 To Ubound(strRequiredNavigation) Step 1
        If SetTo_Click_Object("lnk_Comm_Generic_Object", "name", strRequiredNavigation(intNavigationIndex)) <> "PASS" Then
            Exit Function
        End If
    Next
    
    Wait 15
    fn_Callidus_Sales_Commission_Reports_Navigation = "PASS"
    
End Function


Function fn_Callidus_Validate_Calendar(objDictionary_Mapping,ObjRecordSet_SourceTable)

        call fn_Callidus_Sales_Commission_Reports_Navigation("Global Values||calendars")            
        
            If Test_Object("lnk_MF_Calender").Exist(10)Then
                Test_Object("lnk_MF_Calender").highlight
                Test_Object("lnk_MF_Calender").click
                Test_Object("lnk_MF_Decade").click
                Test_Object("lnk_MF_Year").click
                wait 1
            End If
        icount=icount+1
        wait 2

        
        Set objTableClPosition=Test_object("tbl_MF_Fiscal_Calendar")
    
        strColValue_SourceTable = ObjRecordSet_SourceTable.Fields("Fiscal Year Name").value
                
                            
        row=objTableClPosition.getrowwithcelltext("year")
                    
        If row>0 Then
                        val=objTableClPosition.getcelldata(row,1)
                        name=split(val,"(")
                        
                if trim(cstr(name(0)))=trim(strColValue_SourceTable) then
                            
                    objSheet.cells(intSourceTable_RowLine,i+20).Value = strColValue_SourceTable
                            If RowCount=0 Then
'                                objSheet.cells(intSourceTable_RowLine,i+2).Interior.ColorIndex = 3
'                                        res = strcounter& "No Records Returned in the Callidus application"
'                                        reporter.ReportEvent micFail, "validating the presence of column value for "&strColHeader_SourceTable, res
'                                        objSheet.cells(intSourceTable_RowLine,1).Value = objSheet.cells(intSourceTable_RowLine,1).Value & res & vblf
'                                        call FN_Update_Status("",res)
'                                        strcounter = strcounter + 1
'                            else
'                                If cstr(strColValue_SourceTable)= cstr(strColValue_ApplicationValue) Then
                                            objSheet.cells(intSourceTable_RowLine,i+20).Interior.ColorIndex = 4
                                            Reporter.ReportEvent micPass, "TABLEAU APPLICATION VALUE MATCHES WITH CALLIDUS APPLICATION","TABLEAU APPLICATION VALUE FOR " &strColHeader_SourceTable& " is " &name(0)& "MATCHES WITH CALLIDUS APPLICATION FOR" &strColValue_Application& " is " &strColValue_ApplicationValue 
                                            strpass = 0
                                    
                                ELSE        
                                
                                
                                            If strColValue_SourceTable="" Then
                                            strColValue_SourceTable="BLANK"
                                            End If
                                            If name(0)="" Then
                                            name(0)="BLANK"
                                            End If
                                        
                                            objSheet.cells(intSourceTable_RowLine,i+20).Interior.ColorIndex = 3
                                            res = strcounter& ".validation of column value for "&strColHeader_SourceTable&" is failed. TABLEAU APPLICATION VALUE is "&strColValue_SourceTable&" and CALLIDUS APPLICATION value is "&strColValue_ApplicationValue
                                            reporter.ReportEvent micFail, "validating the presence of column value for "&strColHeader_SourceTable, res
                                            objSheet.cells(intSourceTable_RowLine,1).Value = objSheet.cells(intSourceTable_RowLine,1).Value & res & vblf
                                            call FN_Update_Status("",res)
                                            strcounter = strcounter + 1
                                    
                                End If    
                End If
        End If 


                
val2=objTableClPosition.getcelldata(row,2)
                
        if trim(cstr(name(0)))=trim(val2) then
                    
            objSheet.cells(intSourceTable_RowLine,i+20).Value = strColValue_SourceTable
                If RowCount=0 Then
                
                    objSheet.cells(intSourceTable_RowLine,i+20).Interior.ColorIndex = 4
                    Reporter.ReportEvent micPass, "TABLEAU APPLICATION VALUE MATCHES WITH CALLIDUS APPLICATION","TABLEAU APPLICATION VALUE FOR " &strColHeader_SourceTable& " is " &strColValue_SourceTable& "MATCHES WITH CALLIDUS APPLICATION FOR" &strColValue_Application& " is " &name(0) 
                    strpass = 0
                                    
                 ELSE        
                    If strColValue_SourceTable="" Then
                    strColValue_SourceTable="BLANK"
                    End If
                    If strColValue_ApplicationValue="" Then
                    strColValue_ApplicationValue="BLANK"
                    End If
                        
                        objSheet.cells(intSourceTable_RowLine,i+20).Interior.ColorIndex = 3
                        res = strcounter& ".validation of column value for "&strColHeader_SourceTable&" is failed. TABLEAU APPLICATION VALUE is "&strColValue_SourceTable&" and CALLIDUS APPLICATION value is "&strColValue_ApplicationValue
                        reporter.ReportEvent micFail, "validating the presence of column value for "&strColHeader_SourceTable, res
                        objSheet.cells(intSourceTable_RowLine,1).Value = objSheet.cells(intSourceTable_RowLine,1).Value & res & vblf
                        call FN_Update_Status("",res)
                        strcounter = strcounter + 1
                End If    
        End If
        
        Test_Object("ele_Comm_Apps").click
        
        Test_Object("lnk_Comm_Commissions").click
        
        wait(10)
        
End Function
'*************************************************************End of Local Functions****************************************************************************************



'***********************************************************************************************************************************************************************************
Function get_Source_Base_Mapping_field_Names(strMappingFileName, strMappingSheetName, strSrcColumnName, strBseColumnName)
'    strPath_arr = split(strMappingFilePath,"\")
'    strFileName = strPath_arr(ubound(strPath_arr))
    if instr(strMappingFileName,".xls")>0 then 
'        strFolderLocation = mid(strMappingFilePath,1,len(strMappingFilePath)-len(strFileName)-1)
        strFolderLocation = strpath&"Test_Data\Callidus\1C_Mapping"
'        call QCGetResource(strFileName,strFolderLocation)
        call QCGetResource(strMappingFileName,strFolderLocation)
        
        Set objDictionary_MappingFile = CreateObject("Scripting.Dictionary")
        objDictionary_MappingFile.RemoveAll
'        set objconnection_Excel = createobject("ADODB.connection")
'        objconnection_Excel.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" &strFolderLocation&"\"& strMappingFileName & ";Extended Properties= ""Excel 8.0;HDR=Yes;IMEX=1"""
'        objconnection_Excel.Open
        Call get_Excel_DB_Connection(strFolderLocation&"\"& strMappingFileName)
        strQuery_Excel = "Select * From [" & strMappingSheetName & "$] "
'        Set objRecord_Excel = objconnection_Excel.Execute(strQuery_Excel)
        Set objRecord_Excel = execute_Excel_DB_SQL_Query(strQuery_Excel)
        Do  Until objRecord_Excel.EOF
            src_Col = objRecord_Excel.Fields(strSrcColumnName)
            bse_Col = objRecord_Excel.Fields(strBseColumnName)
            If src_Col = "NA" or src_Col = "N/A" or src_Col = "TBC*" or src_Col = ""or isNull(src_Col) or bse_Col = "NA" or bse_Col = "N/A" or bse_Col = "TBC*" or bse_Col = "" or isNull(bse_Col) Then
            else
                objDictionary_MappingFile.Add src_Col, bse_Col
            End If
            
            objRecord_Excel.MoveNext
        Loop
        set get_Source_Base_Mapping_field_Names = objDictionary_MappingFile
'        objconnection_Excel.Close
'        Set objconnection_Excel = nothing
        close_Excel_DB_Connection()
    Else
        reporter.ReportEvent micFail, "Reading Mapping", "Reading Mapping file failed. Please enter the correct file name" &strMappingFileName
    End If
End Function

'*****************************************************************************************************************************************************************************************
'Function Name    : Add_Tabl1_BaseTable_Headers_In_Output_File
'Description           : Add Source and BaseTable name and column headers in output file
'Arguments            : ObjRecordSet_Source - Record set object of 1st sql query
'                        objDictionary_Mapping - mapping column names dictionary object
'                        strSourceTable_Name - Source table name
'                        strBaseTable_Name - Base table name
'Return value        : 
'Example                : Call Add_Tabl1_BaseTable_Headers_In_Output_File(ObjRecordSet_Source, objDictionary_Mapping)
'******************************************************************************************************************************************************************************************
Function Add_Source_Base_Headers_In_Output_File(ObjRecordSet_SourceTable,objDictionary_Mapping, strSourceTable_Name, strBaseTable_Name)
    totalRecords_SourceTable=0
    Do Until ObjRecordSet_SourceTable.EOF
        totalRecords_SourceTable = totalRecords_SourceTable+1
        ObjRecordSet_SourceTable.MoveNext
        BaseTableStartRow = totalRecords_SourceTable+4
    Loop
    
    objSheet.cells(1,1).Value = strSourceTable_Name
    objSheet.cells(BaseTableStartRow,1).Value = strBaseTable_Name
    
    intRecCounter = 1
    intSourceTable_RowLine = 2
    intBaseTable_RowLine = BaseTableStartRow+1
    
    SourceTableCols_arr = objDictionary_Mapping.Keys
    'add Headers
    For i=0 To ubound(SourceTableCols_arr) Step 1
        If intRecCounter = 1 Then
            objSheet.cells(intSourceTable_RowLine,1).Value = "REMARKS"  
            objSheet.Columns("A").ColumnWidth = 80 
            objSheet.cells(intSourceTable_RowLine,i+2).Value = SourceTableCols_arr(i)
            objSheet.cells(intBaseTable_RowLine,i+2).Value = objDictionary_Mapping.item(SourceTableCols_arr(i))
            intRecCounter =intRecCounter+1 
        else
            objSheet.cells(intSourceTable_RowLine,i+2).Value = SourceTableCols_arr(i)
            objSheet.cells(intBaseTable_RowLine,i+2).Value = objDictionary_Mapping.item(SourceTableCols_arr(i))
        End If
    Next
    
    intSourceTable_RowLine = intSourceTable_RowLine+1
    intBaseTable_RowLine = intBaseTable_RowLine+1
End Function



Function create_1A1C_Output_File_UAT(strIntegrationName,strSheetName)

    strOutputFilePath = strpath&"Test_Result_Log\1A1C_FUNCTIONAL_PREPROD\"&strIntegrationName&"\"&strIntegrationName&"_Output.xlsx"
    Environment.Value("Output_Result_Reference_File") = strOutputFilePath
    Call create_Output_Folder_Location(strOutputFilePath)

    Set objfso = CreateObject("Scripting.FileSystemObject")
    Set objExcel_OP = CreateObject("Excel.Application")
    objExcel_OP.visible = True
    objExcel_OP.DisplayAlerts = False
    
    If not objfso.FileExists(strOutputFilePath) Then
        Set ObjWorkbook = objExcel_OP.Workbooks.Add
        On Error Resume Next
        objExcel_OP.ActiveWorkbook.SaveAs strOutputFilePath
        On Error Resume Next
        Set objSheet = ObjWorkbook.Sheets(strSheetName)
        If Err <> 0 Then
            Set objSheet = ObjWorkbook.Sheets.Add
               objSheet.Name = strSheetName
           End If
           Err.Clear
           Set ObjWorkbook = objExcel_OP.Workbooks.Open(strOutputFilePath)
           On Error Resume Next
        Set objSheet = ObjWorkbook.Sheets(strSheetName)
        If Err <> 0 Then
            Set objSheet = ObjWorkbook.Sheets.Add
               objSheet.Name = strSheetName
        End If
           Err.Clear
           objSheet.Activate
           ObjWorkbook.Save strOutputFilePath
           set create_1A1C_Output_File = objSheet
       Else
           Set ObjWorkbook = objExcel_OP.Workbooks.Open(strOutputFilePath)
           On Error Resume Next
        Set objSheet = ObjWorkbook.Sheets(strSheetName)
        If Err <> 0 Then
            Set objSheet = ObjWorkbook.Sheets.Add
               objSheet.Name = strSheetName
           else
               objSheet.UsedRange.EntireRow.Delete
           End If
           Err.Clear
           objSheet.Activate
           ObjWorkbook.Save
           set create_1A1C_Output_File = objSheet
    End If
End Function

Function fnSqlQuery(objExcel1,strfilename,strFolderLocation,strOutputSheetName,SheetName,objDictionary_Mapping,strperiod,empid)
                Set objWorkbook1 = objExcel1.Workbooks.Open (strfilename)
                Set objWorksheet1 = objWorkbook1.Worksheets(1)
                    objWorksheet1.Name = SheetName
                    objWorkbook1.SaveAs (strFolderLocation&"\"&strOutputSheetName&".xls")
                    'objWorkbook1.save
                    objWorkbook1.close
                    objExcel1.Quit
                    objconnection_Excel.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source="&strFolderLocation&"\" &strOutputSheetName&".xls;Extended Properties= ""Excel 8.0;HDR=Yes;IMEX=1"""
                    objconnection_Excel.Open
                    'Execute the Excel query
                    'fnSqlQuery = "select Top "&parameter("P_Num_Rec_Per_Employee")&" * from ["&SheetName&"$] Where [Period Start Date] IN ("&strperiod&") and [Employee ID]="&empid      
                    fnSqlQuery= "select Top "&parameter("P_Num_Rec_Per_Employee")&" * from ["&SheetName&"$] Where [Period Start Date] IN ("&strperiod&")"
End Function

Function fnSqlQuery1(objExcel1,strfilename,strFolderLocation,strOutputSheetName,SheetName,objDictionary_Mapping,strperiod1,empid1)
                    objconnection_Excel.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source="&strFolderLocation&"\" &strOutputSheetName&".xls;Extended Properties= ""Excel 8.0;HDR=Yes;IMEX=1"""
                    objconnection_Excel.Open
                    'Execute the Excel query
                    fnSqlQuery1 = "select Top "&parameter("P_Num_Rec_Per_Employee")&" * from ["&SheetName&"$] Where [Period Start Date] IN ("&strperiod1&") and [Employee ID]="&empid1
                    
                    
End Function

Public Function Click_Object(strObjectLogicalName)
' This function 
   Click_Object ="FAIL"
   Set objTemp = Test_Object(strObjectLogicalName)
   If  objTemp.Exist(40) Then
	      objTemp.Click
		  Click_Object ="PASS"
	Else
		Call LogResult_And_CaptureImage(strObjectLogicalName,"Click on Object: " & strObjectLogicalName,"FAIL","'" & strObjectLogicalName &"' object does not exist.")
	End If
	CntClick=CntClick+1
End Function

Function fn_Convert_Xlsx_TO_Xls(StrCalenderName,strFolderLocation,StrCalenderFile) 

Set objExcel1 = CreateObject("Excel.Application")
	Set objWorkbook1 = objExcel1.Workbooks.Open (StrCalenderName)
					objWorkbook1.SaveAs (strFolderLocation&"\"&StrCalenderFile)
					'objWorkbook1.save
					objWorkbook1.close
					objExcel1.Quit

End Function

Function fn_Convert_Xlsx_TO_Xls(ByVal vFilePath) 
path=Left(vFilePath, Len(vFilePath) - 4) & ".xlsx"
'fnDeleteExcelFile(path)
 Dim xlApp 
 'Dim iSeconds as Integer
 iSeconds = 20
 Set xlApp = CreateObject("excel.application") 
 With xlApp.Workbooks.Open(vFilePath) 'open file 
 		xlApp.visible=False
 		xlApp.application.DisplayAlerts=False
  With .ActiveSheet 
   'Set a freeze under column 1 so that the header is always present at the top 
   .Range("A7").Select 
   xlApp.ActiveWindow.FreezePanes = True 
  End With 
  'Sleep iSeconds * 1000 '20 * 1 second
  wait (10)
 ' .SaveAs Left(vFilePath, Len(vFilePath) - 4) & ".xlsx", 51 '-4143 '-4143=xlWorkbookNormal
  .SaveAs Left(vFilePath, Len(vFilePath) - 5) &  ".xls", -4143
  '.SaveAs Left(vFilePath, Len(vFilePath) - 4) &  ".xls", 56
  .Close True 'save and close 
 End With 
 xlApp.Quit 
 Set xlApp = Nothing 
 fn_Convert_Xlsx_TO_Xls=Left(vFilePath, Len(vFilePath) - 5) &  ".xls"
 'fn_Convert_CSV_TO_Xlsx=vFilePath
End Function

Function execute_SQL_Query2(strSQLQuery,objconnection_Excel)
	'Set ObjRecordSet = objConnection.Execute(strSQLQuery)
	Set objRecord_Excel = objconnection_Excel.Execute(strSQLQuery)
	If objRecord_Excel.EOF = true Then
		Reporter.ReportEvent micFail, "Execute SQL Query "&sqlQuery, "No result found for the sql query "&sqlQuery
		Set execute_SQL_Query2 = objRecord_Excel
	Else
		objRecord_Excel.MoveFirst	
		Set execute_SQL_Query2 = objRecord_Excel
	End If
End Function

Function execute_SQL_Query1(strSQLQuery)
	'Set ObjRecordSet = objConnection.Execute(strSQLQuery)
	Set objRecord_Excel = objconnection_Excel.Execute(strSQLQuery)
	If objRecord_Excel.EOF = true Then
		Reporter.ReportEvent micFail, "Execute SQL Query "&sqlQuery, "No result found for the sql query "&sqlQuery
		Set execute_SQL_Query1 = ObjRecordSet
	Else
		objRecord_Excel.MoveFirst	
		Set execute_SQL_Query1 = objRecord_Excel
	End If
End Function



