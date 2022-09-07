'*********************************Getting FIELD NAMES from the Anaplan Application*************************************************************
'Header names
Fieldname=parameter("p_First_Field_Name")
'getting the complete header names using the below function
Target_Field_Name= fn_Anaplan_Field_names(Fieldname)
'splitting the header names got from the application and comparing it with the header names passed using the parameter
source_Field_Name=split(parameter("p_Source_Fields_names"),";")
For i = 0 To ubound(source_Field_Name) Step 1
	If  instr(1,source_Field_Name(i),Target_Field_Name)<0   Then
		LogResult_And_CaptureImage "br_Anaplan_Login", "Column Header"&source_Field_Name(i)& "Name Not found", "FAIL", "Source and column Header Names do not Match."
		Else
		LogResult_And_CaptureImage "br_Anaplan_Login", "Column Header" &source_Field_Name(i)& "Name found", "PASS", "Source and column Header Names Match."
	End If
Next

'*********************************Export file in xlsx format*********************************************************************************
'changing the extension from xls or csv to xlsx
'export_file_path1 = strLocalPath & "Test_Data\Callidus\" & parameter("p_Integration_name") & "\" & parameter("p_export_file_name") 
export_file_path1 = strLocalPath & "\Test_Data\Callidus\" & parameter("p_Integration_name") 
create_Output_Folder_Location(export_file_path1)
export_file_path=export_file_path1& "\" & parameter("p_export_file_name") 

'xlsx_path=Left(parameter("p_excel_path"), Len(parameter("p_excel_path")) - 4) & ".xlsx"
xls_path=Left(export_file_path, Len(export_file_path) -4) & ".xls"

If instr(export_file_path,".csv")>0 Then

	call f_csvexport()
	'wait(100)
	
	Do
		If Test_Object("ele_Save_Export_DefinitionExport").exist = "True" then 
			wait 5
			Else
			Exit Do
		End If
	Loop until Test_Object("ele_Save_Export_DefinitionExport").exist = "False"
	
	If parameter("p_view_name")="EXPORT_Account_Attributes" or parameter("p_view_name")="Sales_Account_Data_from_SFDC" Then
		wait 50
	End If
	
	
	
	'Deleting any existing files
	fnDeleteExcelFile(xls_path)
	'Converting from xls,csv to xlsx format
	fn_TO_Xlsx(export_file_path)
	parameter("p_xls_outputfile")=xls_path
Else  
	'export
	call f_excelexport()
	'Deleting any existing files
	'fnDeleteExcelFile(xls_path)
	'Converting from xls to xlsx format
	'fn_TO_Xlsx(export_file_path)
	parameter("p_xls_outputfile")=xls_path
End If
'*****************************************************************************End**************************************************************************

'Function for exporting anaplan application datat to xls 
Function f_excelexport()
	excelpath=export_file_path
	fnDeleteExcelFile(excelpath)
	Browser("br_Anaplan_Login").Page("pg_Anaplan_MF_FY19_HUB").Webbutton("btn_Data").Click
	Browser("br_Anaplan_Login").Page("pg_Anaplan_MF_FY19_HUB").Webelement("uniqName_35_74_text").Click
	Browser("br_Anaplan_Login").Page("pg_Anaplan_MF_FY19_HUB").WebList("widget_fileType").highlight
	Browser("br_Anaplan_Login").Page("pg_Anaplan_MF_FY19_HUB").WebEdit("txt_filetype").Set "Excel (.xls)"
	Browser("br_Anaplan_Login").Page("pg_Anaplan_MF_FY19_HUB").Webbutton("btn_Run_Export").Click
'	Window("wnd_Google Chrome").Dialog("dlg_GC_SaveAs").WinEdit("txt_GC_File_name").Set excelpath
'	Window("wnd_Google Chrome").Dialog("dlg_GC_SaveAs").WinButton("btn_GC_Save").Click
	'Window("wnd_Google Chrome").Dialog("dlg_GC_SaveAs").WinEdit("txt_A_File_name").Set excelpath
	Test_object("txt_A_File_name").set excelpath
	Window("wnd_Google Chrome").Dialog("dlg_GC_SaveAs").WinButton("win_crl_Save").Click

End Function

'function to export file into csv format
Function f_csvexport()
	excelpath=export_file_path
	fnDeleteExcelFile(excelpath)
	Browser("br_Anaplan_Login").Page("pg_Anaplan_MF_FY19_HUB").Webbutton("btn_Data").Click
	Browser("br_Anaplan_Login").Page("pg_Anaplan_MF_FY19_HUB").Webelement("uniqName_35_74_text").Click
	Browser("br_Anaplan_Login").Page("pg_Anaplan_MF_FY19_HUB").WebElement("anaplanDialog_title").Click
	Browser("br_Anaplan_Login").Page("pg_Anaplan_MF_FY19_HUB").WebList("widget_fileType").highlight
	Browser("br_Anaplan_Login").Page("pg_Anaplan_MF_FY19_HUB").WebEdit("txt_filetype").Set "Comma Separated Values (.csv)"
	Browser("br_Anaplan_Login").Page("pg_Anaplan_MF_FY19_HUB").Webbutton("btn_Run_Export").Click
'	Window("wnd_Google Chrome").Dialog("dlg_GC_SaveAs").WinEdit("txt_A_File_name").Set excelpath
'	Window("wnd_Google Chrome").Dialog("dlg_GC_SaveAs").WinButton("btn_Save").Click
	'Window("wnd_Google Chrome").Dialog("dlg_GC_SaveAs").WinEdit("txt_A_File_name").Set excelpath
	Test_object("txt_A_File_name").set excelpath
	Window("wnd_Google Chrome").Dialog("dlg_GC_SaveAs").WinButton("win_crl_Save").Click
End Function

Function f_XLSXexport()
	excelpath=export_file_path
	fnDeleteExcelFile(excelpath)
	Browser("br_Anaplan_Login").Page("pg_Anaplan_MF_FY19_HUB").Webbutton("btn_Data").Click
	Browser("br_Anaplan_Login").Page("pg_Anaplan_MF_FY19_HUB").Webelement("uniqName_35_74_text").Click
	Browser("br_Anaplan_Login").Page("pg_Anaplan_MF_FY19_HUB").WebElement("anaplanDialog_title").Click
	Browser("br_Anaplan_Login").Page("pg_Anaplan_MF_FY19_HUB").WebList("widget_fileType").highlight
	Browser("br_Anaplan_Login").Page("pg_Anaplan_MF_FY19_HUB").WebEdit("txt_filetype").Set "Excel Open XML (.xlsx)"
	Browser("br_Anaplan_Login").Page("pg_Anaplan_MF_FY19_HUB").Webbutton("btn_Run_Export").Click
	'Window("wnd_Google Chrome").Dialog("dlg_GC_SaveAs").WinEdit("txt_A_File_name").Set excelpath
	Test_object("txt_A_File_name").set excelpath
	Window("wnd_Google Chrome").Dialog("dlg_GC_SaveAs").WinButton("btn_Save").Click
End Function




'function to convert xls file to xlsx file
Function fn_TO_Xlsx(ByVal vFilePath) 
path=Left(vFilePath, Len(vFilePath) - 4) & ".xls"
fnDeleteExcelFile(path)
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
  .SaveAs Left(vFilePath, Len(vFilePath) - 4) &  ".xls", -4143
  '.SaveAs Left(vFilePath, Len(vFilePath) - 4) &  ".xls", 56
  .Close True 'save and close 
 End With 
 xlApp.Quit 
 Set xlApp = Nothing 
End Function


Function fnDeleteExcelFile(strPath)
On error resume next
	strTempResultPath=strPath
	'strTempResultPath="C:\FAST_Test_Automation\Execution_Output_Data\Execution_Test_Details_Temp.xlsx"

	set objFSO = createobject("Scripting.FileSystemObject")
	if objFSO.FileExists(strTempResultPath) then
		'msgbox "File Exist"
		objFSO.DeleteFile(strTempResultPath)
		set objFSO =  nothing 
	end if
End Function
	

 Function fn_FilterSetUp(strLogicalnameofColumn)
 Test_Object(strLogicalnameofColumn).click
 wait 5
 if Browser("br_Anaplan_Login").Page("pg_Anaplan_MF_FY19_HUB").webbutton("html id:=uniqName_5_13").exist(10) then

	Browser("br_Anaplan_Login").Page("pg_Anaplan_MF_FY19_HUB").webbutton("html id:=uniqName_5_13").highlight

	Browser("br_Anaplan_Login").Page("pg_Anaplan_MF_FY19_HUB").webbutton("html id:=uniqName_5_13").click


	if strLogicalnameofColumn="ele_Accelerator_is_deactivated" then 
		'call Select_CheckBox_anaplan ("chk_accelerator_is_deactivated", "ON")
		call Select_CheckBox ("chk_accelerator_is_deactivated", "ON")
	else
		call Select_Item_From_ListBox("lst_select_Filter", "is not blank")
	end if

wait 3

Click_Object("btn_Filter_OK")

End  if
 End Function

Function fn_TO_Xls(ByVal vFilePath) 
path=Left(vFilePath, Len(vFilePath) - 4) & ".csv"
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
  .SaveAs Left(vFilePath, Len(vFilePath) - 4) &  ".xls", -4143
  '.SaveAs Left(vFilePath, Len(vFilePath) - 4) &  ".xls", 56
  .Close True 'save and close 
 End With 
 xlApp.Quit 
 Set xlApp = Nothing 
End Function

Function Select_CheckBox_anaplan (strObjectLogicalName, Val)
   Select_CheckBox_anaplan ="FAIL"
   Set objTemp = Test_Object(strObjectLogicalName)
	If  objTemp.Exist Then
         If  UCase(Val) = "ON" Then
                 	 If ucase(objTemp.GetROProperty("Value")) <> Val Then
					objTemp.Set "On"
					Select_CheckBox_anaplan ="PASS"
					else
					Select_CheckBox_anaplan ="PASS"
				 End If
		ElseIf UCase(Val) = "OFF" Then
				If ucase(objTemp.GetROProperty("Value")) <> Val Then
					objTemp.Set "Off"	
					Select_CheckBox_anaplan ="PASS"
					else
					Select_CheckBox_anaplan ="PASS"
				 End If
		Else
				Call LogResult_And_CaptureImage(strObjectLogicalName,"Select Checkbox: " & strObjectLogicalName,"FAIL", _
				"The specified value '"& Val & "' is not defined for the object '" & strObjectLogicalName & "'.")
		End If
	Else
			Call LogResult_And_CaptureImage(strObjectLogicalName,"Select Checkbox: " & strObjectLogicalName,"FAIL" _
			,"'" & strObjectLogicalName &"' object does not exist.")
	End If
End Function
