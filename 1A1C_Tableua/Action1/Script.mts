'MBT_UFT
'Call Fn_Component_Start()
systemutil.CloseProcessByName("Chrome.exe")
call Launch_Application(Parameter("P_URL_In"))
'new code arpil 2021 by dhanesh
WAIT 10
set WshShell = CreateObject("WScript.Shell")
WshShell.SendKeys "{ESC}"

'added by Suprita Nov25,2020
'strTestInstance = Parameter("P_TestInstance_In")
'If strTestInstance = "FY20" Then
'	Browser("br_Tableau_Server").Sync
	'Enter username and password and click on Sign in
	If Test_Object("txt_crl_username").exist(30) Then
	   call Enter_Value_In_Edit_Field("txt_crl_username",Parameter("P_UserName_In"),"No")
	   call Enter_Value_In_Edit_Field("txt_crl_password",Parameter("P_Password_In"),"No")
	   Click_Object("btn_crl_Sign_In")
	'End If
Else
	If Test_Object("txt_Ecom_User_ID").exist(30) Then
	   call Enter_Value_In_Edit_Field("txt_Ecom_User_ID",Parameter("P_UserName_In"),"No")
	   call Enter_Value_In_Edit_Field("txt_Ecom_Password",Parameter("P_Password_In"),"No")
	   Click_Object("btn_Login_sso")
	   wait(2)
	End If	
End If

'Select the correct environment
str_tableau_env=parameter("P_Tableau_Env")
If Test_Object("ele_crl_Select_Site").exist(30) Then
Call SetTo_Object("ele_CommissionsPreProd", "innertext", str_tableau_env)
Call Click_Object("ele_CommissionsPreProd")
End If

'Click on the Link Back to all views
If Test_Object("lnk_Back_To_All_Views").exist(5) Then
	Click_Object("lnk_Back_To_All_Views")
End If
'Clicl on user define reports 
If Test_Object("lnk_UDR_User_Defined_Reports").exist(5) Then
	Test_Object("lnk_UDR_User_Defined_Reports").click
End If

'Check whether the user is navigate to the user defined reports page
if Test_Object("pg_User_Defined_Reports").exist(60) then
	LogResult_And_CaptureImage "pg_User_Defined_Reports", "Login Authorization", "PASS", "Login to Tableau is successful."
else
	LogResult_And_CaptureImage "pg_User_Defined_Reports", "Login Authorization", "FAIL", "Login to Tableau is failed."
	ExitTest
End if	
'Call Fn_Component_End()


Function Launch_Application(strPath)
   'Call Close_All_Browser_Instances_SAP()
   Set g_objFileSystemObject = CreateObject("Scripting.FileSystemObject")
   ' If the Application type is not a web based application then, verify for the existance of the executable file
   If left(strPath,4)<> "http" Then
		   If g_objFileSystemObject.FileExists(Trim(strPath)) Then
			   InvokeApplication strPath
		   Else
				Reporter.ReportEvent micFail,"Launch_Application",strPath & " path not found. "
                Exit Function
		   End If
   Else
			
			' If Browser to be launched is Internet Explorer
			Set WshShell = CreateObject("WScript.Shell")
			If strBrowser_Type="InternetExplorer" Then
				Call WshShell.Run("iexplore.exe " & strPath, 3, false)
            ' If Browser to be launched is Mozilla Firefox				
			Elseif strBrowser_Type="MozillaFirefox" Then
				Call WshShell.Run("firefox.exe "& strPath, 3, false)
            ' If Browser to be launched is Netscape	
			Elseif strBrowser_Type="Netscape" Then
				Call WshShell.Run("netscape.exe "& strPath, 3, false)
			'If Browser to be launched is google chrome
			Elseif strBrowser_Type="GoogleChrome" or strBrowser_Type="chrome" or  strBrowser_Type="Chrome" or strBrowser_Type="CHROME" Then
				Call WshShell.Run("chrome.exe "& strPath, 3, false)				
            End If
				Set WshShell =Nothing  
			Wait(5)
   End If
End Function




