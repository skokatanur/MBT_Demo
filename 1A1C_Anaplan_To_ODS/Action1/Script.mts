strURLPATH=parameter("P_URL_In")
strUsername=parameter("P_UserName_In")
strPassword=parameter("P_Password_In")

NetsuiteConfigPath="C:\FAST_Test_Automation\Test_Data\1A1C_Test_Data.xml"
	set objFSOconfig = createobject("Scripting.FileSystemObject")
	if objFSOconfig.FileExists(NetsuiteConfigPath) then
		Environment.LoadFromFile(NetsuiteConfigPath)
		strUserName=Environment.Value("Anaplan_UserName")
		strPassword=Environment.Value("Anaplan_Password")
		strURLPATH=Environment.Value("Anaplan_URL")
	End If


SystemUtil.CloseProcessByName("chrome.exe")  
SystemUtil.Run "chrome.exe",strURLPATH
wait 10
'************new code
'Browser("Anaplan - Frontdoor").Page("Anaplan - Frontdoor")
if Browser("Anaplan - Frontdoor").Page("Anaplan - Frontdoor").link("xpath:=//a[text()='Log in with Single Sign-on (SSO)']").exist(10) then

	Browser("Anaplan - Frontdoor").Page("Anaplan - Frontdoor").link("xpath:=//a[text()='Log in with Single Sign-on (SSO)']").Click
	Browser("Anaplan - Frontdoor").Page("Anaplan - Frontdoor").webedit("xpath:=//form[@id='ssoForm']//*[@type='email']").set "sharada.kokatanur@microfocus.com"
	wait 2
	
	Browser("Anaplan - Frontdoor").Page("Anaplan - Frontdoor").WebButton("xpath:=//button[@class='continue_button']").Click
	
	if Browser("Anaplan - Frontdoor").Page("Anaplan - Frontdoor").link("xpath:=//a[@id='mcro']").exist(10) then
		Browser("Anaplan - Frontdoor").Page("Anaplan - Frontdoor").link("xpath:=//a[@id='mcro']").Click
	End If
else
'*************

	If Test_object("btn_sso").exist(30) Then
		Test_object("btn_sso").click
	End If
	
	If Test_object("txt_EmailAddress").exist(30) then
		If Enter_Value_In_Edit_Field("txt_EmailAddress","sharada.kokatanur@microfocus.com","No") <> "PASS" then 
						'Exit For
		End If
		Test_object("btn_Next_sso").click
	End If
End If
wait 10
set mySendKeys = CreateObject("WScript.shell")
mySendKeys.SendKeys("{ESC}")

	If Test_Object("pg_Micro_Focus_Stack_A_Login").Exist(60) Then
'				If Enter_Value_In_Edit_Field("txt_Ecom_User_ID",Parameter("P_Username_In"),"No") <> "PASS" then 
				If Enter_Value_In_Edit_Field("txt_Ecom_User_ID",strUserName,"No") <> "PASS" then 
					'Exit For
				End If
		  
'				If Enter_Value_In_Edit_Field("txt_Ecom_Password",Parameter("P_Password_In"),"Yes") <> "PASS" then 
				If Enter_Value_In_Edit_Field("txt_Ecom_Password",strPassword,"Yes") <> "PASS" then 
					'Exit For
				End If
			
		       'Click the Login_Button button.
		    If Click_Object("btn_Login_sso") <> "PASS" then 
				'Exit For
			End If
		End If
