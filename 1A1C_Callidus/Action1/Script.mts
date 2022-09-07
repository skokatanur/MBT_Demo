
strURLPATH = Parameter("P_URL_In")
strUsername = Parameter("P_UserName_In")
strPassword = Parameter("P_Password_In")

SystemUtil.CloseProcessByName("chrome.exe")  
SystemUtil.Run "chrome.exe",strURLPATH
wait 60
'username and password and click on Sign in
	If Test_Object("txt_callidus_username").exist(60) then 
		Call Enter_Value_In_Edit_Field("txt_callidus_username", strUsername, "No")
		Call Enter_Value_In_Edit_Field("txt_callidus_password", strPassword, "No")
		call Click_Object("btn_Login")
	End if	
call Tableau_Buffer()

If Test_Object("pg_Sales_Performance_Home").exist(60) then
	LogResult_And_CaptureImage "br_Sales_Performance_Home", "Login Authorization", "PASS", "Login to Callidus is successful."
else
	LogResult_And_CaptureImage "br_Sales_Performance_Home", "Login Authorization", "FAIL", "Login to Callidus is failed."
	ExitTest
End if	


