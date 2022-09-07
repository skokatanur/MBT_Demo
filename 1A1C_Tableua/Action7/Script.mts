'Click on login and enter username and pwd
If Test_Object("btn_My_Content_and_Settings").exist(10) = True then
	Click_Object("btn_My_Content_and_Settings")
	wait 2
	Click_Object("ele_Sign_Out1")
Else
	LogResult_And_CaptureImage "br_Tableau_Server", "Logout Link", "Failed", "Logout Link not available."
End If

wait 3
'verify if the user is navigated to model page
'if Object_Exists("txt_crl_username") = "PASS" then
'	LogResult_And_CaptureImage "br_Tableau_Server", "User Logged out", "PASS", "User logged out from Tableau successful."
'else
'	LogResult_And_CaptureImage "br_Tableau_Server", "Logout Failed", "FAIL", "Logout from TAbleau failed."
'End if	

call Close_All_Browser_Instances()
