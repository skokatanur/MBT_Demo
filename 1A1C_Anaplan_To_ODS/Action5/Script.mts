'Click on logout
	
If Test_Object("btn_Anaplan_Logout_Option").exist(10) = True then
	Click_Object("btn_Anaplan_Logout_Option")
	wait 3
	call Click_Object("btn_Anaplan_Logout")
	LogResult_And_CaptureImage "br_Anaplan_Login", "User Logged out", "PASS", "User logged out from Anaplan successful."
	Wait 2
Else
	LogResult_And_CaptureImage "br_Anaplan_Login", "Logout Link", "Failed", "Logout Link not available."
End If


