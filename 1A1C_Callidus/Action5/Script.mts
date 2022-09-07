If Test_Object("img_Profile").exist(10) = True Then
	Click_Object("img_Profile")
	If Test_Object("lnk_Sign_Out").exist(5)=True Then
		Click_Object("lnk_Sign_Out")
		wait 5
		LogResult_And_CaptureImage "pg_Sales_Performance_Home", "LOGOUT SUCCESSFUL", "PASS", "LOGOUT SUCCESSFUL"
		else
		LogResult_And_CaptureImage "lnk_Sign_Out", "Log out Failed", "FAIL", "LOG OUT Step Failed"
	End If
End If
