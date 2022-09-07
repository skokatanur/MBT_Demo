'Click on the Model name

'Click_Object("ele_MF_FY19HUB_ CSIT_MFI")

strModelName=parameter("P_Model_Name")
'Pre prd fy21
Test_Object("lnk_Models").Highlight
Test_Object("lnk_Models").click

NetsuiteConfigPath="C:\FAST_Test_Automation\Test_Data\1A1C_Test_Data.xml"
	set objFSOconfig = createobject("Scripting.FileSystemObject")
	if objFSOconfig.FileExists(NetsuiteConfigPath) then
		Environment.LoadFromFile(NetsuiteConfigPath)
		strModelName=Environment.Value("Anaplan_Model")
	End If



If ucase("PREPROD FY22 Hub")=ucase(strModelName) Then
	Browser("Anaplan - Frontdoor").Page("Anaplan - Home").Link("text:=PREPROD FY22 Hub").Highlight
	Browser("Anaplan - Frontdoor").Page("Anaplan - Home").Link("text:=PREPROD FY22 Hub").Click
	wait 20

ELSE
	Browser("Anaplan - Frontdoor").Page("Anaplan - Home").Link("text:=CSIT FY22 Hub").Highlight

	Browser("Anaplan - Frontdoor").Page("Anaplan - Home").Link("text:=CSIT FY22 Hub").Click
	wait 20
End If

