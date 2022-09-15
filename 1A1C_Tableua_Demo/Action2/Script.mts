''Call Fn_Component_Start()
'reportName = Parameter("P_Tableau_Report_Name_in")
'strSearchSite = Parameter("P_Search_Site_In")
'
''Select Site
'call Enter_Value_In_Edit_Field("txt_Search_Site",strSearchSite,"No")
'
'If Test_Object("ele_CommissionsPreProd").exist(5) Then
'	Click_Object("ele_CommissionsPreProd")
'End If
'

wait(30)

Browser("br_Tableau_Server").Page("pg_Tableau_Server").WebElement("ele_CommissionsPreProd_Old").Highlight
Browser("br_Tableau_Server").Page("pg_Tableau_Server").WebElement("ele_CommissionsPreProd_Old").Click
