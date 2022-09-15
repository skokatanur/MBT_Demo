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

wait(60)
Browser("br_User_Defined_Reports").Page("pg_User_Defined_Reports").WebList("tab-shared-widget-166323452040").Select "CommissionsPreProd"



