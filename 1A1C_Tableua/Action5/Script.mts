strpath="C:\FAST_Test_Automation\"

reportName=parameter("p_Report_Name_In")

str_Employee_Name=parameter("p_Employee_Name_In")
str_Fiscal_Year = parameter("p_Fiscal_Year_In")

Call Tableau_FiscalYear(str_Fiscal_Year)
wait 3
Call MIPR_Report(str_Employee_Name)


Function Tableau_FiscalYear(str_Fiscal_Year)

		If Test_Object("lst_Filter_Fiscal_Year_Name").exist(20) then
			Test_Object("lst_Filter_Fiscal_Year_Name").hovertap
			Test_Object("lst_Filter_Fiscal_Year_Name").highlight
			Test_Object("lst_Filter_Fiscal_Year_Name").click
		End If

			
	wait 5
	call Tableau_Buffer()	

	'call SetTo_Click_Object("ele_Draw_year","xpath","//DIV[@role='option' and normalize-space()="&str_Fiscal_Year&"]/DIV[2]")
	'call SetTo_Click_Object("ele_Draw_year","xpath","//span[@class='tabMenuItemName' and normalize-space() ="&str_Fiscal_Year&"]")
	'call SetTo_Click_Object("ele_Draw_year","xpath","//*[text() ="&str_Fiscal_Year&"]")
	call Tableau_Buffer()
	
End Function
