p_view_name =Parameter("p_view_name")


'*********************************Navigating to View*************************************************************
Select Case parameter("p_view_name")
	Case "Quota_by_Position_FA"
			call Pane_Click("ele_Other_Contents", "ele_EXPORT_I-08-005_QuotaByPosition")
	Case "Compensation_Plan_FA"
			Call Pane_Click("ele_Compensation_Plan_FA", "ele_EXPORT_I-08-004_Comp_Plan")
	Case "EXPORT_Account_Attributes"
			Call Pane_Click("ele_Other_Contents", "ele_Export_AA_view")
	Case "EXPORT_Assignment_Portfolio"
			Call Pane_Click("ele_Other_Contents", "ele_Export_APortfolio_view")		
	Case "Employee_Data_from_Workday"
			call Pane_Click("ele_IMPORT", "ele_Emp_data_Import_View")
	Case "Sales_Account_Data_from_SFDC"
			call Pane_Click("ele_IMPORT", "ele_Sales_acc_Import_View")
	Case "Sales_Structure"		
			Call Pane_Click("ele_Other_Contents", "ele_Sales_Structure_export")
	Case "Sales_Role"		
			Call Pane_Click("ele_Sales_Role_FA", "ele_Export_Sales_Role")		
	Case "Cluster"		
			Call Pane_Click("ele_Other_Contents", "ele_Export_Cluster")	
	Case "Accelerator"		
			Call Pane_Click("ele_Accelerators_Group_FA", "ele_Export_Accelerators")
			Fieldname=parameter("p_First_Field_Name")
			'getting the complete header names using the below function
			Target_Field_Name= fn_Anaplan_Field_names(Fieldname)	
			call fn_FilterSetUp("ele_Accelerator_is_deactivated")			
	Case "SalesManHierStructure"
			Call Pane_Click("ele_Other_Contents", "ele_EXPORT_Sales_Management")
	Case "PSA_ZipCodeMapping"
			Call Pane_Click("ele_Other_Contents", "ele_EXPORT_PSA_ZipCodeMapping")
	Case "TerritoryDefinitions"
			Call Pane_Click("ele_Other_Contents", "ele_EXPORT_Territory_Definitions")	
			'function to setup filter on the selected PSA column
			call fn_FilterSetUp("ele_PSA_Column")
			
	Case "PositionToAccountTerr"
			Call Pane_Click("ele_Other_Contents", "ele_EXPORT_Position_Assignments")	
End Select
