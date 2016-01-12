if( ObjVer.Type = 184) then
	Dim ACL 
	Set ACL = CreateObject("MFilesAPI.AccessControlList")

	Dim ID_Kategoria
	 ID_Kategoria = Vault.ObjectPropertyOperations.GetProperty( ObjVer, 100).TypedValue.GetLookupID
		
	'gg
	if( ID_Kategoria = "261" or ID_Kategoria = "199") then
	
	dim ster_subdep
	
	ster_subdep = true
	
		if ( IsNull (Vault.ObjectPropertyOperations.GetProperty( ObjVer, 11802).Value) or IsEmpty(Vault.ObjectPropertyOperations.GetProperty( ObjVer, 11802).Value)) then 'pole recznego wpisywania numeru
		'	Err.Raise mfscriptcancel, "SUB DEPARTMENT IS EMPTY"
			ster_subdep = false
		End If
		Dim oLookUps_Dep

		if ster_subdep then
			'Get the value as lookup reference. Odczytanie warości pola SALES LEAD
			
			Set oLookUps_Dep = Vault.ObjectPropertyOperations.GetProperty( ObjVer, 1212).TypedValue.GetValueAsLookups	
			set Department = oLookUps_Dep.Item(1)

			Dim oLookUps_SubDep	
			Set oLookUps_SubDep = Vault.ObjectPropertyOperations.GetProperty( ObjVer, 11802).TypedValue.GetValueAsLookups	
			set SubDepartment = oLookUps_SubDep.Item(1)
			

			Set ACL = Ustal_Uprawnienia (Department.Item, SubDepartment.Item)

		else
		
			'Get the value as lookup reference. Odczytanie warości pola SALES LEAD
				
			Set oLookUps_Dep = Vault.ObjectPropertyOperations.GetProperty( ObjVer, 1212).TypedValue.GetValueAsLookups	
			set Department = oLookUps_Dep.Item(1)
			
			Set ACL = Ustal_Uprawnienia (Department.Item, 0)
		
		End If
	End if

call Vault.ObjectOperations.ChangePermissionsToACL(ObjVer, ACL, True)

End if

Function Ustal_Uprawnienia (Dep_ID, SubDep_ID)


'-+DEPARTMENTS+-================================
'Dział IT						1
'Dział Sprzedaży i Marketingu	2
'Dział Produkcji				3
'Dział R&D						4
'Administracja					6
'Dział Utrzymania Jakości		7
'Board							9
'===============================================


'-+SUB-DEPARTMENTS+-=============================
'Products					1
'SMT Services				2
'Quality Control			3
'Production Processing		4
'Logistic					5
'Service					6
'Accountancy				8
'HR							9
'Health & Safety			10
'Goods						11
'SMT Services				12
'Board Office				13
'===============================================

Dim oACL 
Dim oACEK 
Dim oACED 

Set oACED = CreateObject("MFilesAPI.AccessControlEntryData")
Set oACEK = CreateObject("MFilesAPI.AccessControlEntryKey")
Set oACL = CreateObject("MFilesAPI.AccessControlList")

oACED.ReadPermission = MFilesAPI.MFPermission.MFPermissionAllow
oACED.EditPermission = MFilesAPI.MFPermission.MFPermissionAllow
oACED.DeletePermission = MFilesAPI.MFPermission.MFPermissionAllow

Select Case Dep_ID
	Case 1 'Dział IT
		Select Case SubDep_ID
			Case Else
			call oACEK.SetUserOrGroupID(135, true)

		End Select

	Case 2 'Dział Sprzedaży i Marketingu
		Select Case SubDep_ID
			Case 11 'Goods
			call oACEK.SetUserOrGroupID(141, true)

			Case 12 'SMT Services
			call oACEK.SetUserOrGroupID(142, true)

			Case Else
			call oACEK.SetUserOrGroupID(132, true)

		End Select

	Case 3 'Dział Produkcji
		Select Case SubDep_ID
			Case 1 'Products
			call oACEK.SetUserOrGroupID(143, true)

			Case 2 'SMT Services
			call oACEK.SetUserOrGroupID(144, true)

			Case 3 'Quality Control
			call oACEK.SetUserOrGroupID(145, true)

			Case 4 'Production Processing
			call oACEK.SetUserOrGroupID(146, true)

			Case 5 'Logistic
			call oACEK.SetUserOrGroupID(147, true)

			Case 6 'Service
			call oACEK.SetUserOrGroupID(148, true)

			Case Else
			call oACEK.SetUserOrGroupID(133, true)

		End Select

	Case 4 'Dział R&D
		Select Case SubDep_ID
			Case Else
			call oACEK.SetUserOrGroupID(136, true)

		End Select

	Case 6 'Administracja
		Select Case SubDep_ID
			Case 8 'Acountancy
			call oACEK.SetUserOrGroupID(138, true)

			Case 9 'HR
			call oACEK.SetUserOrGroupID(139, true)

			Case 10 'Health & Safety
			call oACEK.SetUserOrGroupID(140, true)

			Case Else
			call oACEK.SetUserOrGroupID(131, true)

		End Select

	Case 7 'Dział Utrzymania Jakości
		Select Case SubDep_ID
			Case Else
			call oACEK.SetUserOrGroupID(137, true)

		End Select

	Case 9 'Board	
		Select Case SubDep_ID
			Case 13 'Board Office
			call oACEK.SetUserOrGroupID(149, true)

			Case Else
			call oACEK.SetUserOrGroupID(134, true)

		End Select

End Select


call oACL.CustomComponent.AccessControlEntries.Add(oACEK, oACED)
'Err.Raise mfscriptcancel, "STOP"
Set Ustal_Uprawnienia = oACL

End Function
