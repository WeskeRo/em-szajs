'Dział IT _________________ 1
'Dział S i M - Produkty ___ 2
'Dział Produkcji __________ 3
'Dział R&D ________________ 4
'Dział Serwisu ____________ 5
'Administracja ____________ 6
'Dział Utrzymania Jakości _ 7
'Dział S i M - Usługi _____ 8
'Zarząd ___________________ 9

'===================================================================================================
DIM value

'ObjVer
'	kategoria_leada = Vault.ObjectPropertyOperations.GetProperty( connected_lead, 100).TypedValue.GetLookupID
'value = PropertyValue.TypedValue.GetLookupID





Set Obiekt_ResponsiblePerson = LookupObject(11604,ObjVer)
set oLookups_ResponsiblePerson_Departments = Vault.ObjectPropertyOperations.GetProperty( Obiekt_ResponsiblePerson, 1212).TypedValue.GetValueAsLookups



Dim Departments_Array()
ReDim Preserve Departments_Array(oLookups_ResponsiblePerson_Departments.Count)

For i = 1 To oLookups_ResponsiblePerson_Departments.Count

	Set oLookup = oLookups_ResponsiblePerson_Departments.Item(i)
	Departments_Array(i) = oLookup.DisplayID			
			
Next

Select Case ObjVer.Type
	
	Case 1 'SalesLead
	
	Case 2 'Offering Process
	
	Case 3 'Commercial Process
	
	Case 4 'Production Process
	
	Case 5 'Product
		'------Dopuszczalne działy------
		' 3,
		
		if IsInArray(CStr(3), Departments_Array) then
		Err.Raise MFScriptCancel, "Jestem na liście"
		else 
		Err.Raise MFScriptCancel, "Nie jestem na liście"
		End If
	
	Case 6 'Test Order
	
	Case 7 'Support Process
	
	Case 8 'Service Process
	
	Case 9 'SoftwareProcess
	
	Case 10 'Development Process
	
	Case 11 'Development Project
	
	Case 12 'Development Task
	
End Select	


	Err.Raise MFScriptCancel, "pole " & olookup.DisplayID


If Len(value) <> 100 Then

Err.Raise MFScriptCancel, "The value you enter must be 10 characters long. " & kategoria_leada

End If


'====================================================================================================================
Function LookupObject(ID_Metadana, oOBJ_VER)
    
	Dim oLookUpRef	
	Set oLookUpRef = Vault.ObjectPropertyOperations.GetProperty( oOBJ_VER, ID_Metadana).TypedValue.GetValueAsLookup	

	'Create ObjID
	dim oObjID: Set oObjID = CreateObject("MFilesAPI.ObjID")
	oObjID.ID   =  oLookUpRef.item 'ID for related Member object
	oObjID.Type =  oLookUpRef.ObjectType 'related member ObjType

	'dim oProces_ObjVer
	'set oProces_ObjVer = Vault.ObjectOperations.GetLatestObjVer(oObjID,true)  '
	
	Set LookupObject = Vault.ObjectOperations.GetLatestObjVer(oObjID,true)
	
End Function


'====================================================================================================================
Function CheckDepartment(oDepartments, Tablica_ , oOBJ_VER)
    
	Dim oLookUpRef	
	Set oLookUpRef = Vault.ObjectPropertyOperations.GetProperty( oOBJ_VER, ID_Metadana).TypedValue.GetValueAsLookup	

	'Create ObjID
	dim oObjID: Set oObjID = CreateObject("MFilesAPI.ObjID")
	oObjID.ID   =  oLookUpRef.item 'ID for related Member object
	oObjID.Type =  oLookUpRef.ObjectType 'related member ObjType

	'dim oProces_ObjVer
	'set oProces_ObjVer = Vault.ObjectOperations.GetLatestObjVer(oObjID,true)  '
	
	Set LookupObject = Vault.ObjectOperations.GetLatestObjVer(oObjID,true)
	
End Function


'====================================================================================================================
Function IsInArray(strIn, arrCheck)
    'IsInArray: Checks for a value inside an array
    'Author: Justin Doles - www.DigitalDeviation.com
    Dim bFlag 
	bFlag = false
 
    If IsArray(arrCheck) AND Not IsNull(strIn) Then
        Dim i
        For i = 0 to UBound(arrCheck)
            If arrcheck(i) = strIn Then
                bFlag = true
                Exit For
            End If
        Next
    End If
    IsInArray = bFlag
End Function