Dim iPropertyId

dim CurentSalesLeadStatus 'aktualny status wątku do którego podpięty jest dokument.
dim PreviousDocumentStatus 'stastus dokumentu z którego nastąpiła zmiana
dim CurentDocumentStatus 'status dokumentu na który nastąpiła zmiana
dim objCurentDocumentStatus


'Odszukanie wartości pola SALES LEAD
Set myPropertyValues = Vault.ObjectPropertyOperations.GetProperties(ObjVer)

'Set ID_GrupaKategori = myPropertyValues.SearchForProperty(101)

Dim ID_GrupaKategori, ID_Kategori
Set ID_GrupaKategori = Vault.ObjectPropertyOperations.GetProperty( ObjVer, 101).Value.GetValueAsLookup

Set ID_Kategori = Vault.ObjectPropertyOperations.GetProperty( ObjVer, 100).Value.GetValueAsLookup
'if (ID_Kategoria.DisplayID = "20" or ID_Kategoria.DisplayID = "21" ) then

	dim metadana
	dodajDoLog("Przed Cint")
	przelacznik = Cint(ID_GrupaKategori.DisplayID)
	dodajDoLog("Przelacznik = " & przelacznik)
	Select Case przelacznik
		Case 1 'Sales&Marketing

			metadana = 1455
			'ID właściwości po ktorej bedziemy poszukiwać obiektu podrzędnego
			iPropertyId = 1484


		Case 4 'Production

			if(ID_Kategori.DisplayID = 156) then
				metadana = 11604
				'ID właściwości po ktorej bedziemy poszukiwać obiektu podrzędnego
				iPropertyId = 1048
				dodajDoLog("Typ obiektu 156 = " & iPropertyId)
			elseif (ID_Kategori.DisplayID = 159) then
				metadana = 11604
				'ID właściwości po ktorej bedziemy poszukiwać obiektu podrzędnego
				iPropertyId = 1050
				dodajDoLog("Typ obiektu 159 = " & iPropertyId)
			end if


	End Select

	dodajDoLog("Po selectci")

Set myPropertyValue = myPropertyValues.SearchForProperty(iPropertyId)



'ID właściwości po ktorej bedziemy poszukiwać obiektu podrzędnego


Dim myPropertyValues
Dim myPropertyValue


dodajDoLog("ID POLA = " & iPropertyId)


    'Get the value as lookup reference. Odczytanie warości pola SALES LEAD
    Dim oLookUpRef	
	Set oLookUpRef = Vault.ObjectPropertyOperations.GetProperty( ObjVer, iPropertyId).TypedValue.GetValueAsLookup	

    ' Resolve the target object type for the property value.
    Dim oPropertyDef 
	Set oPropertyDef = Vault.PropertyDefOperations.GetPropertyDef(myPropertyValue.PropertyDef)

    Dim oValListObjType
	Set oValListObjType = Vault.ValueListOperations.GetValueList(oPropertyDef.ValueList)

	'Create ObjID
	dim oObjID: Set oObjID = CreateObject("MFilesAPI.ObjID")
	oObjID.ID   =  oLookUpRef.item 'ID for related Member object
	oObjID.Type =  oLookUpRef.ObjectType 'related member ObjType
	
	'Get Latest version of external object. Pobranie aktualnego biektu SALES LEAD celem odczytania z niego metadanych
	dim oMember_ObjVer
	set oMember_ObjVer = vault.ObjectOperations.GetLatestObjVer(oObjID,true)

'--------------------------------------------------------------------------------------------------------
' ---  --  - Tu sprawdzane jest czy osoba która utworzyła bierzący dokument jest właścicielem SalesLeada
'--------------------------------------------------------------------------------------------------------
dim ID_SalesLead_ResponsiblePerson
dim ID_CurentDocument_Creator
dodajDoLog("Sprawdzanie własciciela sales leada")

' oMember_ObjVer <---- to jest obiekt Sales Lead powiązany z bierzącym dokumentem.

dim ObiektWypluty_ResponsiblePerson  ' <---- to jest obiekt Responsible Person z SalesLeada
Set ObiektWypluty_ResponsiblePerson = znajdzSupervisora (metadana , oMember_ObjVer)

dim ObiektWypluty_Manager  ' <---- to jest obiekt Responsible Person z SalesLeada
Set ObiektWypluty_Manager = znajdzSupervisora (1162, oMember_ObjVer)



ID_SalesLead_ResponsiblePerson = Cint(Vault.ObjectPropertyOperations.GetProperty( ObiektWypluty_ResponsiblePerson.ObjVer, 1213).TypedValue.GetLookupID)
ID_CurentDocument_Creator = Cint(myPropertyValues.SearchForProperty(MFBuiltInPropertyDefCreatedBy).TypedValue.GetLookupID)

dodajDoLog("Sales Lead responsible person to - " & ID_SalesLead_ResponsiblePerson)
dodajDoLog("Obecny dokument został utworzony przez - " & ID_CurentDocument_Creator)


Set objCurentDocumentStatus = CreateObject("MFilesAPI.PropertyValue")
objCurentDocumentStatus.PropertyDef = 11585





	'Create the current user ID as typed value
	Dim typed: Set typed = CreateObject("MFilesAPI.TypedValue")
	Dim oLUUsers: Set oLUUsers = CreateObject("MFilesAPI.Lookups")
	Dim oLUOneUser: Set oLUOneUser = CreateObject("MFilesAPI.Lookup")
	

if ID_SalesLead_ResponsiblePerson = ID_CurentDocument_Creator then
	objCurentDocumentStatus.TypedValue.SetValue MFilesAPI.MFDataType.MFDatatypeLookup, 1
	

else 
	objCurentDocumentStatus.TypedValue.SetValue MFilesAPI.MFDataType.MFDatatypeLookup, 1
	











end if
Vault.ObjectPropertyOperations.SetProperty ObjVer, objCurentDocumentStatus





'--------------------------------------------------------------------------------------------------------
' 
'--------------------------------------------------------------------------------------------------------

	CurentSalesLeadStatus = Cint(Vault.ObjectPropertyOperations.GetProperty( oMember_ObjVer, 11585).TypedValue.GetValueAsLocalizedText)
	CurentDocumentStatus = Cint(Vault.ObjectPropertyOperations.GetProperty( ObjVer, 11585).TypedValue.GetValueAsLocalizedText)

	Dim oPropValsNew
	Dim myStatusValue
	dim moja

	If ObjVer.Version > 1 then

		Dim staryObjVer
		Set staryObjVer = ObjVer.Clone()
		dim ver
		ver= staryObjVer.Version
		ver=ver-1 
		staryObjVer.Version = ver 
		Set myPropertyValues1 = Vault.ObjectPropertyOperations.GetProperties(staryObjVer)

		'id naszego salesleada oObjID.ID
		
		Set moja = myPropertyValues1.SearchForProperty(11585)
		dodajDoLog("Przed PD status")
	End if	

	if ( IsNull(moja) or IsEmpty(moja)) then 
		PreviousDocumentStatus = 11    'Jeśli wersja dokumentu = 1
	elseif ( IsNull(moja.TypedValue) or IsEmpty(moja.TypedValue)) then 
		PreviousDocumentStatus = 11    'Jeśli wersja dokumentu > 1 ale pole StatusDokumentu nie ma wartości
	else
		PreviousDocumentStatus = Cint(moja.TypedValue.GetValueAsLocalizedText)
	End if

	if CurentSalesLeadStatus > CurentDocumentStatus then
		
		Set oPropValsNew = CreateObject("MFilesAPI.PropertyValues")
		Set myStatusValue= CreateObject("MFilesAPI.PropertyValue")
		myStatusValue.PropertyDef = 11585
		myStatusValue.TypedValue.SetValue MFDatatypeLookup , CurentDocumentStatus
	
		oPropValsNew.Add -1, myStatusValue
		Vault.ObjectPropertyOperations.SetProperties oMember_ObjVer, oPropValsNew
		
	Elseif CurentSalesLeadStatus = PreviousDocumentStatus then


		dodajDoLog("Przed funkcją - ZNAJDZ DOKUMENTY")
		ib = Cint(znajdzDokumenty( PreviousDocumentStatus, oObjID.ID))
		dodajDoLog("Po wyszukiwaniu statusów - " & ib)
		if ib = 0 then 
			ib =CurentDocumentStatus
		end if
		Set oPropValsNew = CreateObject("MFilesAPI.PropertyValues")
		Set myStatusValue= CreateObject("MFilesAPI.PropertyValue")
		myStatusValue.PropertyDef = 11585
		myStatusValue.TypedValue.SetValue MFDatatypeLookup , ib
	
		oPropValsNew.Add -1, myStatusValue
		Vault.ObjectPropertyOperations.SetProperties oMember_ObjVer, oPropValsNew

	End if



				' Set last modified by user.
		Set oLastModifiedBy = CreateObject("MFilesAPI.TypedValue")
		oLastModifiedBy.SetValue MFDatatypeLookup, CurrentUserID	
		Vault.ObjectPropertyOperations.SetLastModificationInfoAdmin ObjVer, True, oLastModifiedBy, false, Nothing



 '   ZAPIS DANYCH DO LOGA   ////////////////////
sub dodajDoLog(message)
		ForAppending = 8
		 
		set objFSO = CreateObject("Scripting.FileSystemObject")
		set objFile = objFSO.OpenTextFile("E:\test\USTAW_ACTION_LOOP.txt", ForAppending, True)
	 
		objFile.WriteLine(message)
		objFile.Close
end sub


Function znajdzDokumenty(nr_statusu, id_SalesLead)



Dim oSearchConditions1: Set oSearchConditions1 = CreateObject("MFilesAPI.SearchConditions") 
' Create a search condition for the object class (i.e., document in this case).
Dim oSearchCondition1: Set oSearchCondition1 = CreateObject("MFilesAPI.SearchCondition") 
oSearchCondition1.ConditionType = MFilesAPI.MFConditionType.MFConditionTypeEqual
oSearchCondition1.Expression.DataStatusValueType = MFStatusTypeObjectTypeID
oSearchCondition1.TypedValue.SetValue MFilesAPI.MFDataType.MFDatatypeLookup, MFilesAPI.MFBuiltInObjectType.MFBuiltInObjectTypeDocument
oSearchConditions1.Add -1, oSearchCondition1

' Create a search condition for property SALES LEAD.
Dim oSearchCondition2: Set oSearchCondition2 = CreateObject("MFilesAPI.SearchCondition") 
oSearchCondition2.ConditionType = MFilesAPI.MFConditionType.MFConditionTypeEqual
oSearchCondition2.Expression.SetPropertyValueExpression 1484,  MFilesAPI.MFParentChildBehavior.MFParentChildBehaviorNone, Nothing
oSearchCondition2.TypedValue.SetValue MFilesAPI.MFDataType.MFDatatypeLookup, id_SalesLead
oSearchConditions1.Add -1, oSearchCondition2


' Create a search condition for property STATUS = Approve (OUTGOING)
Dim oSearchCondition3: Set oSearchCondition3 = CreateObject("MFilesAPI.SearchCondition") 
oSearchCondition3.ConditionType = MFilesAPI.MFConditionType.MFConditionTypeEqual
oSearchCondition3.Expression.SetPropertyValueExpression 11585,  MFilesAPI.MFParentChildBehavior.MFParentChildBehaviorNone, Nothing

' Create a search condition for property DELETED. /nie bierzemy pod uwagę dokumentów usuniętych
Dim oSearchCondition4: Set oSearchCondition4 = CreateObject("MFilesAPI.SearchCondition") 
oSearchCondition4.ConditionType = MFilesAPI.MFConditionType.MFConditionTypeEqual
oSearchCondition4.Expression.DataStatusValueType = MFStatusTypeDeleted
oSearchCondition4.TypedValue.SetValue MFilesAPI.MFDataType.MFDatatypeBoolean, False
oSearchConditions1.Add -1, oSearchCondition4

for i=nr_statusu To 10 Step 1

dodajDoLog("w forze - " & nr_statusu & " -- " & i )
	oSearchCondition3.TypedValue.SetValue MFilesAPI.MFDataType.MFDatatypeLookup, i
	oSearchConditions1.Add -1, oSearchCondition3

	Dim oObjectVersions1

	Set oObjectVersions1 = Vault.ObjectSearchOperations.SearchForObjectsByConditions ( oSearchConditions1, MFilesAPI.MFSearchFlags.MFSearchFlagNone, true)
	dodajDoLog("Znalazłem tyle dokumentów - " & oObjectVersions1.Count)
	if oObjectVersions1.Count > 0 then
		znajdzDokumenty = i
		Exit Function
	end if
	
	oSearchConditions1.Remove(3)



next

end Function


'-----------------------------------------------------------------------------------------------------------------------------------
'       Funkcja wyszukująca ID użytkownika na podstawie pola M-Files user
'
'
'-----------------------------------------------------------------------------------------------------------------------------------


Function znajdzSupervisora(ID_WLASCIWOSCI, oOBJ_VER)

Dim iPropertyId 

'ID właściwości po ktorej bedziemy poszukiwać obiektu podrzędnego
iPropertyId = ID_WLASCIWOSCI

Dim myPropertyValues
Dim myPropertyValue
dodajDoLog("Wszedłem do funkcji")
'Odszukanie wartości pola ID_WLASCIWOSCI
'Set myPropertyValues = Vault.ObjectPropertyOperations.GetProperties(oOBJ_VER)
'Set myPropertyValue = myPropertyValues.SearchForProperty(iPropertyId)

    'Get the value as lookup reference.
    Dim oLookUpRef	
	Set oLookUpRef = Vault.ObjectPropertyOperations.GetProperty( oOBJ_VER, iPropertyId).TypedValue.GetValueAsLookup	


	'Create ObjID
	dim oObjID: Set oObjID = CreateObject("MFilesAPI.ObjID")
	oObjID.ID   =  oLookUpRef.item 'ID for related Member object
	oObjID.Type =  oLookUpRef.ObjectType 'related member ObjType
	
	'Get Latest version of external object
'	dim oMember_ObjVer
'	set oMember_ObjVer = vault.ObjectOperations.GetLatestObjVer(oObjID,true)


	' Initialize an array of search conditions.
Dim oSearchConditions1: Set oSearchConditions1 = CreateObject("MFilesAPI.SearchConditions")  

' Create a search condition for the object class (i.e., document in this case).
Dim oSearchCondition1: Set oSearchCondition1 = CreateObject("MFilesAPI.SearchCondition") 
oSearchCondition1.ConditionType = MFilesAPI.MFConditionType.MFConditionTypeEqual
oSearchCondition1.Expression.DataStatusValueType = MFStatusTypeObjectID
oSearchCondition1.TypedValue.SetValue MFilesAPI.MFDataType.MFDatatypeInteger, oLookUpRef.item
oSearchConditions1.Add -1, oSearchCondition1

Dim oSearchCondition2: Set oSearchCondition2 = CreateObject("MFilesAPI.SearchCondition") 
oSearchCondition2.ConditionType = MFilesAPI.MFConditionType.MFConditionTypeEqual
oSearchCondition2.Expression.DataStatusValueType = MFStatusTypeObjectTypeID
oSearchCondition2.TypedValue.SetValue MFilesAPI.MFDataType.MFDatatypeLookup, 103
oSearchConditions1.Add -1, oSearchCondition2


Dim oObjectVersions1
Set oObjectVersions1 = Vault.ObjectSearchOperations.SearchForObjectsByConditions ( oSearchConditions1, MFilesAPI.MFSearchFlags.MFSearchFlagNone, False)
dodajDoLog("Znalazłem Responsible Person dla nowego salesleada - "& oObjectVersions1.Count)
dodajDoLog("Funkcja =znajdzSupervisora= zakonczyła działanie.")
Set znajdzSupervisora = oObjectVersions1.Item(1)
'-------------wypluj obiekt znaleziony --------------


end Function



'-----------------------------------------------------------------------------------------------------------------------------------------------------
'
'							Funkcja tworząca nowy ASSIGNMENT
'
'-----------------------------------------------------------------------------------------------------------------------------------------------------

Sub pAssignment(szDescriptionL, szAssignedToL, szMonitoredByL, szLinkDocID)
	Dim oPropertyValues
	Set oPropertyValues = CreateObject("MFilesAPI.PropertyValues")
	Dim oPropertyValue 
	Set oPropertyValue = CreateObject("MFilesAPI.PropertyValue")
	' NameOrTitle
	oPropertyValue.PropertyDef = MFBuiltInPropertyDefNameOrTitle
	oPropertyValue.TypedValue.SetValue MFDataTypeText, szDescriptionL
	oPropertyValues.Add -1, oPropertyValue	
	' Description
	oPropertyValue.PropertyDef = MFBuiltInPropertyDefAssignmentDescription
	oPropertyValue.TypedValue.SetValue MFDatatypeMultiLineText, "Please check linked document"
	oPropertyValues.Add -1, oPropertyValue		
	' Object class  MFBuiltInObjectClassGenericAssignment (-100=assignment)
	oPropertyValue.PropertyDef = MFBuiltInPropertyDefClass
	oPropertyValue.TypedValue.SetValue MFDatatypeLookup, MFBuiltInObjectClassGenericAssignment
	oPropertyValues.Add -1, oPropertyValue
	' AssignedToUsers
	oPropertyValue.PropertyDef = MFBuiltInPropertyDefAssignedTo
	oPropertyValue.Value.SetValueToMultiSelectLookup szAssignedToL
	oPropertyValues.Add -1, oPropertyValue
	'MonitoredBy 
	oPropertyValue.PropertyDef = MFBuiltInPropertyDefMonitoredBy
	oPropertyValue.Value.SetValueToMultiSelectLookup szMonitoredByL
	oPropertyValues.Add -1, oPropertyValue
	'SFD/MFD
	oPropertyValue.PropertyDef = MFBuiltInPropertyDefSingleFileObject
	oPropertyValue.TypedValue.SetValue MFDatatypeBoolean, False
	oPropertyValues.Add -1, oPropertyValue
	'MFBuiltInPropertyDefDeadline
	'oPropertyValue.PropertyDef = MFBuiltInPropertyDefDeadline
	'oPropertyValue.TypedValue.SetValue MFDatatypeDate, DateAdd("d",2,Now())
	'oPropertyValues.Add -1, oPropertyValue
	'Linked document 
	oPropertyValue.PropertyDef = 1263	'1241 DocumentList PropertyID
	oPropertyValue.TypedValue.SetValue MFDatatypeLookup, szLinkDocID	
	oPropertyValues.Add -1, oPropertyValue
	
	dim props
	set props = vault.ObjectOperations.CreateNewObject (MFBuiltInObjectTypeAssignment, oPropertyValues, Nothing, Nothing)
	Vault.ObjectOperations.CheckIn( props.ObjVer )		
	
End Sub

