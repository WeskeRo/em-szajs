	dodajDoLog("---------------------------")
	dodajDoLog("SalesLead_NAME - enter")



'//////////SPRAWDZANIE CZY OBIEKT TO WĄTEK SPRZEDAZOWY   /////////////
if( ObjVer.Type = 182) then

	Dim nazwaZm 
	dim objVersionAndProps
	Dim wiadomosc
	Dim sCreatedBy  'Get Created by 
	Dim Rok
	Dim Nazwa_watku
	Dim Licznik
	Dim oSCs
	Dim oSC
	Dim myPropertyValue
	dim postfix

	Set objVersionAndProps = Vault.ObjectPropertyOperations.GetProperties(ObjVer)
	sCreatedBy = objVersionAndProps.SearchForProperty(MFBuiltInPropertyDefCreatedBy).TypedValue.GetLookupID
	Rok = Year(Now)
	'Numer_watku  = Vault.ObjectPropertyOperations.GetProperty( ObjVer, 1021)

'if ( Vault.ObjectPropertyOperations.GetProperty( ObjVer, 100).Value = 127) then 'pole recznego wpisywania numeru
		'	Err.Raise mfscriptcancel, "Offer deadline cant be greater than Tender closing date."

'	if ( Vault.ObjectPropertyOperations.GetProperty( ObjVer, 1451).Value > Vault.ObjectPropertyOperations.GetProperty( ObjVer, 1450).Value) then 'pole recznego wpisywania numeru
	'		Err.Raise mfscriptcancel, "Offer deadline cant be greater than Tender closing date."
	'	end if

'end if


	' ///////////////  FUNKCJA USTAWIANIA LICZNIKA WĄTKÓW  ///////////////////  

	nazwaZm = Cstr(objVer.Type)& "-" &Cstr(Year(now)) 'Nazwa zmiennej generowana na podstawie nazwy typu obiektu oraz aktualnego roku tworzenia obiektu
	licznik = VaultSharedVariables(nazwaZm)

	if IsNull(licznik) or IsEmpty(licznik) or licznik = "" Then ' wartość jest niezainicjowana lub zerowa
		VaultSharedVariables(nazwaZm) = 1
		licznik = 1
	end if
	dodajDoLog("test - " &licznik )
	' ///////////////  KONIEC FUNKCJI USTAWIANIA LICZNIKA WĄTKÓW  ///////////////////   	 

	' ///////////////  FUNKCJA ODCZYTU POSTFIXU  ///////////////////   
	
	Set oSCs = CreateObject("MFilesAPI.SearchConditions")
	Set oSC = CreateObject("MFilesAPI.SearchCondition")

	'The property "M-Files User" has the ID 1213
	oSC.Expression.SetPropertyValueExpression 1213,MFilesAPI.MFParentChildBehavior.MFParentChildBehaviorNone, Nothing
	oSC.ConditionType = MFilesAPI.MFConditionType.MFConditionTypeEqual
	oSC.TypedValue.SetValue MFilesAPI.MFDataType.MFDatatypeLookup, CurrentUserID
	oSCs.Add -1,oSC

	Dim oResults
	Set oResults = Vault.ObjectSearchOperations.SearchForObjectsByConditions(oSCs, MFilesAPI.MFSearchFlags.MFSearchFlagNone, False)
	
	If oResults.Count>0 Then
	dodajDoLog("search cond w if - " &licznik )
		Dim oEmp
		Set oEmp = oResults.Item(1)

		'get checked out Version & properties
		Set oEmp = Vault.ObjectOperations.GetLatestObjectVersionAndProperties(oEmp.ObjVer.ObjID,True)

		'get a property (id=1431)  POSTFIX KONTYNENTU                                  
		postfix = Vault.ObjectPropertyOperations.GetProperty(oEmp.ObjVer, 1431).TypedValue.GetValueAsLocalizedText()

	dodajDoLog(postfix)
	Else
		
	End If

	'//////////////////   KONIEC FUNKCJI ODCZYTU POSTFIXU   /////////////////




	Dim oPropValsNew
	Set oPropValsNew = CreateObject("MFilesAPI.PropertyValues")
			

	dodajDoLog("przed ustawieniem sterowania")
	Dim sterowanie1
	Dim sterowanie2
	
	'sterowanie2 = Vault.ObjectPropertyOperations.GetProperty( ObjVer, 1437)
	dodajDoLog("ustawianie sterowanie2 ")
'	if ( not (IsNull(Vault.ObjectPropertyOperations.GetProperty( ObjVer, 1437).Value) or IsEmpty(Vault.ObjectPropertyOperations.GetProperty( ObjVer, 1437).Value))) then
	if (Vault.ObjectPropertyOperations.GetProperty( ObjVer, 1437).Value = false)  then
		'Err.Raise mfscriptcancel, "wartośc combo false"
		if ( IsNull(Vault.ObjectPropertyOperations.GetProperty( ObjVer, 1436).Value) or IsEmpty(Vault.ObjectPropertyOperations.GetProperty( ObjVer, 1436).Value)) then 'pole recznego wpisywania numeru
			Err.Raise mfscriptcancel, "When 'Autonumber' field is set to 'NO', You need to provide manual number in form 'YEAR/XXX'."
		else

			Dim Temp_value
			Temp_value = Vault.ObjectPropertyOperations.GetProperty( ObjVer, 1436).Value
	
			Nazwa_watku = Temp_value

			Set myPropertyValue= CreateObject("MFilesAPI.PropertyValue")
			myPropertyValue.PropertyDef = 1020
			myPropertyValue.TypedValue.SetValue MFDatatypeText , Nazwa_watku
			oPropValsNew.Add -1, myPropertyValue

			'Set myPropertyValueACTION = CreateObject("MFilesAPI.PropertyValue")
			'myPropertyValueAction.PropertyDef = 11585
			'myPropertyValueAction.TypedValue.SetValue MFDatatypeLookup , 10
	
			'oPropValsNew.Add -1, myPropertyValueAction


			Vault.ObjectPropertyOperations.SetProperties ObjVer, oPropValsNew
			
			
			'Vault.ObjectPropertyOperations.SetProperty ObjVer,myPropertyValue

			dodajDoLog("Drugi if")
		End if
	else
	Nazwa_watku = Rok &"/"& licznik &"/"& postfix 
	
	
	Set myPropertyValue= CreateObject("MFilesAPI.PropertyValue")
	myPropertyValue.PropertyDef = 1020
	myPropertyValue.TypedValue.SetValue MFDatatypeText , Nazwa_watku

	oPropValsNew.Add -1, myPropertyValue

	'Set myPropertyValueAction= CreateObject("MFilesAPI.PropertyValue")
	'myPropertyValueAction.PropertyDef = 1565
	'myPropertyValueAction.TypedValue.SetValue MFDatatypeLookup , 2
	'oPropValsNew.Add -1, myPropertyValueAction
	
'	Set myPropertyValueACTION = CreateObject("MFilesAPI.PropertyValue")
'	myPropertyValueAction.PropertyDef = 11585
'	myPropertyValueAction.TypedValue.SetValue MFDatatypeLookup , 10
	
'	oPropValsNew.Add -1, myPropertyValueAction

	'Set myPropertyValueACTION = CreateObject("MFilesAPI.PropertyValue")
	'	myPropertyValueACTION.PropertyDef = 1565
	'	myPropertyValueACTION.TypedValue.SetValue MFDatatypeLookup , 1
	
	'	oPropValsNew.Add -1, myPropertyValueACTION

	Vault.ObjectPropertyOperations.SetProperties ObjVer, oPropValsNew
'	Vault.ObjectPropertyOperations.SetProperty ObjVer, myPropertyValue

	dodajDoLog(Nazwa_watku & " NAZWA KONCOWA ")

	dim n_licznik
	n_licznik = licznik+1
	VaultSharedVariables(nazwaZm) = n_licznik

	End if	

	
'----------------------------------------------
'dim ObiektWypluty_Manager
'Set ObiektWypluty_Manager = znajdzSupervisora (1455, ObjVer) 

'	Dim oManager	
'	Set oManager = Vault.ObjectPropertyOperations.GetProperty( ObiektWypluty_Manager.ObjVer, 1214).TypedValue.GetValueAsLookup

	'Create ObjID
'	dim oObjID: Set oObjID = CreateObject("MFilesAPI.ObjID")
'	oObjID.ID   =  oManager.item 'ID for related Member object
'	oObjID.Type =  oManager.ObjectType 'related member ObjType
	
	'Get Latest version of external object
'	dim oMember_ObjVer
'	set oMember_ObjVer = vault.ObjectOperations.GetLatestObjVer(oObjID,true)


'Dim oLookupManager: Set oLookupManager = CreateObject("MFilesAPI.Lookup")

'oLookupManager.ObjectType = 103 ' ID of your external object type
'oLookupManager.Item = oManager.item ' ID of specific object

'Dim oPropValManager: Set oPropValManager = PropertyValues.SearchForProperty( 1162 )
'oPropValManager.Value.SetValueToLookup oLookupManager
'Vault.ObjectPropertyOperations.SetProperty ObjVer, oPropValManager
'dodajDoLog("Ustawiłem Managera")

    'Get the value as lookup reference.
    'Dim oSupervisor	
'	Set oSupervisor = Vault.ObjectPropertyOperations.GetProperty( oMember_ObjVer, 1214).TypedValue.GetValueAsLookup

'Dim oLookupSupervisor: Set oLookupSupervisor = CreateObject("MFilesAPI.Lookup")
'oLookupSupervisor.ObjectType = 103 ' ID of your external object type
'oLookupSupervisor.Item = oSupervisor.item ' ID of specific object

'Dim oPropValSupervisor: Set oPropValSupervisor = PropertyValues.SearchForProperty( 1456 )
'oPropValSupervisor.Value.SetValueToLookup oLookupSupervisor
'Vault.ObjectPropertyOperations.SetProperty ObjVer, oPropValSupervisor
'dodajDoLog("Ustawiłem Supervisora")

'------------------------------------------
end if 'koniec ifa sprawdzającego czy WĄTEK SPRZEDAŻOWY
	
	dodajDoLog("SalesLead_Name - leave")
	dodajDoLog("---------------------------")


'   ZAPIS DANYCH DO LOGA   ////////////////////
sub dodajDoLog(message)
		ForAppending = 8
		 
		set objFSO = CreateObject("Scripting.FileSystemObject")
		set objFile = objFSO.OpenTextFile("E:\test\SalesLead_create.txt", ForAppending, True)
	 
		objFile.WriteLine(message)
		objFile.Close
end sub

'-----------------------------------------------------------------------------------------------------------------------------------------------------


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

'-------