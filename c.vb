	dodajDoLog("---------------------------")
	dodajDoLog("CreateBinders - enter")

Dim przelacznik
przelacznik = Cint(ObjVer.Type)
Dim oLookUpRef
Select Case przelacznik

	Case 4 'Production

	Case 14 'Dokumenty QA

	Case 182 'Dokumenty Sales&Marketing
	
	Dim PropertyValues 
	Set PropertyValues = Vault.ObjectPropertyOperations.GetProperties( ObjVer,False )
	Dim Meta_REP
	Meta_REP = PropertyValues.SearchForProperty( 1459 ).TypedValue.GetLookupID    'alternative code .TypeValue.GetValueAsLookup().Item
	dodajDoLog("NOWY LEAD")
	DIM ID_Customer, ID_Representative, ID_Country
	ID_Customer = PropertyValues.SearchForProperty( 1457 ).TypedValue.GetLookupID
	ID_Representative = PropertyValues.SearchForProperty( 1459 ).TypedValue.GetLookupID
	ID_Country = PropertyValues.SearchForProperty( 1507 ).TypedValue.GetLookupID
	
	if IsNull(Meta_REP) Then
		ID_Representative = 0
	End if 
		
		
		
		
		
		NEW_SalesLead (ID_Customer, ID_Representative, ID_Country)
		



	Case 183 ' Nowy obiekt typu DEVELOPMENT
	
		Call NEW_Development ()	

	Case 206 ' Nowy obiekt typu Zlecenie na TEST
	
		kategoria_test = Vault.ObjectPropertyOperations.GetProperty( objVer, 100).TypedValue.GetLookupID
'err.raise mfscriptcancel, "Kategoria " & kategoria_test
		Select Case kategoria_test
			Case 161 'FTR
				oLookUpRef = Vault.ObjectPropertyOperations.GetProperty( objVer, 11609).TypedValue.GetLookupID	
				'err.raise mfscriptcancel, "id dla U " & oLookUpRef
				
				Select Case oLookUpRef
					
					Case 1 'Acu
						Call NEW_FTR_TEST (1, oLookUpRef, 23605)

					Case 2 'Astel-1
						Call NEW_FTR_TEST (2, oLookUpRef, 24018)
	
					Case 3 'Astel-3
						Call NEW_FTR_TEST (2, oLookUpRef, 24018)
					
					Case 4 'RD
					'	Call NEW_FTR_TEST (1, oLookUpRef)
	
					Case 5 'CTS
						Call NEW_FTR_TEST (2, oLookUpRef, 22434)

					Case 6 'CIS
						Call NEW_FTR_TEST (2, oLookUpRef, 22438)

					Case 7 'VTS
						Call NEW_FTR_TEST (2, oLookUpRef, 22435)
						
					Case 8 'VIS
						Call NEW_FTR_TEST (2, oLookUpRef, 22444)
	
					Case 9 'PS-1
						'Call NEW_FTR_TEST (1, oLookUpRef)
					
					Case 10 'PS-3
						'Call NEW_FTR_TEST (1, oLookUpRef)
					
					Case 11 'Inne
						'Call NEW_FTR_TEST (1, oLookUpRef)
					
					Case 12 'SR
						'Call NEW_FTR_TEST (1, oLookUpRef)
					
					Case 13 'Components
						'Call NEW_FTR_TEST (1, oLookUpRef)
					
					Case 14 'OPTO
						'Call NEW_FTR_TEST (1, oLookUpRef)
					
					Case 15 'Software
						'Call NEW_FTR_TEST (1, oLookUpRef)
					
					Case 16 'Montaż SMT
						'Call NEW_FTR_TEST (1, oLookUpRef)
						
				End Select
				
			Case 164 'FCT
			Case 165 'FAT
			Case 166 'SAT
				oLookUpRef = Vault.ObjectPropertyOperations.GetProperty( objVer, 11609).TypedValue.GetLookupID	
				
				Call NEW_SAT_Test( 4, oLookUpRef, 24194)	
				'err.raise mfscriptcancel, "BRAKE end"
			Case 175 'SCT
			Case 176 'STR
			
					
		'NEW_FTR_TEST ()
		End Select

End Select

	dodajDoLog("CreateBinders - leave")
	dodajDoLog("---------------------------")

	'--------------------------------------------------------------------
		'--------------------------------------------------------------------
			'--------------------------------------------------------------------
				'--------------------------------------------------------------------
					'--------------------------------------------------------------------
						'--------------------------------------------------------------------
							'--------------------------------------------------------------------

sub dodajDoLog(message)
		ForAppending = 8
		 
		set objFSO = CreateObject("Scripting.FileSystemObject")
		set objFile = objFSO.OpenTextFile("E:\test\Developent_DODAJ_PLIKI.txt", ForAppending, True)
	 
		objFile.WriteLine(message)
		objFile.Close
end sub

	'--------------------------------------------------------------------
		'--------------------------------------------------------------------
			'--------------------------------------------------------------------
				'--------------------------------------------------------------------
					'--------------------------------------------------------------------
						'--------------------------------------------------------------------
							'--------------------------------------------------------------------

Sub NEW_Development()
	Dim idClass_Development: idClass_Development = 1486 ' Development property definition ID
	Dim idType_InputDoc: idType_InputDoc = 9 ' Input_documents Object Type ID
	Dim idClass_InputDoc: idClass_InputDoc = 132 ' Input_documents Class ID
	Dim idClass_OutputDoc: idClass_OutputDoc = 133 ' output_documents Class ID
	Dim idClass_CatalogElem: idClass_CatalogElem = 134 ' Catalog Elementow Class ID
	Dim idClass_TestResult: idClass_TestResult = 137 ' Test result Class ID
	
	Dim idType_KartRozw: idType_KartRozw = 0 ' Karta rozwoju Object Type ID
	Dim idClass_KartaRozw: idClass_KartaRozw = 43 ' Karta rozwoju Class ID
	
	' Create a new INPUT_Documents object
	Dim oPropVals_Input: Set oPropVals_Input = CreateObject("MFilesAPI.PropertyValues")
	Dim oOnePropVal_Input: Set oOnePropVal_Input = CreateObject("MFilesAPI.PropertyValue")
	
	' Class = Input Documents
	oOnePropVal_Input.PropertyDef = 100
	oOnePropVal_Input.TypedValue.SetValue MFDatatypeLookup, idClass_InputDoc
	oPropVals_Input.Add -1, oOnePropVal_Input
	
	' Development
	oOnePropVal_Input.PropertyDef = idClass_Development
	oOnePropVal_Input.TypedValue.SetValue MFDatatypeMultiSelectLookup, ObjVer.ID
	oPropVals_Input.Add -1, oOnePropVal_Input
	
	'--------------------------------------------------------------------
	
	' Create a new Output_Documents object
	Dim oPropVals_Output: Set oPropVals_Output = CreateObject("MFilesAPI.PropertyValues")
	Dim oOnePropVal_Output: Set oOnePropVal_Output = CreateObject("MFilesAPI.PropertyValue")
	
	' Class = Output Documents
	oOnePropVal_Output.PropertyDef = 100
	oOnePropVal_Output.TypedValue.SetValue MFDatatypeLookup, idClass_OutputDoc
	oPropVals_Output.Add -1, oOnePropVal_Output
	
	'Development
	oPropVals_Output.Add -1, oOnePropVal_Input
	'--------------------------------------------------------------------
	
	'Create a new Catalog of elements object
	Dim oPropVals_Catalog: Set oPropVals_Catalog = CreateObject("MFilesAPI.PropertyValues")
	Dim oOnePropVal_Catalog: Set oOnePropVal_Catalog = CreateObject("MFilesAPI.PropertyValue")
	
	'Class = Catalog of elements
	oOnePropVal_Catalog.PropertyDef = 100
	oOnePropVal_Catalog.TypedValue.SetValue MFDatatypeLookup, idClass_CatalogElem
	oPropVals_Catalog.Add -1, oOnePropVal_Catalog
	
	'Development
	oPropVals_Catalog.Add -1, oOnePropVal_Input
	'--------------------------------------------------------------------
	
	'Create a new Test Results object
	Dim oPropVals_Tests: Set oPropVals_Tests = CreateObject("MFilesAPI.PropertyValues")
	Dim oOnePropVal_Tests: Set oOnePropVal_Tests = CreateObject("MFilesAPI.PropertyValue")
	
	' Class = Test Results
	oOnePropVal_Tests.PropertyDef = 100
	oOnePropVal_Tests.TypedValue.SetValue MFDatatypeLookup, idClass_TestResult
	oPropVals_Tests.Add -1, oOnePropVal_Tests
	
	'Development
	oPropVals_Tests.Add -1, oOnePropVal_Input
	'--------------------------------------------------------------------
	
	' Create a new FB 1.1 KARTA ROZWOJU object
	Dim oPropVals_KartaRoz: Set oPropVals_KartaRoz = CreateObject("MFilesAPI.PropertyValues")
	Dim oPropVals_Overview: Set oPropVals_Overview = CreateObject("MFilesAPI.PropertyValues")
	Dim oPropVals_Veryfication: Set oPropVals_Veryfication = CreateObject("MFilesAPI.PropertyValues")
	Dim oPropVals_Validation: Set oPropVals_Validation = CreateObject("MFilesAPI.PropertyValues")
	Dim oPropVals_Decision: Set oPropVals_Decision = CreateObject("MFilesAPI.PropertyValues")

	Dim oOnePropVal_KartaRoz: Set oOnePropVal_KartaRoz = CreateObject("MFilesAPI.PropertyValue")
	Dim oOnePropVal_Overview: Set oOnePropVal_Overview = CreateObject("MFilesAPI.PropertyValue")
	Dim oOnePropVal_Veryfication: Set oOnePropVal_Veryfication = CreateObject("MFilesAPI.PropertyValue")
	Dim oOnePropVal_Validation: Set oOnePropVal_Validation = CreateObject("MFilesAPI.PropertyValue")
	Dim oOnePropVal_Decision: Set oOnePropVal_Decision = CreateObject("MFilesAPI.PropertyValue")
'	Dim ObjVer_KartaRoz: Set ObjVer_KartaRoz = CreateObject("MFilesAPI.PropertyValue")
	
	' Class = Test Results
	oOnePropVal_KartaRoz.PropertyDef = 100
	oOnePropVal_KartaRoz.TypedValue.SetValue MFDatatypeLookup, idClass_KartaRozw
	oPropVals_KartaRoz.Add -1, oOnePropVal_KartaRoz

'	oOnePropVal_Overview.TypedValue.SetValue MFDatatypeLookup, 44
'	oPropVals_Overview.Add -1, oOnePropVal_Overview
'
'	oOnePropVal_Veryfication.TypedValue.SetValue MFDatatypeLookup, 46
'	oPropVals_Veryfication.Add -1, oOnePropVal_Veryfication

''	oOnePropVal_Validation.TypedValue.SetValue MFDatatypeLookup, 47
'	oPropVals_Validation.Add -1, oOnePropVal_Validation
	oOnePropVal_Decision.PropertyDef = 100
	oOnePropVal_Decision.TypedValue.SetValue MFDatatypeLookup, 49
	oPropVals_Decision.Add -1, oOnePropVal_Decision
	
	' Development
	oPropVals_KartaRoz.Add -1, oOnePropVal_Input
'	oPropVals_Overview.Add -1, oOnePropVal_Input
	'oPropVals_Veryfication.Add -1, oOnePropVal_Input
	'oPropVals_Validation.Add -1, oOnePropVal_Input
	oPropVals_Decision.Add -1, oOnePropVal_Input

	'--------------------------------------------------------------------

	
	Call Add_SFD_from_template (oPropVals_KartaRoz, 10852)
'	Call Add_SFD_from_template (oPropVals_Overview, 28269)
'	Call Add_SFD_from_template (oPropVals_Veryfication, 28270)
'	Call Add_SFD_from_template (oPropVals_Validation, 28271)
	Call Add_SFD_from_template (oPropVals_Decision, 28272)



	' Files
	Dim oACL: Set oACL = CreateObject("MFilesAPI.AccessControlList")
	set oACL = nothing
	
	Dim oFiles: Set oFiles = CreateObject("MFilesAPI.SourceObjectFiles")
	
	Call Vault.ObjectOperations.CreateNewObjectEx(idType_InputDoc, oPropVals_Input, oFiles, False, True, oACL)
	Call Vault.ObjectOperations.CreateNewObjectEx(idType_InputDoc, oPropVals_Output, oFiles, False, True, oACL)
	Call Vault.ObjectOperations.CreateNewObjectEx(idType_InputDoc, oPropVals_Catalog, oFiles, False, True, oACL)		
	Call Vault.ObjectOperations.CreateNewObjectEx(idType_InputDoc, oPropVals_Tests, oFiles, False, True, oACL)

End Sub

	'--------------------------------------------------------------------
		'--------------------------------------------------------------------
			'--------------------------------------------------------------------
				'--------------------------------------------------------------------
					'--------------------------------------------------------------------
						'--------------------------------------------------------------------
							'--------------------------------------------------------------------

Sub NEW_SalesLead(ID_Customer, ID_Rep, ID_Country)

	dodajDoLog("-----NEW_SalesLead function--------")

'err.raise mfscriptcancel, "Error: No such !Template found on no files in template."
	Dim oLookupsProfilKlienta : Set oLookups = CreateObject("MFilesAPI.Lookups")
	'set oLookupsProfilKlienta = WyszukajProfilKlienta(ID_Customer, false)
	if WyszukajProfilKlienta(ID_Customer, ID_Rep, ID_Country, false) = true then


		Dim idType_DocBinder : idType_DocBinder  = 9 		' ID Typu Obiektu - Document Binder 
		Dim idClass_SalesLead: idClass_SalesLead = 1484 ' ID metadanej Sales Lead 
		Dim idClass_Customer : idClass_Customer  = 1457 ' ID metadanej Customer 
		
		Dim idClass_ProfilKlienta : idClass_ProfilKlienta  = 205 	' ID Klasy Obiektu (kategoria - Profil Klienta)
	
		'--------------------------------------------------------------------
		
		' Create a new DOCUMENT BINDER object
		Dim oPropVals_BINDER: Set oPropVals_BINDER = CreateObject("MFilesAPI.PropertyValues")
		Dim oOnePropVal_BINDER: Set oOnePropVal_BINDER = CreateObject("MFilesAPI.PropertyValue")
		'Dim looks: Set looks = CreateObject("MFilesAPI.PropertyValue")
			
		 'Klasa = Profil Klienta
		oOnePropVal_BINDER.PropertyDef = 100
		oOnePropVal_BINDER.TypedValue.SetValue MFDatatypeLookup, idClass_ProfilKlienta
		oPropVals_BINDER.Add -1, oOnePropVal_BINDER
	
		'METADANA - Name or Title 
		oOnePropVal_BINDER.PropertyDef = 0
		oOnePropVal_BINDER.TypedValue.SetValue MFDatatypeText, "Profil Klienta - Prosze się zapoznać!"
		oPropVals_BINDER.Add -1, oOnePropVal_BINDER
		
		' Sales Lead
		oOnePropVal_BINDER.PropertyDef = 1484
		oOnePropVal_BINDER.TypedValue.SetValue MFDatatypeMultiSelectLookup, ObjVer.ID
		oPropVals_BINDER.Add -1, oOnePropVal_BINDER

		' Collection members
		oOnePropVal_BINDER.PropertyDef = 46
		oOnePropVal_BINDER.TypedValue.SetValueToMultiSelectLookup WyszukajProfilKlienta(ID_Customer, ID_Rep, ID_Country, true)
		oPropVals_BINDER.Add -1, oOnePropVal_BINDER

		' WorkFlow
		oOnePropVal_BINDER.PropertyDef = 38
		oOnePropVal_BINDER.TypedValue.SetValue MFDatatypeLookup, 146
		oPropVals_BINDER.Add -1, oOnePropVal_BINDER
		
		' Workflow State
		oOnePropVal_BINDER.PropertyDef = 39
		oOnePropVal_BINDER.TypedValue.SetValue MFDatatypeLookup, 361
		oPropVals_BINDER.Add -1, oOnePropVal_BINDER
	
		'--------------------------------------------------------------------
	
		' Files
		'Dim oACL: Set oACL = CreateObject("MFilesAPI.AccessControlList")
		'set oACL = nothing
		
		'Dim oFiles: Set oFiles = CreateObject("MFilesAPI.SourceObjectFiles")
		
		'Call Vault.ObjectOperations.CreateNewObjectEx(idType_DocBinder, oPropVals_Input, oFiles, False, True, oACL)
		'Call Vault.ObjectOperations.CreateNewObjectEx(idType_DocBinder, oPropVals_Output, oFiles, False, True, oACL)
		'Call Vault.ObjectOperations.CreateNewObjectEx(idType_DocBinder, oPropVals_Catalog, oFiles, False, True, oACL)	
		'err.raise mfscriptcancel, "doszło do tąd"
		Call Vault.ObjectOperations.CreateNewObjectEx(9, oPropVals_BINDER, nothing, False, True, nothing)

	end if

	

End Sub

	'--------------------------------------------------------------------
		'--------------------------------------------------------------------
			'--------------------------------------------------------------------
				'--------------------------------------------------------------------
					'--------------------------------------------------------------------
						'--------------------------------------------------------------------
							'--------------------------------------------------------------------

Sub NEW_FTR_Test(ID_RaportDotyczy, ID_TypProduktu, ID_template)
	Dim idType_Doc				: idType_Doc			 = 0		' ID Typu Obiektu - Document Binder 
	Dim idClass_Badanie			: idClass_Badanie		 = 11607	' ID metadanej Sales Lead 
	Dim idClass_RaportDotyczy	: idClass_RaportDotyczy  = 11657	' ID metadanej Customer 
	
	Dim idClass_RaportQA		: idClass_RaportQA		 = 182 		' ID Klasy Obiektu (kategoria - Profil Klienta)
	Dim idClass_WynikiBadanQA	: idClass_WynikiBadanQA  = 184 		' ID Klasy Obiektu (kategoria - Profil Klienta)

	'--------------------------------------------------------------------
	
	' Create a new DOCUMENT BINDER object
	Dim oPropVals_BINDER: Set oPropVals_BINDER = CreateObject("MFilesAPI.PropertyValues")
	Dim oPropVals_BINDER2: Set oPropVals_BINDER2 = CreateObject("MFilesAPI.PropertyValues")
	Dim oPropVals_BINDER3: Set oPropVals_BINDER3 = CreateObject("MFilesAPI.PropertyValues")
	Dim oPropVals_BINDER4: Set oPropVals_BINDER4 = CreateObject("MFilesAPI.PropertyValues")

	Dim oOnePropVal_BINDER: Set oOnePropVal_BINDER = CreateObject("MFilesAPI.PropertyValue")
	Dim oOnePropVal_BINDER2: Set oOnePropVal_BINDER2 = CreateObject("MFilesAPI.PropertyValue")
	Dim oOnePropVal_BINDER3: Set oOnePropVal_BINDER3 = CreateObject("MFilesAPI.PropertyValue")
	Dim oOnePropVal_BINDER4: Set oOnePropVal_BINDER4 = CreateObject("MFilesAPI.PropertyValue")

	'Dim looks: Set looks = CreateObject("MFilesAPI.PropertyValue")
	
	 'Klasa = Raport QA
	oOnePropVal_BINDER.PropertyDef = 100
	oOnePropVal_BINDER.TypedValue.SetValue MFDatatypeLookup, idClass_RaportQA
	oPropVals_BINDER.Add -1, oOnePropVal_BINDER
	oPropVals_BINDER3.Add -1, oOnePropVal_BINDER


	'Klasa = Wyniki Badań QA
	oOnePropVal_BINDER2.PropertyDef = 100
	oOnePropVal_BINDER2.TypedValue.SetValue MFDatatypeLookup, idClass_WynikiBadanQA
	oPropVals_BINDER2.Add -1, oOnePropVal_BINDER2
	
	'Klasa = Zdjęcia QA
	oOnePropVal_BINDER4.PropertyDef = 100
	oOnePropVal_BINDER4.TypedValue.SetValue MFDatatypeLookup, 221
	oPropVals_BINDER4.Add -1, oOnePropVal_BINDER4


	' Badanie
	oOnePropVal_BINDER.PropertyDef = 11607
	oOnePropVal_BINDER.TypedValue.SetValue MFDatatypeMultiSelectLookup, ObjVer.ID
	oPropVals_BINDER.Add -1, oOnePropVal_BINDER
	oPropVals_BINDER3.Add -1, oOnePropVal_BINDER
	oPropVals_BINDER4.Add -1, oOnePropVal_BINDER

	' Badanie
	oOnePropVal_BINDER2.PropertyDef = 11607
	oOnePropVal_BINDER2.TypedValue.SetValue MFDatatypeMultiSelectLookup, ObjVer.ID
	oPropVals_BINDER2.Add -1, oOnePropVal_BINDER2

	' Typ produktu
	oOnePropVal_BINDER.PropertyDef = 11609
	oOnePropVal_BINDER.TypedValue.SetValue MFDatatypeLookup, ID_TypProduktu
	oPropVals_BINDER.Add -1, oOnePropVal_BINDER
	
	' Raport dotyczy (BADANIA)
	oOnePropVal_BINDER.PropertyDef = 11657
	oOnePropVal_BINDER.TypedValue.SetValue MFDatatypeMultiSelectLookup, ID_RaportDotyczy
	oPropVals_BINDER.Add -1, oOnePropVal_BINDER

	' Raport dotyczy (URUCHOMIENIA)
	oOnePropVal_BINDER3.PropertyDef = 11657
	oOnePropVal_BINDER3.TypedValue.SetValue MFDatatypeMultiSelectLookup, 1
	oPropVals_BINDER3.Add -1, oOnePropVal_BINDER3


		' WorkFlow
'	oOnePropVal_BINDER.PropertyDef = 38
'	oOnePropVal_BINDER.TypedValue.SetValue MFDatatypeLookup, 134
'	oPropVals_BINDER.Add -1, oOnePropVal_BINDER
'	oPropVals_BINDER2.Add -1, oOnePropVal_BINDER
'	oPropVals_BINDER3.Add -1, oOnePropVal_BINDER
'	oPropVals_BINDER4.Add -1, oOnePropVal_BINDER
	
	' Workflow State
'	oOnePropVal_BINDER.PropertyDef = 39
'	oOnePropVal_BINDER.TypedValue.SetValue MFDatatypeLookup, 287
'	oPropVals_BINDER.Add -1, oOnePropVal_BINDER
'	oPropVals_BINDER2.Add -1, oOnePropVal_BINDER
'	oPropVals_BINDER3.Add -1, oOnePropVal_BINDER
'	oPropVals_BINDER4.Add -1, oOnePropVal_BINDER

	'--------------------------------------------------------------------
'err.raise mfscriptcancel, "id dla U " & oLookUpRef
	Add_SFD_from_template oPropVals_BINDER , ID_template
	
	if ID_TypProduktu >=2 AND ID_TypProduktu <=8 then
		Add_MFD_Empty (oPropVals_BINDER2)		
		Add_SFD_from_template oPropVals_BINDER3 , 23605

		if ID_TypProduktu =6 OR ID_TypProduktu =8 then
			Add_MFD_Empty (oPropVals_BINDER4)	
		End if
	End if
End Sub

	'--------------------------------------------------------------------
		'--------------------------------------------------------------------
			'--------------------------------------------------------------------
				'--------------------------------------------------------------------
					'--------------------------------------------------------------------
						'--------------------------------------------------------------------
							'--------------------------------------------------------------------

Sub NEW_SAT_Test(ID_RaportDotyczy, ID_TypProduktu, ID_template)
	Dim idType_Doc				: idType_Doc			 = 0		' ID Typu Obiektu - Document Binder 
	Dim idClass_Badanie			: idClass_Badanie		 = 11607	' ID metadanej Sales Lead 
	Dim idClass_RaportDotyczy	: idClass_RaportDotyczy  = 11657	' ID metadanej Customer 
	
	Dim idClass_RaportQA		: idClass_RaportQA		 = 182 		' ID Klasy Obiektu (kategoria - Profil Klienta)
	Dim idClass_WynikiBadanQA	: idClass_WynikiBadanQA  = 184 		' ID Klasy Obiektu (kategoria - Profil Klienta)

	'--------------------------------------------------------------------
	
	' Create a new DOCUMENT BINDER object
	Dim oPropVals_BINDER: Set oPropVals_BINDER = CreateObject("MFilesAPI.PropertyValues")
	Dim oOnePropVal_BINDER: Set oOnePropVal_BINDER = CreateObject("MFilesAPI.PropertyValue")
	'Dim looks: Set looks = CreateObject("MFilesAPI.PropertyValue")
	
	 'Klasa = Raport QA
	oOnePropVal_BINDER.PropertyDef = 100
	oOnePropVal_BINDER.TypedValue.SetValue MFDatatypeLookup, idClass_RaportQA
	oPropVals_BINDER.Add -1, oOnePropVal_BINDER

	' Badanie
	oOnePropVal_BINDER.PropertyDef = 11607
	oOnePropVal_BINDER.TypedValue.SetValue MFDatatypeMultiSelectLookup, ObjVer.ID
	oPropVals_BINDER.Add -1, oOnePropVal_BINDER

	' Typ produktu
	oOnePropVal_BINDER.PropertyDef = 11609
	oOnePropVal_BINDER.TypedValue.SetValue MFDatatypeLookup, ID_TypProduktu
	oPropVals_BINDER.Add -1, oOnePropVal_BINDER
	
	' Raport dotyczy
	oOnePropVal_BINDER.PropertyDef = 11657
	oOnePropVal_BINDER.TypedValue.SetValue MFDatatypeMultiSelectLookup, ID_RaportDotyczy
	oPropVals_BINDER.Add -1, oOnePropVal_BINDER

	' WorkFlow
	oOnePropVal_BINDER.PropertyDef = 38
	oOnePropVal_BINDER.TypedValue.SetValue MFDatatypeLookup, 134
	oPropVals_BINDER.Add -1, oOnePropVal_BINDER
	
	' Workflow State
	oOnePropVal_BINDER.PropertyDef = 39
	oOnePropVal_BINDER.TypedValue.SetValue MFDatatypeLookup, 287
	oPropVals_BINDER.Add -1, oOnePropVal_BINDER

	'--------------------------------------------------------------------
'err.raise mfscriptcancel, "id dla U " & oLookUpRef
	Add_SFD_from_template oPropVals_BINDER , ID_template
'	err.raise mfscriptcancel, "BRAKE"
	'if ID_TypProduktu >=5 AND ID_TypProduktu <=8 then
	'	Add_MFD_Empty (oPropVals_BINDER2)
	'End if
End Sub

	'--------------------------------------------------------------------
		'--------------------------------------------------------------------
			'--------------------------------------------------------------------
				'--------------------------------------------------------------------
					'--------------------------------------------------------------------
						'--------------------------------------------------------------------
							'--------------------------------------------------------------------


Sub Add_SFD_from_template (Bind_Properties , Template_ID)	
	
	Dim oACL: Set oACL = CreateObject("MFilesAPI.AccessControlList")
	set oACL = nothing
	
	
	'--------------------------------------------------------------------
	
	
	'znajdowanie templejta
	Dim oLookupObj: set oLookupObj = CreateObject("MFilesAPI.ObjVer")
	oLookupObj.SetIDs 0, Template_ID, -1 


	'Pobieranie pliku z templejta do folderu tymczasowego
	Dim oObjectInfo : Set oObjectInfo = CreateObject("MFilesAPI.ObjectVersion")
	Set oObjectInfo = Vault.ObjectOperations.GetObjectInfo(oLookupObj, True, False)
	Dim oObjectFile : Set oObjectFile = CreateObject("MFilesAPI.ObjectFile")
	'Dim oObjectFiles : Set oObjectFiles = CreateObject("MFilesAPI.ObjectFiles")				'This operation is not allowed for ObjectFiles, but not necessary either (nor is it above!)
	Dim oObjectFiles : Set oObjectFiles = Vault.ObjectFileOperations.GetFiles(oObjectInfo.ObjVer)		'The GetFiles function will create the Files collection this way :-)
	
	If oObjectFiles.Count = 0 Then
		err.raise mfscriptcancel, "Error: No such !Template found"
	Else
		Set oObjectFile = oObjectFiles.Item(1)			
		Dim szExt : szExt = oObjectFile.Extension 
		Dim szName : szName = oObjectFile.GetNameForFileSystem
		Dim szID : szID = oObjectFile.ID
		Dim szVersion : szVersion = oObjectFile.Version
		Dim szPath: szPath = "E:\test\TEMP\"	& szName
		Vault.ObjectFileOperations.DownloadFile szID, szVersion, szPath					' Download the file. The path must be available on the server!

		'Upload nowego pliku na serwer i zakładanie nowego dokumentu
		Dim oFiles1: Set oFiles1 = CreateObject("MFilesAPI.SourceObjectFiles")
		Dim oSourceFile1: Set oSourceFile1 = CreateObject("MFilesAPI.SourceObjectFile")
		oSourceFile1.SourceFilePath = szPath
		oSourceFile1.Title = "demo"
		oSourceFile1.Extension = oObjectFile.Extension
		oFiles1.Add -1, oSourceFile1

		Call Vault.ObjectOperations.CreateNewObjectEx(0, Bind_Properties, oFiles1, true, True, oACL)

		'Kasowanie pliku tymczasowego
		Dim fso: Set fso = CreateObject("Scripting.FileSystemObject")
		fso.DeleteFile szPath
	end if

End Sub

	'--------------------------------------------------------------------
		'--------------------------------------------------------------------
			'--------------------------------------------------------------------
				'--------------------------------------------------------------------
					'--------------------------------------------------------------------
						'--------------------------------------------------------------------
							'--------------------------------------------------------------------

Sub Add_MFD_from_template (Bind_Properties , Template_ID)	
	
	Dim oACL: Set oACL = CreateObject("MFilesAPI.AccessControlList")
	set oACL = nothing
		
	'--------------------------------------------------------------------	
	
	'znajdowanie templejta
	Dim oLookupObj: set oLookupObj = CreateObject("MFilesAPI.ObjVer")
	oLookupObj.SetIDs 0, Template_ID, -1 

	'Pobieranie pliku z templejta do folderu tymczasowego
	Dim oObjectInfo : Set oObjectInfo = CreateObject("MFilesAPI.ObjectVersion")
	Set oObjectInfo = Vault.ObjectOperations.GetObjectInfo(oLookupObj, True, False)
	Dim oObjectFile : Set oObjectFile = CreateObject("MFilesAPI.ObjectFile")
	'Dim oObjectFiles : Set oObjectFiles = CreateObject("MFilesAPI.ObjectFiles")				'This operation is not allowed for ObjectFiles, but not necessary either (nor is it above!)
	Dim oObjectFiles : Set oObjectFiles = Vault.ObjectFileOperations.GetFiles(oObjectInfo.ObjVer)		'The GetFiles function will create the Files collection this way :-)
	
	If oObjectFiles.Count = 0 Then
		err.raise mfscriptcancel, "Error: No such !Template found on no files in template."
	Else
	'	Dim NowyDokumentMFD
	'	Set MFD = Vault.ObjectOperations.CreateNewObjectEx(9, Bind_Properties, oFiles1, False, True, oACL)'
		Dim oFiles1: Set oFiles1 = CreateObject("MFilesAPI.SourceObjectFiles")

		Dim adresy_plikow()
		ReDim adresy_plikow(0)

		ReDim Preserve numbers(WskaznikTablicy+1)

		For i = 1 To oObjectFiles.Count

			ReDim Preserve adresy_plikow(i)
			adresy_plikow(i)=szPath


			Set oObjectFile = oObjectFiles.Item(i)			
			Dim szExt : szExt = oObjectFile.Extension 
			Dim szName : szName = oObjectFile.GetNameForFileSystem
			Dim szID : szID = oObjectFile.ID
			Dim szVersion : szVersion = oObjectFile.Version
			Dim szPath: szPath = "E:\test\TEMP\"	& szName
			adresy_plikow(i)=szPath
			Vault.ObjectFileOperations.DownloadFile szID, szVersion, szPath					' Download the file. The path must be available on the server!

			'Upload nowego pliku na serwer i zakładanie nowego dokumentu
			
			Dim oSourceFile1: Set oSourceFile1 = CreateObject("MFilesAPI.SourceObjectFile")
			oSourceFile1.SourceFilePath = szPath
			oSourceFile1.Title = oObjectFile.Title
			oSourceFile1.Extension = oObjectFile.Extension
			oFiles1.Add -1, oSourceFile1
			dodajDoLog(oObjectFile.GetNameForFileSystem)
			
		Next

		Call Vault.ObjectOperations.CreateNewObjectEx(0, Bind_Properties, oFiles1, False, True, oACL)
			
		'Kasowanie plików tymczasowych
		Dim fso: Set fso = CreateObject("Scripting.FileSystemObject")
		For i = 1 To UBound(adresy_plikow) 
	
			fso.DeleteFile adresy_plikow(i)

		Next
	end if

End Sub

	'--------------------------------------------------------------------
		'--------------------------------------------------------------------
			'--------------------------------------------------------------------
				'--------------------------------------------------------------------
					'--------------------------------------------------------------------
						'--------------------------------------------------------------------
							'--------------------------------------------------------------------

Sub Add_MFD_Empty (Bind_Properties)	
	
Call Vault.ObjectOperations.CreateNewObjectEx(0, Bind_Properties, nothing, False, True, nothing)

End Sub

	'--------------------------------------------------------------------
		'--------------------------------------------------------------------
			'--------------------------------------------------------------------
				'--------------------------------------------------------------------
					'--------------------------------------------------------------------
						'--------------------------------------------------------------------
							'--------------------------------------------------------------------

Function WyszukajProfilKlienta(CustomerID, RepID, CountryID, sterownik)
	
	Const iOTDocument = 0 'Builtin, do not change
	Const iPDClass  = 100 'Builtin, do not change
	Const iClassManual  = 205 'ID kategorii dokumentu "0-Unclassified document"

	' Build the search conditions
	Dim oOneSC: Set oOneSC = CreateObject("MFilesAPI.SearchCondition")
	Dim oSCs_customer: Set oSCs_customer = CreateObject("MFilesAPI.SearchConditions")
	Dim oSCs_representative: Set oSCs_representative = CreateObject("MFilesAPI.SearchConditions")
	Dim oSCs_country: Set oSCs_country = CreateObject("MFilesAPI.SearchConditions")

	' Deleted = no
	oOneSC.ConditionType = MFConditionTypeEqual
	oOneSC.Expression.DataStatusValueType = MFStatusTypeDeleted
	oOneSC.TypedValue.SetValue MFDatatypeBoolean, False
	oSCs_customer.Add -1, oOneSC
	oSCs_representative.Add -1, oOneSC
	oSCs_country.Add -1, oOneSC

	' Object type = Document
	oOneSC.ConditionType = MFConditionTypeEqual
	oOneSC.Expression.DataStatusValueType = MFStatusTypeObjectTypeID
	oOneSC.TypedValue.SetValue MFDatatypeLookup, iOTDocument
	oScs_customer.Add -1, oOneSC
	oSCs_representative.Add -1, oOneSC
	oSCs_country.Add -1, oOneSC
	
'	' Class = Manuals (general)
	oOneSC.ConditionType = MFConditionTypeEqual
	oOneSC.Expression.DataPropertyValuePropertyDef = iPDClass
	oOneSC.TypedValue.SetValue MFDatatypeLookup, iClassManual
	oScs_customer.Add -1, oOneSC
	oSCs_representative.Add -1, oOneSC
	oSCs_country.Add -1, oOneSC
	
	' Customer = CustomerID
	oOneSC.ConditionType = MFConditionTypeEqual
	oOneSC.Expression.SetPropertyValueExpression 1457,  MFilesAPI.MFParentChildBehavior.MFParentChildBehaviorNone, Nothing
	oOneSC.TypedValue.SetValue MFilesAPI.MFDataType.MFDatatypeLookup,  CustomerID
	oScs_customer.Add -1, oOneSC

	' Representative = RepresentativeID
	oOneSC.ConditionType = MFConditionTypeEqual
	oOneSC.Expression.SetPropertyValueExpression 1457,  MFilesAPI.MFParentChildBehavior.MFParentChildBehaviorNone, Nothing
	oOneSC.TypedValue.SetValue MFilesAPI.MFDataType.MFDatatypeLookup,  RepresentativeID
	oSCs_representative.Add -1, oOneSC

	' Country = CountryID
	oOneSC.ConditionType = MFConditionTypeEqual
	oOneSC.Expression.SetPropertyValueExpression 1507,  MFilesAPI.MFParentChildBehavior.MFParentChildBehaviorNone, Nothing
	oOneSC.TypedValue.SetValue MFilesAPI.MFDataType.MFDatatypeLookup,  CountryID
	oSCs_country.Add -1, oOneSC	

	Dim oSearchResults_customer, oSearchResults_representative, oSearchResults_country
	Set oSearchResults_customer 		= Vault.ObjectSearchOperations.SearchForObjectsByConditions(oScs_customer, MFSearchFlagNone, false)
	Set oSearchResults_representative 	= Vault.ObjectSearchOperations.SearchForObjectsByConditions(oScs_representative, MFSearchFlagNone, false)
	Set oSearchResults_country 			= Vault.ObjectSearchOperations.SearchForObjectsByConditions(oScs_country, MFSearchFlagNone, false)
			
	Dim oLookups : Set oLookups = CreateObject("MFilesAPI.Lookups")
	Dim oLookup : Set oLookup = CreateObject("MFilesAPI.Lookup")

	dim sprawdzacz_profilu
	sprawdzacz_profilu = false
'	err.raise mfscriptcancel, "znaleziono plik " & oSearchResults.Count



	'TUTAJ WYNIKI WYSZUKIWANIA PROFILU KLIENTA
	If oSearchResults_customer.Count > 0 Then
		if sterownik = false Then
			
			WyszukajProfilKlienta = true

		elseif sterownik = true Then
		sprawdzacz_profilu = true
			For Each oSearchResult In oSearchResults_customer
				oLookup.Item = oSearchResult.ObjVer.ID
				oLookups.Add -1, oLookup
			Next		
	'	err.raise mfscriptcancel, "ile oSearchResults znaleziono? - " & oLookups.Count
		End if
	End if
	
	'TUTAJ WYNIKI WYSZUKIWANIA PROFILU REPA
	If oSearchResults_representative.Count > 0 Then
		if sterownik = false Then
			
			WyszukajProfilKlienta = true

		elseif sterownik = true Then
		sprawdzacz_profilu = true
			For Each oSearchResult In oSearchResults_representative
				oLookup.Item = oSearchResult.ObjVer.ID
				oLookups.Add -1, oLookup
			Next		
	'	err.raise mfscriptcancel, "ile oSearchResults znaleziono? - " & oLookups.Count
		End if
	End if
	
	'TUTAJ WYNIKI WYSZUKIWANIA PROFILU REPA
	If oSearchResults_country.Count > 0 Then
		if sterownik = false Then
			
			WyszukajProfilKlienta = true

		elseif sterownik = true Then
		sprawdzacz_profilu = true
			For Each oSearchResult In oSearchResults_country
				oLookup.Item = oSearchResult.ObjVer.ID
				oLookups.Add -1, oLookup
			Next		
	'	err.raise mfscriptcancel, "ile oSearchResults znaleziono? - " & oLookups.Count
		End if
	End if
	
	if sprawdzacz_profilu = true Then
		set WyszukajProfilKlienta = oLookups
	End if
	
End Function

	'--------------------------------------------------------------------
		'--------------------------------------------------------------------
			'--------------------------------------------------------------------
				'--------------------------------------------------------------------
					'--------------------------------------------------------------------
						'--------------------------------------------------------------------
							'--------------------------------------------------------------------