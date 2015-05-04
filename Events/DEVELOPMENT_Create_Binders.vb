Dim przelacznik
przelacznik = Cint(ObjVer.Type)
Select Case przelacznik
	Case 183 ' Nowy obiekt typu DEVELOPMENT
	
		NEW_Development()	

	Case 4 'Production

	Case 14 'Dokumenty QA

End Select



sub dodajDoLog(message)
	ForAppending = 8
	
	set objFSO = CreateObject("Scripting.FileSystemObject")
	set objFile = objFSO.OpenTextFile("E:\test\Developent.txt", ForAppending, True)
	
	objFile.WriteLine(message)
	objFile.Close
end sub

'-------------------------------------------------------
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
	Dim oOnePropVal_KartaRoz: Set oOnePropVal_KartaRoz = CreateObject("MFilesAPI.PropertyValue")
	Dim ObjVer_KartaRoz: Set ObjVer_KartaRoz = CreateObject("MFilesAPI.PropertyValue")
	
	' Class = Test Results
	oOnePropVal_KartaRoz.PropertyDef = 100
	oOnePropVal_KartaRoz.TypedValue.SetValue MFDatatypeLookup, idClass_KartaRozw
	oPropVals_KartaRoz.Add -1, oOnePropVal_KartaRoz
	
	' Development
	oPropVals_KartaRoz.Add -1, oOnePropVal_Input
	
	
	'--------------------------------------------------------------------
	
	Call Add_SFD_from_template (oPropVals_KartaRoz, 10852)
	
	' Files
	Dim oACL: Set oACL = CreateObject("MFilesAPI.AccessControlList")
	set oACL = nothing
	
	Dim oFiles: Set oFiles = CreateObject("MFilesAPI.SourceObjectFiles")
	
	Call Vault.ObjectOperations.CreateNewObjectEx(idType_InputDoc, oPropVals_Input, oFiles, False, True, oACL)
	Call Vault.ObjectOperations.CreateNewObjectEx(idType_InputDoc, oPropVals_Output, oFiles, False, True, oACL)
	Call Vault.ObjectOperations.CreateNewObjectEx(idType_InputDoc, oPropVals_Catalog, oFiles, False, True, oACL)
	Call Vault.ObjectOperations.CreateNewObjectEx(idType_InputDoc, oPropVals_Tests, oFiles, False, True, oACL)
End Sub

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
		oFiles1.Add 0, oSourceFile1

		Call Vault.ObjectOperations.CreateNewObjectEx(0, Bind_Properties, oFiles1, true, True, oACL)

		'Kasowanie pliku tymczasowego
		Dim fso: Set fso = CreateObject("Scripting.FileSystemObject")
		fso.DeleteFile szPath
	end if

End Sub

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
		'Dim NowyDokumentMFD
		'Set MFD = Vault.ObjectOperations.CreateNewObjectEx(9, Bind_Properties, oFiles1, False, True, oACL)'
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


Function WyszukajProfilKlienta(LeadID, ClientID)
	
	Const iOTDocument = 0 'Builtin, do not change
	Const iPDClass  = 100 'Builtin, do not change
	Const iClassManual  = 0 'ID kategorii dokumentu "0-Unclassified document"

	' Build the search conditions
	Dim oOneSC: Set oOneSC = CreateObject("MFilesAPI.SearchCondition")
	Dim oSCs: Set oSCs = CreateObject("MFilesAPI.SearchConditions")

	' Deleted = no
	oOneSC.ConditionType = MFConditionTypeEqual
	oOneSC.Expression.DataStatusValueType = MFStatusTypeDeleted
	oOneSC.TypedValue.SetValue MFDatatypeBoolean, False
	oSCs.Add -1, oOneSC

	' Object type = Document
	oOneSC.ConditionType = MFConditionTypeEqual
	oOneSC.Expression.DataStatusValueType = MFStatusTypeObjectTypeID
	oOneSC.TypedValue.SetValue MFDatatypeLookup, iOTDocument
	oScs.Add -1, oOneSC
	
	' Class = Manuals (general)
	oOneSC.ConditionType = MFConditionTypeEqual
	oOneSC.Expression.DataPropertyValuePropertyDef = iPDClass
	oOneSC.TypedValue.SetValue MFDatatypeLookup, iClassManual
	oScs.Add -1, oOneSC
	
	' Customer = ClientID
	oOneSC.ConditionType = MFConditionTypeEqual
	oOneSC.Expression.SetPropertyValueExpression 1457,  MFilesAPI.MFParentChildBehavior.MFParentChildBehaviorNone, Nothing
	oOneSC.TypedValue.SetValue MFDatatypeLookup, 1011
	oScs.Add -1, oOneSC

	
	
	Dim oSearchResults
	Set oSearchResults = Vault.ObjectSearchOperations.SearchForObjectsByConditions(oScs, MFSearchFlagNone, False)
	err.raise mfscriptcancel, "znaleziono plik " & oSearchResults.Count
	
End Function