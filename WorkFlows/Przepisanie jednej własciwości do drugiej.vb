'--------------------------------------------------------------------
'--------------------------------------------------------------------
Dim id_OjectType			 : id_ObjectType			= 227 		' ID Typu Obiektu - Document Binder 
Dim idProperty_Category		 : idProperty_Category		= 100		' ID metadanej Sales Lead 
Dim idProperty_SalesLead	 : idProperty_SalesLead		= 1484		' ID metadanej Customer 
Dim idProperty_RespPers_SL	 : idProperty_RespPers_SL	= 1455		' ID metadanej Sales Lead 
Dim idProperty_Manager		 : idProperty_Manager		= 1162		' ID metadanej Customer 
Dim idProperty_Supervisor_SL : idProperty_Supervisor_SL	= 1456		' ID metadanej Sales Lead 
'--------------------------------------------------------------------
Dim idValue_CommercialProcess : idValue_CommercialProcess = 203
'--------------------------------------------------------------------
		
	'kategoria_maila = Vault.ObjectPropertyOperations.GetProperty( objVer, 100).TypedValue.GetLookupID

		Dim staryObjVer
		Set staryObjVer = ObjVer.Clone()
		'dim ver
		'ver= staryObjVer.Version
	'	ver=ver-1 
		staryObjVer.Version = staryObjVer.Version - 1 
		Set myPropertyValues1 = Vault.ObjectPropertyOperations.GetProperties(staryObjVer)

		'id naszego salesleada oObjID.ID
		
		moja = myPropertyValues1.SearchForProperty(39).TypedValue.GetLookupID
		
		tytul = myPropertyValues1.SearchForProperty(0).TypedValue.GetValueAsLocalizedText




Dim oOnePropVal_PK: Set oOnePropVal_PK = CreateObject("MFilesAPI.PropertyValue")
'Dim looks: Set looks = CreateObject("MFilesAPI.PropertyValue")
		
'Klasa = Proces Ofertowy
oOnePropVal_PK.PropertyDef = 11671
oOnePropVal_PK.TypedValue.SetValue MFDatatypeText, tytul

Vault.ObjectPropertyOperations.SetProperty ObjVer, oOnePropVal_PK	


			oOnePropVal_PK.PropertyDef = 39
			oOnePropVal_PK.TypedValue.SetValue MFDatatypeLookup, moja




Vault.ObjectPropertyOperations.SetProperty ObjVer, oOnePropVal_PK	
