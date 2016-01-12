	dodajDoLog("---------------------------")
	dodajDoLog("CreateBinders - enter")


Dim oLookUpRef
Dim przelacznik, check_dep, check_subdep, VSV_name, project_number, output_value
dim id_dep, id_dep_sufix, id_subdep, id_subdep_sufix
dim dep_sufix, subdep_sufix
check_dep = false
check_subdep = false

przelacznik = Cint(ObjVer.Type)

Select Case przelacznik

	Case 184 'Typ obiektu = PROJEKT



	if IsNull(licznik) or IsEmpty(licznik) or licznik = "" Then ' wartość jest niezainicjowana lub zerowa
		VaultSharedVariables(nazwaZm) = 1
		licznik = 1
	end if
	dodajDoLog("test - " &licznik )




	if ( IsEmpty(Vault.ObjectPropertyOperations.GetProperty( ObjVer, 1212).Value)) then 'pole recznego wpisywania numeru
			Err.Raise mfscriptcancel, "Dział musi być wypełniony zawsze!"

	else
		
		check_dep = true
		
		id_dep = Vault.ObjectPropertyOperations.GetProperty( objVer, 1212).TypedValue.GetLookupID

		Select Case id_dep
			
			case 1
				subdep_sufix="P"
			case 2
				subdep_sufix="SS"
			case 3
				subdep_sufix="QC"
			case 4
				subdep_sufix="PP"
			case 5
				subdep_sufix="L"
			case 6
				subdep_sufix="S"
			case 8
				subdep_sufix="A"
			case 9
				subdep_sufix="HR"
			case 10
				subdep_sufix="HS"
			case 11
				subdep_sufix="S"
			case 12
				subdep_sufix="S"
			case 13
				subdep_sufix="BO"

		End Select

	





		if ( IsEmpty(Vault.ObjectPropertyOperations.GetProperty( ObjVer, 11802).Value)) then 'pole recznego wpisywania numeru
			Err.Raise mfscriptcancel, "Sub Dział musi być wypełniony zawsze!"

		else
			check_subdep = true
			id_subdep = Vault.ObjectPropertyOperations.GetProperty( objVer, 11802).TypedValue.GetLookupID

			Select Case id_subdep
			
			case 1
				dep_sufix="IT"
			case 2
				dep_sufix="SM"
			case 3
				dep_sufix="PD"
			case 4
				dep_sufix="RD"
			case 6
				dep_sufix="IT"
			case 7
				dep_sufix="QA"
			case 9
				dep_sufix="BR"

		End Select

		end if 'if od subdziału

	end if 'ifo od działu


	if check_dep= true And check_subdep = false
	
		VSV_name = Year(Now)& "/"& dep_sufix 
		project_number = VaultSharedVariables(VSV_name) 


		if IsNull(project_number) or IsEmpty(project_number) or project_number = "" Then ' wartość jest niezainicjowana lub zerowa
			VaultSharedVariables(VSV_name) = 1
			project_number = 1
		end if
		
		output_value = VSV_name & "/" & Right("000" & project_number,3)

		VaultSharedVariables(VSV_name) = project_number+1

		dodajDoLog("NUMER PROJEKTU DZIAŁU - " & output_value)


	else if (check_dep = true AND check_subdep = true) then
	
		VSV_name = Year(Now)& "/"& dep_sufix & "/"& subdep_sufix
		project_number = VaultSharedVariables(VSV_name) 


		if IsNull(project_number) or IsEmpty(project_number) or project_number = "" Then ' wartość jest niezainicjowana lub zerowa
			VaultSharedVariables(VSV_name) = 1
			project_number = 1
		end if
		
		output_value = VSV_name & "/" &  Right("000" & project_number,3)
		VaultSharedVariables(VSV_name) = project_number+1
		
		dodajDoLog("NUMER PROJEKTU SUBDZIAŁU - " & output_value)


	End If
		
End Select





	





































if( ObjVer.Type = 184) then

	Dim yy, xx, nn
	yy = "MG/" & Year(Date())
	xx = Year(Date())
	If isnull(LastUsed) or (LastUsed="") Then
		nn = 1
	ElseIf Left(LastUsed,7)=yy Then	
		nn = CInt(Right(LastUsed,3)) +1
	Else
		nn = 1
	End If 
			
	Output ="MG/" & xx & "/" & Right("000" & nn,3)

end if
