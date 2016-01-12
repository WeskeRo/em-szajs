%PROPERTY_23% wprowadził obiekt <a href=" %MFILESURL%"> <b>%PROPERTY_0% </b></a> w stan DO REALIZACJI. <br>
 Proszę o przystąpienie do realizacji.  <br> <br>. 


<hr>
<b>Dodatkowe informacje:</b><br>
Kategoria dokumentu: <b>%PROPERTY_100%</b> <br>
Typ produktu: <b>%PROPERTY_11609%</b> <br>
Metryka: <b> %PROPERTY_11612%</b> <br>
Link do dokumentu: <a href="%MFILESURL%"><b> LINK </b></a>



Dim dok_Commercial
		dok_Commercial = array("197", "200", "201", "202", "208", "220", "226", "227", "235", "241", "248", "250")
	
		If IsInArray(CStr(ID_Kategori.DisplayID), dok_Commercial) Then
			'Tutaj obsługujemy dokumenty podłaczane pod proces KOMERCYJNY
			L_ResponsiblePerson = 1455
			L_Proces = 11677
			dodajDoLog("Typ obiektu 156 = " & L_Proces)
		else 
			'Tutaj obsługujemy dokumenty podłaczane pod proces OFERTOWY
			L_ResponsiblePerson = 1455
			L_Proces = 11712
			dodajDoLog("Typ obiektu 159 = " & L_Proces)
		End if
		

		
		
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
