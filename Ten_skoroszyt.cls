Option Explicit

Private Sub Workbook_Open()
  Dim ldate As String
  ldate = Date + 1
  
  Select Case Application.UserName 
	Case "Paweł Machał": userID = 1
	Case "Witold Stawiarski": userID = 2
	Case Else userID = 99
	End Select
  Application.ScreenUpdating = False
  Workbooks("prognoza Gaz PM.xlsm").Sheets("Analiza - 1").Range("F4") = VBA.Format(ldate, "yyyy/mm/dd")
        
	With Workbooks("prognoza Gaz PM.xlsm").Sheets("Analiza - 1").ComboBox1
		.Text = ""
		.AddItem "Gdynia"
		.AddItem "Grudziądz"
		.AddItem "Warszawa"
		.AddItem "Suwałki"
		.AddItem "Białystok"
		.AddItem "Olsztyn"
		.AddItem "Koszalin"
		.AddItem "Szczecin"
		.AddItem "Zielona Góra"
		.AddItem "Poznań"
		.AddItem "Kalisz"
		.AddItem "Lódź"
		.AddItem "Lublin"
		.AddItem "Rzeszów"
		.AddItem "Częstochowa"
		.AddItem "Katowice"
		.AddItem "Gliwice"
		.AddItem "Kraków"
		.AddItem "Wrocław"
		.AddItem "Bielsko-Biała"
		End With
	With Workbooks("prognoza Gaz PM.xlsm").Sheets("Analiza - 1").ComboBox2
		.Text = ""
		.AddItem "PSG Nienominowani"
		.AddItem "Jaro"
		.AddItem "Stalprodukt"
		.AddItem "PCO Żarów"
		.AddItem "Szpital Grudziądz"
		.AddItem "Gates"
		.AddItem "Zakłady Mięsne Duda"
		.AddItem "Lajkonik"
		.AddItem "Grana"
		.AddItem "EBCC"
		.AddItem "Certech"
		.AddItem "Biała Góra"
		.AddItem "Lerg"
		.AddItem "Bobrek"
		.AddItem "Rafako"
		.AddItem "Pini Polonia"
		.AddItem "Veolia"
		.AddItem "Animex"
		.AddItem "General Motors"
		.AddItem "Philip Morris"
		.AddItem "Kopalnia Czatkowice"
		.AddItem "Effect System"
		.AddItem "Pini Polska"
		.AddItem "Tauron Ciepło"
		.AddItem "HCM"
		.AddItem "Beskidy"
		.AddItem "Boryszew"
		End With
  Application.ScreenUpdating = True
End Sub