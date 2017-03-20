Option Explicit

Private Sub ComboBox1_Change()
    Select Case ComboBox1.Value
    Case "Gdynia": Obszar_pogoda = 1
    Case "Grudzi¹dz": Obszar_pogoda = 2
    Case "Warszawa": Obszar_pogoda = 3
    Case "Suwa³ki": Obszar_pogoda = 4
    Case "Bia³ystok": Obszar_pogoda = 5
    Case "Olsztyn": Obszar_pogoda = 6
    Case "Koszalin": Obszar_pogoda = 7
    Case "Szczecin": Obszar_pogoda = 8
    Case "Zielona Góra": Obszar_pogoda = 9
    Case "Poznañ": Obszar_pogoda = 10
    Case "Kalisz": Obszar_pogoda = 11
    Case "LódŸ": Obszar_pogoda = 12
    Case "Lublin": Obszar_pogoda = 13
    Case "Rzeszów": Obszar_pogoda = 14
    Case "Czêstochowa": Obszar_pogoda = 15
    Case "Katowice": Obszar_pogoda = 16
    Case "Gliwice": Obszar_pogoda = 17
    Case "Kraków": Obszar_pogoda = 18
    Case "Wroc³aw": Obszar_pogoda = 19
    Case "Bielsko-Bia³a": Obszar_pogoda = 20
    End Select
End Sub

Private Sub ComboBox2_Change()
 
 Select Case ComboBox2.Value
    Case "PSG Nienominowani": ID_Klient = 1
    Case "Jaro": ID_Klient = 1
    Case "Stalprodukt": ID_Klient = 1
    Case "PCO ¯arów": ID_Klient = 1
    Case "Szpital Grudzi¹dz": ID_Klient = 1
    Case "Gates": ID_Klient = 1
    Case "Zak³ady Miêsne Duda": ID_Klient = 1
    Case "Lajkonik": ID_Klient = 1
    Case "Grana": ID_Klient = 1
    Case "EBCC": ID_Klient = 1
    Case "Certech": ID_Klient = 1
    Case "Bia³a Góra": ID_Klient = 1
    Case "Lerg": ID_Klient = 1
    Case "Bobrek": ID_Klient = 1
    Case "Rafako": ID_Klient = 1
    Case "Pini Polonia": ID_Klient = 1
    Case "Veolia": ID_Klient = 1
    Case "Animex": ID_Klient = 1
    Case "General Motors": ID_Klient = 1
    Case "Philip Morris": ID_Klient = 1
    Case "Kopalnia Czatkowice": ID_Klient = 1
    Case "Effect System": ID_Klient = 1
    Case "Pini Polska": ID_Klient = 1
    Case "Tauron Ciep³o": ID_Klient = 1
    Case "HCM": ID_Klient = 1
    Case "Beskidy": ID_Klient = 1
    End Select
End Sub

Private Sub CommandButton1_Click()
    'pobieranie danych pogodowych
    Dim strChaine As Date
    strChaine = VBA.Format(Workbooks("prognoza Gaz PM.xlsm").Sheets("Analiza - 1").Range("F4"), "yyyy-mm-dd")
    Dim day_od As String
    Dim month_od As String
    Dim Year_od As String
    Dim strSQL As String
    Dim cn As ADODB.Connection
    Set cn = New ADODB.Connection
    Set rst = New ADODB.Recordset
    Dim objaccess As Object
    Set objaccess = CreateObject("Access.Application")
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    cn.Open "Provider=Microsoft.ACE.OLEDB.15.0; Data Source=" & baza_gaz
    Const acImport = 0
            
   'data_od
  Year_od = VBA.Format(Workbooks("prognoza Gaz PM.xlsm").Sheets("Analiza - 1").Range("F4"), "yyyy")
  day_od = VBA.Format(Workbooks("prognoza Gaz PM.xlsm").Sheets("Analiza - 1").Range("F4"), "dd")
  month_od = VBA.Format(Workbooks("prognoza Gaz PM.xlsm").Sheets("Analiza - 1").Range("F4"), "mm")

  Workbooks("prognoza Gaz PM.xlsm").Sheets("Analiza - 1").Range("C12:E35").ClearContents
  
  'ustawienie daty prognozy w zak³adce Wykresy - 2
  Workbooks("prognoza Gaz PM.xlsm").Sheets("Wykresy - 2").Cells(53, 2) = day_od
  Workbooks("prognoza Gaz PM.xlsm").Sheets("Wykresy - 2").Cells(112, 2) = day_od
  
  'Prognoza pogody
  strSQL = "SELECT Prognoza!Temperatura, Prognoza!Naslon FROM Prognoza WHERE (Prognoza!Godzina > 6 AND Prognoza!Obszar_ID =" & Obszar_pogoda _
  & " and Prognoza.data_prognozy = '" & Year_od & month_od & day_od & "') and Prognoza.Doba = #" & strChaine + 1 & "# ;"
  rst.Open strSQL, cn
  Workbooks("prognoza Gaz PM.xlsm").Sheets("Analiza - 1").Cells(12, 3).CopyFromRecordset rst
  strSQL = "SELECT Prognoza!Temperatura, Prognoza!Naslon FROM Prognoza WHERE (Prognoza!Godzina < 7 AND Prognoza!Obszar_ID =" & Obszar_pogoda _
  & " and Prognoza.data_prognozy = '" & Year_od & month_od & day_od & "') and Prognoza.Doba = #" & strChaine + 2 & "# ;"
  Set rst = New ADODB.Recordset
  rst.Open strSQL, cn
  Workbooks("prognoza Gaz PM.xlsm").Sheets("Analiza - 1").Cells(29, 3).CopyFromRecordset rst
  Application.ScreenUpdating = True
  Application.DisplayAlerts = True
  Set rst = Nothing
  cn.Close
End Sub

Private Sub CommandButton2_Click()
    'wyszukiwanie dni podobnych pogodowo
    UserForm2.Show
End Sub

Private Sub CommandButton3_Click()
  
    Arkusz2.Range("a46").Value = Sheets("Analiza - 1").Range("B37")
    Dim i As Integer
    For i = 0 To 26
     Arkusz2.Cells(46, 3 + i).Value = Arkusz1.Cells(41 + i, 6)
     Next i
     Application.ScreenUpdating = True
    
End Sub

Private Sub CommandButton4_Click()
    Arkusz2.Range("a47").Value = Sheets("Analiza - 1").Range("G37")
    Dim i As Integer
    For i = 0 To 26
     Arkusz2.Cells(47, 3 + i).Value = Arkusz1.Cells(41 + i, 11)
     Next i
     Application.ScreenUpdating = True
End Sub

Private Sub CommandButton5_Click()
    Arkusz2.Range("a48").Value = Sheets("Analiza - 1").Range("L37")
    Dim i As Integer
    For i = 0 To 26
    Arkusz2.Cells(48, 3 + i).Value = Arkusz1.Cells(41 + i, 16)
    Next i
    Application.ScreenUpdating = True
End Sub

Private Sub CommandButton6_Click()
    Arkusz2.Range("a49").Value = Sheets("Analiza - 1").Range("Q37")
      Dim i As Integer
    For i = 0 To 26
     Arkusz2.Cells(49, 3 + i).Value = Arkusz1.Cells(41 + i, 21)
     Next i
     Application.ScreenUpdating = True
     
End Sub

Private Sub czy_wybrany_klient_progoda()
	If Obszar_pogoda = "" Then
	  MsgBox "Nie wybrano obszaru pogody!"
	  End
	  ElseIf ID_Klient = "" Then
	  MsgBox "Nie wybrano klienta!"
	  End
	  End If
End Sub

Private Sub SpinButton2_Change()
  Call czy_wybrany_klient_progoda
  Dim strSQL As String
  Set cn = New ADODB.Connection
  Set rst = New ADODB.Recordset
  Set objaccess = CreateObject("Access.Application")
  Dim ldate As Date
  Dim start_row
  Application.ScreenUpdating = False
  Application.DisplayAlerts = False
  cn.Open "Provider=Microsoft.ACE.OLEDB.15.0; Data Source=" & baza_gaz
  ldate = VBA.Format(Workbooks("prognoza Gaz PM.xlsm").Sheets("Analiza - 1").Range("B37"), "yyyy-mm-dd")
  Workbooks("prognoza Gaz PM.xlsm").Sheets("Analiza - 1").Range("C41:F67").ClearContents
'Wykonanie pogody
  
  strSQL = "SELECT Wykonanie!Temperatura, Wykonanie!Naslon FROM Wykonanie WHERE (Wykonanie!Godzina > 3 AND Wykonanie!Obszar_ID =" & Obszar_pogoda _
  & " and Wykonanie.Doba = #" & ldate & "#)  ;"
  rst.Open strSQL, cn
  Workbooks("prognoza Gaz PM.xlsm").Sheets("Analiza - 1").Cells(41, 3).CopyFromRecordset rst
  strSQL = "SELECT Wykonanie!Temperatura, Wykonanie!Naslon FROM Wykonanie WHERE (Wykonanie!Godzina < 7 AND Wykonanie!Obszar_ID =" & Obszar_pogoda _
  & " and Wykonanie.Doba = #" & ldate + 1 & "#);"
  Set rst = New ADODB.Recordset
  rst.Open strSQL, cn
  Workbooks("prognoza Gaz PM.xlsm").Sheets("Analiza - 1").Cells(61, 3).CopyFromRecordset rst
  
   'Wykonanie zu¿ycia kWh
 
'Wykonanie zuzycia
  Set rst = New ADODB.Recordset
  strSQL = "SELECT Wolumen_wykon!Doba, Wolumen_wykon!H7, Wolumen_wykon!H8, Wolumen_wykon!H9, Wolumen_wykon!H10, Wolumen_wykon!H11, Wolumen_wykon!H12, Wolumen_wykon!H13, Wolumen_wykon!H14, Wolumen_wykon!H15, Wolumen_wykon!H16, Wolumen_wykon!H17," _
  & "Wolumen_wykon!H18, Wolumen_wykon!H19, Wolumen_wykon!H20, Wolumen_wykon!H21, Wolumen_wykon!H22, Wolumen_wykon!H23, Wolumen_wykon!H0, Wolumen_wykon!H1, Wolumen_wykon!H2, Wolumen_wykon!H3, Wolumen_wykon!H4, Wolumen_wykon!H5," _
  & "Wolumen_wykon!H6 FROM Wolumen_wykon WHERE (Wolumen_wykon!ID_Klient =" & ID_Klient & " and (Wolumen_wykon.Doba BETWEEN #" & ldate - 1 & "# and #" & ldate & "#));"
  rst.Open strSQL, cn
  start_row = 44
  Dim objFields As ADODB.Fields
  Set objFields = rst.Fields
  
  Do While (rst.EOF = False)
  If objFields.Item("Wolumen_wykon!Doba").Value = ldate - 1 Then
    Workbooks("prognoza Gaz PM.xlsm").Sheets("Analiza - 1").Cells(start_row - 3, 6) = objFields.Item("Wolumen_wykon!H4").Value / 1000
    Workbooks("prognoza Gaz PM.xlsm").Sheets("Analiza - 1").Cells(start_row - 2, 6) = objFields.Item("Wolumen_wykon!H5").Value / 1000
    Workbooks("prognoza Gaz PM.xlsm").Sheets("Analiza - 1").Cells(start_row - 1, 6) = objFields.Item("Wolumen_wykon!H6").Value / 1000
  Else
    Workbooks("prognoza Gaz PM.xlsm").Sheets("Analiza - 1").Cells(start_row - 0, 6) = objFields.Item("Wolumen_wykon!H7").Value / 1000
    Workbooks("prognoza Gaz PM.xlsm").Sheets("Analiza - 1").Cells(start_row + 1, 6) = objFields.Item("Wolumen_wykon!H8").Value / 1000
    Workbooks("prognoza Gaz PM.xlsm").Sheets("Analiza - 1").Cells(start_row + 2, 6) = objFields.Item("Wolumen_wykon!H9").Value / 1000
    Workbooks("prognoza Gaz PM.xlsm").Sheets("Analiza - 1").Cells(start_row + 3, 6) = objFields.Item("Wolumen_wykon!H10").Value / 1000
    Workbooks("prognoza Gaz PM.xlsm").Sheets("Analiza - 1").Cells(start_row + 4, 6) = objFields.Item("Wolumen_wykon!H11").Value / 1000
    Workbooks("prognoza Gaz PM.xlsm").Sheets("Analiza - 1").Cells(start_row + 5, 6) = objFields.Item("Wolumen_wykon!H12").Value / 1000
    Workbooks("prognoza Gaz PM.xlsm").Sheets("Analiza - 1").Cells(start_row + 6, 6) = objFields.Item("Wolumen_wykon!H13").Value / 1000
    Workbooks("prognoza Gaz PM.xlsm").Sheets("Analiza - 1").Cells(start_row + 7, 6) = objFields.Item("Wolumen_wykon!H14").Value / 1000
    Workbooks("prognoza Gaz PM.xlsm").Sheets("Analiza - 1").Cells(start_row + 8, 6) = objFields.Item("Wolumen_wykon!H15").Value / 1000
    Workbooks("prognoza Gaz PM.xlsm").Sheets("Analiza - 1").Cells(start_row + 9, 6) = objFields.Item("Wolumen_wykon!H16").Value / 1000
    Workbooks("prognoza Gaz PM.xlsm").Sheets("Analiza - 1").Cells(start_row + 10, 6) = objFields.Item("Wolumen_wykon!H17").Value / 1000
    Workbooks("prognoza Gaz PM.xlsm").Sheets("Analiza - 1").Cells(start_row + 11, 6) = objFields.Item("Wolumen_wykon!H18").Value / 1000
    Workbooks("prognoza Gaz PM.xlsm").Sheets("Analiza - 1").Cells(start_row + 12, 6) = objFields.Item("Wolumen_wykon!H19").Value / 1000
    Workbooks("prognoza Gaz PM.xlsm").Sheets("Analiza - 1").Cells(start_row + 13, 6) = objFields.Item("Wolumen_wykon!H20").Value / 1000
    Workbooks("prognoza Gaz PM.xlsm").Sheets("Analiza - 1").Cells(start_row + 14, 6) = objFields.Item("Wolumen_wykon!H21").Value / 1000
    Workbooks("prognoza Gaz PM.xlsm").Sheets("Analiza - 1").Cells(start_row + 15, 6) = objFields.Item("Wolumen_wykon!H22").Value / 1000
    Workbooks("prognoza Gaz PM.xlsm").Sheets("Analiza - 1").Cells(start_row + 16, 6) = objFields.Item("Wolumen_wykon!H23").Value / 1000
    Workbooks("prognoza Gaz PM.xlsm").Sheets("Analiza - 1").Cells(start_row + 17, 6) = objFields.Item("Wolumen_wykon!H0").Value / 1000
    Workbooks("prognoza Gaz PM.xlsm").Sheets("Analiza - 1").Cells(start_row + 18, 6) = objFields.Item("Wolumen_wykon!H1").Value / 1000
    Workbooks("prognoza Gaz PM.xlsm").Sheets("Analiza - 1").Cells(start_row + 19, 6) = objFields.Item("Wolumen_wykon!H2").Value / 1000
    Workbooks("prognoza Gaz PM.xlsm").Sheets("Analiza - 1").Cells(start_row + 20, 6) = objFields.Item("Wolumen_wykon!H3").Value / 1000
    Workbooks("prognoza Gaz PM.xlsm").Sheets("Analiza - 1").Cells(start_row + 21, 6) = objFields.Item("Wolumen_wykon!H4").Value / 1000
    Workbooks("prognoza Gaz PM.xlsm").Sheets("Analiza - 1").Cells(start_row + 22, 6) = objFields.Item("Wolumen_wykon!H5").Value / 1000
    Workbooks("prognoza Gaz PM.xlsm").Sheets("Analiza - 1").Cells(start_row + 23, 6) = objFields.Item("Wolumen_wykon!H6").Value / 1000
    End If
    rst.MoveNext
  Loop
  Application.ScreenUpdating = True
  Application.DisplayAlerts = True
  Set rst = Nothing
  Set objFields = Nothing
  cn.Close
  Application.ScreenUpdating = True
  Workbooks("prognoza Gaz PM.xlsm").Sheets("Analiza - 1").Visible = True
End Sub

Private Sub SpinButton3_Change()
	Call czy_wybrany_klient_progoda
  Dim strSQL As String
  Dim cn As ADODB.Connection
  Set cn = New ADODB.Connection
  Set rst = New ADODB.Recordset
  Dim objaccess As Object
  Set objaccess = CreateObject("Access.Application")
  Dim ldate As Date
  Dim start_row
  Application.ScreenUpdating = False
  Application.DisplayAlerts = False
  cn.Open "Provider=Microsoft.ACE.OLEDB.15.0; Data Source=" & baza_gaz
  Const acImport = 0
  ldate = VBA.Format(Workbooks("prognoza Gaz PM.xlsm").Sheets("Analiza - 1").Range("G37"), "yyyy-mm-dd")
  Workbooks("prognoza Gaz PM.xlsm").Sheets("Analiza - 1").Range("H41:K67").ClearContents
'Wykonanie pogody
  strSQL = "SELECT Wykonanie!Temperatura, Wykonanie!Naslon FROM Wykonanie WHERE (Wykonanie!Godzina > 3 AND Wykonanie!Obszar_ID =" & Obszar_pogoda _
  & " and Wykonanie.Doba = #" & ldate & "#)  ;"
  rst.Open strSQL, cn
  Workbooks("prognoza Gaz PM.xlsm").Sheets("Analiza - 1").Cells(41, 8).CopyFromRecordset rst
  strSQL = "SELECT Wykonanie!Temperatura, Wykonanie!Naslon FROM Wykonanie WHERE (Wykonanie!Godzina < 7 AND Wykonanie!Obszar_ID =" & Obszar_pogoda _
  & " and Wykonanie.Doba = #" & ldate + 1 & "#);"
  Set rst = New ADODB.Recordset
  rst.Open strSQL, cn
  Workbooks("prognoza Gaz PM.xlsm").Sheets("Analiza - 1").Cells(61, 8).CopyFromRecordset rst
  
   'Wykonanie zu¿ycia kWh
 
'Wykonanie zuzycia
  Set rst = New ADODB.Recordset
  strSQL = "SELECT Wolumen_wykon!Doba, Wolumen_wykon!H7, Wolumen_wykon!H8, Wolumen_wykon!H9, Wolumen_wykon!H10, Wolumen_wykon!H11, Wolumen_wykon!H12, Wolumen_wykon!H13, Wolumen_wykon!H14, Wolumen_wykon!H15, Wolumen_wykon!H16, Wolumen_wykon!H17," _
  & "Wolumen_wykon!H18, Wolumen_wykon!H19, Wolumen_wykon!H20, Wolumen_wykon!H21, Wolumen_wykon!H22, Wolumen_wykon!H23, Wolumen_wykon!H0, Wolumen_wykon!H1, Wolumen_wykon!H2, Wolumen_wykon!H3, Wolumen_wykon!H4, Wolumen_wykon!H5," _
  & "Wolumen_wykon!H6 FROM Wolumen_wykon WHERE (Wolumen_wykon!ID_Klient =" & ID_Klient & " and (Wolumen_wykon.Doba BETWEEN #" & ldate - 1 & "# and #" & ldate & "#));"
  rst.Open strSQL, cn
  start_row = 44
  Dim objFields As ADODB.Fields
  Set objFields = rst.Fields
  
  Do While (rst.EOF = False)
  If objFields.Item("Wolumen_wykon!Doba").Value = ldate - 1 Then
    Workbooks("prognoza Gaz PM.xlsm").Sheets("Analiza - 1").Cells(start_row - 3, 11) = objFields.Item("Wolumen_wykon!H4").Value / 1000
    Workbooks("prognoza Gaz PM.xlsm").Sheets("Analiza - 1").Cells(start_row - 2, 11) = objFields.Item("Wolumen_wykon!H5").Value / 1000
    Workbooks("prognoza Gaz PM.xlsm").Sheets("Analiza - 1").Cells(start_row - 1, 11) = objFields.Item("Wolumen_wykon!H6").Value / 1000
  Else
    Workbooks("prognoza Gaz PM.xlsm").Sheets("Analiza - 1").Cells(start_row - 0, 11) = objFields.Item("Wolumen_wykon!H7").Value / 1000
    Workbooks("prognoza Gaz PM.xlsm").Sheets("Analiza - 1").Cells(start_row + 1, 11) = objFields.Item("Wolumen_wykon!H8").Value / 1000
    Workbooks("prognoza Gaz PM.xlsm").Sheets("Analiza - 1").Cells(start_row + 2, 11) = objFields.Item("Wolumen_wykon!H9").Value / 1000
    Workbooks("prognoza Gaz PM.xlsm").Sheets("Analiza - 1").Cells(start_row + 3, 11) = objFields.Item("Wolumen_wykon!H10").Value / 1000
    Workbooks("prognoza Gaz PM.xlsm").Sheets("Analiza - 1").Cells(start_row + 4, 11) = objFields.Item("Wolumen_wykon!H11").Value / 1000
    Workbooks("prognoza Gaz PM.xlsm").Sheets("Analiza - 1").Cells(start_row + 5, 11) = objFields.Item("Wolumen_wykon!H12").Value / 1000
    Workbooks("prognoza Gaz PM.xlsm").Sheets("Analiza - 1").Cells(start_row + 6, 11) = objFields.Item("Wolumen_wykon!H13").Value / 1000
    Workbooks("prognoza Gaz PM.xlsm").Sheets("Analiza - 1").Cells(start_row + 7, 11) = objFields.Item("Wolumen_wykon!H14").Value / 1000
    Workbooks("prognoza Gaz PM.xlsm").Sheets("Analiza - 1").Cells(start_row + 8, 11) = objFields.Item("Wolumen_wykon!H15").Value / 1000
    Workbooks("prognoza Gaz PM.xlsm").Sheets("Analiza - 1").Cells(start_row + 9, 11) = objFields.Item("Wolumen_wykon!H16").Value / 1000
    Workbooks("prognoza Gaz PM.xlsm").Sheets("Analiza - 1").Cells(start_row + 10, 11) = objFields.Item("Wolumen_wykon!H17").Value / 1000
    Workbooks("prognoza Gaz PM.xlsm").Sheets("Analiza - 1").Cells(start_row + 11, 11) = objFields.Item("Wolumen_wykon!H18").Value / 1000
    Workbooks("prognoza Gaz PM.xlsm").Sheets("Analiza - 1").Cells(start_row + 12, 11) = objFields.Item("Wolumen_wykon!H19").Value / 1000
    Workbooks("prognoza Gaz PM.xlsm").Sheets("Analiza - 1").Cells(start_row + 13, 11) = objFields.Item("Wolumen_wykon!H20").Value / 1000
    Workbooks("prognoza Gaz PM.xlsm").Sheets("Analiza - 1").Cells(start_row + 14, 11) = objFields.Item("Wolumen_wykon!H21").Value / 1000
    Workbooks("prognoza Gaz PM.xlsm").Sheets("Analiza - 1").Cells(start_row + 15, 11) = objFields.Item("Wolumen_wykon!H22").Value / 1000
    Workbooks("prognoza Gaz PM.xlsm").Sheets("Analiza - 1").Cells(start_row + 16, 11) = objFields.Item("Wolumen_wykon!H23").Value / 1000
    Workbooks("prognoza Gaz PM.xlsm").Sheets("Analiza - 1").Cells(start_row + 17, 11) = objFields.Item("Wolumen_wykon!H0").Value / 1000
    Workbooks("prognoza Gaz PM.xlsm").Sheets("Analiza - 1").Cells(start_row + 18, 11) = objFields.Item("Wolumen_wykon!H1").Value / 1000
    Workbooks("prognoza Gaz PM.xlsm").Sheets("Analiza - 1").Cells(start_row + 19, 11) = objFields.Item("Wolumen_wykon!H2").Value / 1000
    Workbooks("prognoza Gaz PM.xlsm").Sheets("Analiza - 1").Cells(start_row + 20, 11) = objFields.Item("Wolumen_wykon!H3").Value / 1000
    Workbooks("prognoza Gaz PM.xlsm").Sheets("Analiza - 1").Cells(start_row + 21, 11) = objFields.Item("Wolumen_wykon!H4").Value / 1000
    Workbooks("prognoza Gaz PM.xlsm").Sheets("Analiza - 1").Cells(start_row + 22, 11) = objFields.Item("Wolumen_wykon!H5").Value / 1000
    Workbooks("prognoza Gaz PM.xlsm").Sheets("Analiza - 1").Cells(start_row + 23, 11) = objFields.Item("Wolumen_wykon!H6").Value / 1000
    End If
    rst.MoveNext
  Loop
  Application.ScreenUpdating = True
  Application.DisplayAlerts = True
  Set rst = Nothing
  Set objFields = Nothing
  cn.Close
  
  Application.ScreenUpdating = True
  Workbooks("prognoza Gaz PM.xlsm").Sheets("Analiza - 1").Visible = True
End Sub

Private Sub SpinButton4_Change()
	Call czy_wybrany_klient_progoda
  Dim strSQL As String
  Dim cn As ADODB.Connection
  Set cn = New ADODB.Connection
  Set rst = New ADODB.Recordset
  Dim objaccess As Object
  Set objaccess = CreateObject("Access.Application")
  Dim ldate As Date
  Dim start_row
  Application.ScreenUpdating = False
  Application.DisplayAlerts = False
  cn.Open "Provider=Microsoft.ACE.OLEDB.15.0; Data Source=" & baza_gaz
  Const acImport = 0
  ldate = VBA.Format(Workbooks("prognoza Gaz PM.xlsm").Sheets("Analiza - 1").Range("L37"), "yyyy-mm-dd")
  Workbooks("prognoza Gaz PM.xlsm").Sheets("Analiza - 1").Range("M41:P67").ClearContents
'Wykonanie pogody
  strSQL = "SELECT Wykonanie!Temperatura, Wykonanie!Naslon FROM Wykonanie WHERE (Wykonanie!Godzina > 3 AND Wykonanie!Obszar_ID =" & Obszar_pogoda _
  & " and Wykonanie.Doba = #" & ldate & "#)  ;"
  rst.Open strSQL, cn
  Workbooks("prognoza Gaz PM.xlsm").Sheets("Analiza - 1").Cells(41, 13).CopyFromRecordset rst
  strSQL = "SELECT Wykonanie!Temperatura, Wykonanie!Naslon FROM Wykonanie WHERE (Wykonanie!Godzina < 7 AND Wykonanie!Obszar_ID =" & Obszar_pogoda _
  & " and Wykonanie.Doba = #" & ldate + 1 & "#);"
  Set rst = New ADODB.Recordset
  rst.Open strSQL, cn
  Workbooks("prognoza Gaz PM.xlsm").Sheets("Analiza - 1").Cells(61, 13).CopyFromRecordset rst
  
   'Wykonanie zu¿ycia kWh
 
'Wykonanie zuzycia
  Set rst = New ADODB.Recordset
  strSQL = "SELECT Wolumen_wykon!Doba, Wolumen_wykon!H7, Wolumen_wykon!H8, Wolumen_wykon!H9, Wolumen_wykon!H10, Wolumen_wykon!H11, Wolumen_wykon!H12, Wolumen_wykon!H13, Wolumen_wykon!H14, Wolumen_wykon!H15, Wolumen_wykon!H16, Wolumen_wykon!H17," _
  & "Wolumen_wykon!H18, Wolumen_wykon!H19, Wolumen_wykon!H20, Wolumen_wykon!H21, Wolumen_wykon!H22, Wolumen_wykon!H23, Wolumen_wykon!H0, Wolumen_wykon!H1, Wolumen_wykon!H2, Wolumen_wykon!H3, Wolumen_wykon!H4, Wolumen_wykon!H5," _
  & "Wolumen_wykon!H6 FROM Wolumen_wykon WHERE (Wolumen_wykon!ID_Klient =" & ID_Klient & " and (Wolumen_wykon.Doba BETWEEN #" & ldate - 1 & "# and #" & ldate & "#));"
  rst.Open strSQL, cn
  start_row = 44
  Dim objFields As ADODB.Fields
  Set objFields = rst.Fields
  
  Do While (rst.EOF = False)
  If objFields.Item("Wolumen_wykon!Doba").Value = ldate - 1 Then
    Workbooks("prognoza Gaz PM.xlsm").Sheets("Analiza - 1").Cells(start_row - 3, 16) = objFields.Item("Wolumen_wykon!H4").Value / 1000
    Workbooks("prognoza Gaz PM.xlsm").Sheets("Analiza - 1").Cells(start_row - 2, 16) = objFields.Item("Wolumen_wykon!H5").Value / 1000
    Workbooks("prognoza Gaz PM.xlsm").Sheets("Analiza - 1").Cells(start_row - 1, 16) = objFields.Item("Wolumen_wykon!H6").Value / 1000
  Else
    Workbooks("prognoza Gaz PM.xlsm").Sheets("Analiza - 1").Cells(start_row - 0, 16) = objFields.Item("Wolumen_wykon!H7").Value / 1000
    Workbooks("prognoza Gaz PM.xlsm").Sheets("Analiza - 1").Cells(start_row + 1, 16) = objFields.Item("Wolumen_wykon!H8").Value / 1000
    Workbooks("prognoza Gaz PM.xlsm").Sheets("Analiza - 1").Cells(start_row + 2, 16) = objFields.Item("Wolumen_wykon!H9").Value / 1000
    Workbooks("prognoza Gaz PM.xlsm").Sheets("Analiza - 1").Cells(start_row + 3, 16) = objFields.Item("Wolumen_wykon!H10").Value / 1000
    Workbooks("prognoza Gaz PM.xlsm").Sheets("Analiza - 1").Cells(start_row + 4, 16) = objFields.Item("Wolumen_wykon!H11").Value / 1000
    Workbooks("prognoza Gaz PM.xlsm").Sheets("Analiza - 1").Cells(start_row + 5, 16) = objFields.Item("Wolumen_wykon!H12").Value / 1000
    Workbooks("prognoza Gaz PM.xlsm").Sheets("Analiza - 1").Cells(start_row + 6, 16) = objFields.Item("Wolumen_wykon!H13").Value / 1000
    Workbooks("prognoza Gaz PM.xlsm").Sheets("Analiza - 1").Cells(start_row + 7, 16) = objFields.Item("Wolumen_wykon!H14").Value / 1000
    Workbooks("prognoza Gaz PM.xlsm").Sheets("Analiza - 1").Cells(start_row + 8, 16) = objFields.Item("Wolumen_wykon!H15").Value / 1000
    Workbooks("prognoza Gaz PM.xlsm").Sheets("Analiza - 1").Cells(start_row + 9, 16) = objFields.Item("Wolumen_wykon!H16").Value / 1000
    Workbooks("prognoza Gaz PM.xlsm").Sheets("Analiza - 1").Cells(start_row + 10, 16) = objFields.Item("Wolumen_wykon!H17").Value / 1000
    Workbooks("prognoza Gaz PM.xlsm").Sheets("Analiza - 1").Cells(start_row + 11, 16) = objFields.Item("Wolumen_wykon!H18").Value / 1000
    Workbooks("prognoza Gaz PM.xlsm").Sheets("Analiza - 1").Cells(start_row + 12, 16) = objFields.Item("Wolumen_wykon!H19").Value / 1000
    Workbooks("prognoza Gaz PM.xlsm").Sheets("Analiza - 1").Cells(start_row + 13, 16) = objFields.Item("Wolumen_wykon!H20").Value / 1000
    Workbooks("prognoza Gaz PM.xlsm").Sheets("Analiza - 1").Cells(start_row + 14, 16) = objFields.Item("Wolumen_wykon!H21").Value / 1000
    Workbooks("prognoza Gaz PM.xlsm").Sheets("Analiza - 1").Cells(start_row + 15, 16) = objFields.Item("Wolumen_wykon!H22").Value / 1000
    Workbooks("prognoza Gaz PM.xlsm").Sheets("Analiza - 1").Cells(start_row + 16, 16) = objFields.Item("Wolumen_wykon!H23").Value / 1000
    Workbooks("prognoza Gaz PM.xlsm").Sheets("Analiza - 1").Cells(start_row + 17, 16) = objFields.Item("Wolumen_wykon!H0").Value / 1000
    Workbooks("prognoza Gaz PM.xlsm").Sheets("Analiza - 1").Cells(start_row + 18, 16) = objFields.Item("Wolumen_wykon!H1").Value / 1000
    Workbooks("prognoza Gaz PM.xlsm").Sheets("Analiza - 1").Cells(start_row + 19, 16) = objFields.Item("Wolumen_wykon!H2").Value / 1000
    Workbooks("prognoza Gaz PM.xlsm").Sheets("Analiza - 1").Cells(start_row + 20, 16) = objFields.Item("Wolumen_wykon!H3").Value / 1000
    Workbooks("prognoza Gaz PM.xlsm").Sheets("Analiza - 1").Cells(start_row + 21, 16) = objFields.Item("Wolumen_wykon!H4").Value / 1000
    Workbooks("prognoza Gaz PM.xlsm").Sheets("Analiza - 1").Cells(start_row + 22, 16) = objFields.Item("Wolumen_wykon!H5").Value / 1000
    Workbooks("prognoza Gaz PM.xlsm").Sheets("Analiza - 1").Cells(start_row + 23, 16) = objFields.Item("Wolumen_wykon!H6").Value / 1000
    End If
    rst.MoveNext
  Loop
  Application.ScreenUpdating = True
  Application.DisplayAlerts = True
  Set rst = Nothing
  Set objFields = Nothing
  cn.Close
  
  Application.ScreenUpdating = True
  Workbooks("prognoza Gaz PM.xlsm").Sheets("Analiza - 1").Visible = True
End Sub

Private Sub SpinButton5_Change()
	Call czy_wybrany_klient_progoda
  Dim strSQL As String
  Dim cn As ADODB.Connection
  Set cn = New ADODB.Connection
  Set rst = New ADODB.Recordset
  Dim objaccess As Object
  Set objaccess = CreateObject("Access.Application")
  Dim ldate As Date
  Dim start_row
  Application.ScreenUpdating = False
  Application.DisplayAlerts = False
  cn.Open "Provider=Microsoft.ACE.OLEDB.15.0; Data Source=" & baza_gaz
  Const acImport = 0
  ldate = VBA.Format(Workbooks("prognoza Gaz PM.xlsm").Sheets("Analiza - 1").Range("Q37"), "yyyy-mm-dd")
  Workbooks("prognoza Gaz PM.xlsm").Sheets("Analiza - 1").Range("R41:U67").ClearContents
'Wykonanie pogody
  strSQL = "SELECT Wykonanie!Temperatura, Wykonanie!Naslon FROM Wykonanie WHERE (Wykonanie!Godzina > 3 AND Wykonanie!Obszar_ID =" & Obszar_pogoda _
  & " and Wykonanie.Doba = #" & ldate & "#)  ;"
  rst.Open strSQL, cn
  Workbooks("prognoza Gaz PM.xlsm").Sheets("Analiza - 1").Cells(41, 18).CopyFromRecordset rst
  strSQL = "SELECT Wykonanie!Temperatura, Wykonanie!Naslon FROM Wykonanie WHERE (Wykonanie!Godzina < 7 AND Wykonanie!Obszar_ID =" & Obszar_pogoda _
  & " and Wykonanie.Doba = #" & ldate + 1 & "#);"
  Set rst = New ADODB.Recordset
  rst.Open strSQL, cn
  Workbooks("prognoza Gaz PM.xlsm").Sheets("Analiza - 1").Cells(61, 18).CopyFromRecordset rst
  
   'Wykonanie zu¿ycia kWh
 
'Wykonanie zuzycia
  Set rst = New ADODB.Recordset
  strSQL = "SELECT Wolumen_wykon!Doba, Wolumen_wykon!H7, Wolumen_wykon!H8, Wolumen_wykon!H9, Wolumen_wykon!H10, Wolumen_wykon!H11, Wolumen_wykon!H12, Wolumen_wykon!H13, Wolumen_wykon!H14, Wolumen_wykon!H15, Wolumen_wykon!H16, Wolumen_wykon!H17," _
  & "Wolumen_wykon!H18, Wolumen_wykon!H19, Wolumen_wykon!H20, Wolumen_wykon!H21, Wolumen_wykon!H22, Wolumen_wykon!H23, Wolumen_wykon!H0, Wolumen_wykon!H1, Wolumen_wykon!H2, Wolumen_wykon!H3, Wolumen_wykon!H4, Wolumen_wykon!H5," _
  & "Wolumen_wykon!H6 FROM Wolumen_wykon WHERE (Wolumen_wykon!ID_Klient =" & ID_Klient & " and (Wolumen_wykon.Doba BETWEEN #" & ldate - 1 & "# and #" & ldate & "#));"
  rst.Open strSQL, cn
  start_row = 44
  Dim objFields As ADODB.Fields
  Set objFields = rst.Fields
  
  Do While (rst.EOF = False)
  If objFields.Item("Wolumen_wykon!Doba").Value = ldate - 1 Then
    Workbooks("prognoza Gaz PM.xlsm").Sheets("Analiza - 1").Cells(start_row - 3, 21) = objFields.Item("Wolumen_wykon!H4").Value / 1000
    Workbooks("prognoza Gaz PM.xlsm").Sheets("Analiza - 1").Cells(start_row - 2, 21) = objFields.Item("Wolumen_wykon!H5").Value / 1000
    Workbooks("prognoza Gaz PM.xlsm").Sheets("Analiza - 1").Cells(start_row - 1, 21) = objFields.Item("Wolumen_wykon!H6").Value / 1000
  Else
    Workbooks("prognoza Gaz PM.xlsm").Sheets("Analiza - 1").Cells(start_row - 0, 21) = objFields.Item("Wolumen_wykon!H7").Value / 1000
    Workbooks("prognoza Gaz PM.xlsm").Sheets("Analiza - 1").Cells(start_row + 1, 21) = objFields.Item("Wolumen_wykon!H8").Value / 1000
    Workbooks("prognoza Gaz PM.xlsm").Sheets("Analiza - 1").Cells(start_row + 2, 21) = objFields.Item("Wolumen_wykon!H9").Value / 1000
    Workbooks("prognoza Gaz PM.xlsm").Sheets("Analiza - 1").Cells(start_row + 3, 21) = objFields.Item("Wolumen_wykon!H10").Value / 1000
    Workbooks("prognoza Gaz PM.xlsm").Sheets("Analiza - 1").Cells(start_row + 4, 21) = objFields.Item("Wolumen_wykon!H11").Value / 1000
    Workbooks("prognoza Gaz PM.xlsm").Sheets("Analiza - 1").Cells(start_row + 5, 21) = objFields.Item("Wolumen_wykon!H12").Value / 1000
    Workbooks("prognoza Gaz PM.xlsm").Sheets("Analiza - 1").Cells(start_row + 6, 21) = objFields.Item("Wolumen_wykon!H13").Value / 1000
    Workbooks("prognoza Gaz PM.xlsm").Sheets("Analiza - 1").Cells(start_row + 7, 21) = objFields.Item("Wolumen_wykon!H14").Value / 1000
    Workbooks("prognoza Gaz PM.xlsm").Sheets("Analiza - 1").Cells(start_row + 8, 21) = objFields.Item("Wolumen_wykon!H15").Value / 1000
    Workbooks("prognoza Gaz PM.xlsm").Sheets("Analiza - 1").Cells(start_row + 9, 21) = objFields.Item("Wolumen_wykon!H16").Value / 1000
    Workbooks("prognoza Gaz PM.xlsm").Sheets("Analiza - 1").Cells(start_row + 10, 21) = objFields.Item("Wolumen_wykon!H17").Value / 1000
    Workbooks("prognoza Gaz PM.xlsm").Sheets("Analiza - 1").Cells(start_row + 11, 21) = objFields.Item("Wolumen_wykon!H18").Value / 1000
    Workbooks("prognoza Gaz PM.xlsm").Sheets("Analiza - 1").Cells(start_row + 12, 21) = objFields.Item("Wolumen_wykon!H19").Value / 1000
    Workbooks("prognoza Gaz PM.xlsm").Sheets("Analiza - 1").Cells(start_row + 13, 21) = objFields.Item("Wolumen_wykon!H20").Value / 1000
    Workbooks("prognoza Gaz PM.xlsm").Sheets("Analiza - 1").Cells(start_row + 14, 21) = objFields.Item("Wolumen_wykon!H21").Value / 1000
    Workbooks("prognoza Gaz PM.xlsm").Sheets("Analiza - 1").Cells(start_row + 15, 21) = objFields.Item("Wolumen_wykon!H22").Value / 1000
    Workbooks("prognoza Gaz PM.xlsm").Sheets("Analiza - 1").Cells(start_row + 16, 21) = objFields.Item("Wolumen_wykon!H23").Value / 1000
    Workbooks("prognoza Gaz PM.xlsm").Sheets("Analiza - 1").Cells(start_row + 17, 21) = objFields.Item("Wolumen_wykon!H0").Value / 1000
    Workbooks("prognoza Gaz PM.xlsm").Sheets("Analiza - 1").Cells(start_row + 18, 21) = objFields.Item("Wolumen_wykon!H1").Value / 1000
    Workbooks("prognoza Gaz PM.xlsm").Sheets("Analiza - 1").Cells(start_row + 19, 21) = objFields.Item("Wolumen_wykon!H2").Value / 1000
    Workbooks("prognoza Gaz PM.xlsm").Sheets("Analiza - 1").Cells(start_row + 20, 21) = objFields.Item("Wolumen_wykon!H3").Value / 1000
    Workbooks("prognoza Gaz PM.xlsm").Sheets("Analiza - 1").Cells(start_row + 21, 21) = objFields.Item("Wolumen_wykon!H4").Value / 1000
    Workbooks("prognoza Gaz PM.xlsm").Sheets("Analiza - 1").Cells(start_row + 22, 21) = objFields.Item("Wolumen_wykon!H5").Value / 1000
    Workbooks("prognoza Gaz PM.xlsm").Sheets("Analiza - 1").Cells(start_row + 23, 21) = objFields.Item("Wolumen_wykon!H6").Value / 1000
    End If
    rst.MoveNext
  Loop
  Application.ScreenUpdating = True
  Application.DisplayAlerts = True
  Set rst = Nothing
  Set objFields = Nothing
  cn.Close
  
  Application.ScreenUpdating = True
  Workbooks("prognoza Gaz PM.xlsm").Sheets("Analiza - 1").Visible = True
End Sub

Private Sub SpinButton1_SpinUp()
    Dim dtDate  As Date
    dtDate = Format(Workbooks("prognoza Gaz PM.xlsm").Sheets("Analiza - 1").Range("f4").Value, "yyyy-mm-dd")
    dtDate = dtDate + 1
    Workbooks("prognoza Gaz PM.xlsm").Sheets("Analiza - 1").Range("f4").Value = Format(dtDate, "yyyy-mm-dd")
    Workbooks("prognoza Gaz PM.xlsm").Sheets("Wykresy - 2").Range("b53").Value = Format(dtDate, "dd")
End Sub

Private Sub SpinButton1_SpinDown()
    Dim dtDate  As Date
    dtDate = Format(Workbooks("prognoza Gaz PM.xlsm").Sheets("Analiza - 1").Range("f4").Value, "yyyy-mm-dd")
    dtDate = dtDate - 1
    Workbooks("prognoza Gaz PM.xlsm").Sheets("Analiza - 1").Range("f4").Value = Format(dtDate, "yyyy-mm-dd")
    Workbooks("prognoza Gaz PM.xlsm").Sheets("Wykresy - 2").Range("b53").Value = Format(dtDate, "dd")
End Sub

Private Sub Worksheet_SelectionChange(ByVal Target As Range)

End Sub
