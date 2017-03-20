Option Explicit

Private Sub SpinButton1_Change()
  Dim strSQL As String
  Set cn = New ADODB.Connection
  Set rst = New ADODB.Recordset
  Set objaccess = CreateObject("Access.Application")
  Dim ldate As Date
  Dim start_row
  Application.ScreenUpdating = False
  Application.DisplayAlerts = False
  cn.Open "Provider=Microsoft.ACE.OLEDB.15.0; Data Source=" & baza_gaz
  ldate = VBA.Format(Workbooks("prognoza Gaz PM.xlsm").Sheets("Delta 2 dni - 9").Range("B30"), "yyyy-mm-dd")
  Workbooks("prognoza Gaz PM.xlsm").Sheets("Delta 2 dni - 9").Range("C33:F56").ClearContents
'Wykonanie pogody
  strSQL = "SELECT Wykonanie!Temperatura, Wykonanie!Naslon FROM Wykonanie WHERE (Wykonanie!Godzina > 6 AND Wykonanie!Obszar_ID =" & Obszar_pogoda _
  & " and Wykonanie.Doba = #" & ldate & "#)  ;"
  rst.Open strSQL, cn
  Workbooks("prognoza Gaz PM.xlsm").Sheets("Delta 2 dni - 9").Cells(33, 3).CopyFromRecordset rst
  strSQL = "SELECT Wykonanie!Temperatura, Wykonanie!Naslon FROM Wykonanie WHERE (Wykonanie!Godzina < 7 AND Wykonanie!Obszar_ID =" & Obszar_pogoda _
  & " and Wykonanie.Doba = #" & ldate + 1 & "#);"
  Set rst = New ADODB.Recordset
  rst.Open strSQL, cn
  Workbooks("prognoza Gaz PM.xlsm").Sheets("Delta 2 dni - 9").Cells(50, 3).CopyFromRecordset rst
  
   'Wykonanie zużycia kWh
 
'Wykonanie zuzycia
  Set rst = New ADODB.Recordset
  strSQL = "SELECT Wolumen_wykon!Doba, Wolumen_wykon!H7, Wolumen_wykon!H8, Wolumen_wykon!H9, Wolumen_wykon!H10, Wolumen_wykon!H11, Wolumen_wykon!H12, Wolumen_wykon!H13, Wolumen_wykon!H14, Wolumen_wykon!H15, Wolumen_wykon!H16, Wolumen_wykon!H17," _
  & "Wolumen_wykon!H18, Wolumen_wykon!H19, Wolumen_wykon!H20, Wolumen_wykon!H21, Wolumen_wykon!H22, Wolumen_wykon!H23, Wolumen_wykon!H0, Wolumen_wykon!H1, Wolumen_wykon!H2, Wolumen_wykon!H3, Wolumen_wykon!H4, Wolumen_wykon!H5," _
  & "Wolumen_wykon!H6 FROM Wolumen_wykon WHERE (Wolumen_wykon!ID_Klient =" & ID_Klient & " and (Wolumen_wykon.Doba = #" & ldate & "#));"
  rst.Open strSQL, cn
  start_row = 33
  Dim objFields As ADODB.Fields
  Set objFields = rst.Fields
  
  Do While (rst.EOF = False)
    Workbooks("prognoza Gaz PM.xlsm").Sheets("Delta 2 dni - 9").Cells(start_row - 0, 6) = objFields.Item("Wolumen_wykon!H7").Value / 1000
    Workbooks("prognoza Gaz PM.xlsm").Sheets("Delta 2 dni - 9").Cells(start_row + 1, 6) = objFields.Item("Wolumen_wykon!H8").Value / 1000
    Workbooks("prognoza Gaz PM.xlsm").Sheets("Delta 2 dni - 9").Cells(start_row + 2, 6) = objFields.Item("Wolumen_wykon!H9").Value / 1000
    Workbooks("prognoza Gaz PM.xlsm").Sheets("Delta 2 dni - 9").Cells(start_row + 3, 6) = objFields.Item("Wolumen_wykon!H10").Value / 1000
    Workbooks("prognoza Gaz PM.xlsm").Sheets("Delta 2 dni - 9").Cells(start_row + 4, 6) = objFields.Item("Wolumen_wykon!H11").Value / 1000
    Workbooks("prognoza Gaz PM.xlsm").Sheets("Delta 2 dni - 9").Cells(start_row + 5, 6) = objFields.Item("Wolumen_wykon!H12").Value / 1000
    Workbooks("prognoza Gaz PM.xlsm").Sheets("Delta 2 dni - 9").Cells(start_row + 6, 6) = objFields.Item("Wolumen_wykon!H13").Value / 1000
    Workbooks("prognoza Gaz PM.xlsm").Sheets("Delta 2 dni - 9").Cells(start_row + 7, 6) = objFields.Item("Wolumen_wykon!H14").Value / 1000
    Workbooks("prognoza Gaz PM.xlsm").Sheets("Delta 2 dni - 9").Cells(start_row + 8, 6) = objFields.Item("Wolumen_wykon!H15").Value / 1000
    Workbooks("prognoza Gaz PM.xlsm").Sheets("Delta 2 dni - 9").Cells(start_row + 9, 6) = objFields.Item("Wolumen_wykon!H16").Value / 1000
    Workbooks("prognoza Gaz PM.xlsm").Sheets("Delta 2 dni - 9").Cells(start_row + 10, 6) = objFields.Item("Wolumen_wykon!H17").Value / 1000
    Workbooks("prognoza Gaz PM.xlsm").Sheets("Delta 2 dni - 9").Cells(start_row + 11, 6) = objFields.Item("Wolumen_wykon!H18").Value / 1000
    Workbooks("prognoza Gaz PM.xlsm").Sheets("Delta 2 dni - 9").Cells(start_row + 12, 6) = objFields.Item("Wolumen_wykon!H19").Value / 1000
    Workbooks("prognoza Gaz PM.xlsm").Sheets("Delta 2 dni - 9").Cells(start_row + 13, 6) = objFields.Item("Wolumen_wykon!H20").Value / 1000
    Workbooks("prognoza Gaz PM.xlsm").Sheets("Delta 2 dni - 9").Cells(start_row + 14, 6) = objFields.Item("Wolumen_wykon!H21").Value / 1000
    Workbooks("prognoza Gaz PM.xlsm").Sheets("Delta 2 dni - 9").Cells(start_row + 15, 6) = objFields.Item("Wolumen_wykon!H22").Value / 1000
    Workbooks("prognoza Gaz PM.xlsm").Sheets("Delta 2 dni - 9").Cells(start_row + 16, 6) = objFields.Item("Wolumen_wykon!H23").Value / 1000
    Workbooks("prognoza Gaz PM.xlsm").Sheets("Delta 2 dni - 9").Cells(start_row + 17, 6) = objFields.Item("Wolumen_wykon!H0").Value / 1000
    Workbooks("prognoza Gaz PM.xlsm").Sheets("Delta 2 dni - 9").Cells(start_row + 18, 6) = objFields.Item("Wolumen_wykon!H1").Value / 1000
    Workbooks("prognoza Gaz PM.xlsm").Sheets("Delta 2 dni - 9").Cells(start_row + 19, 6) = objFields.Item("Wolumen_wykon!H2").Value / 1000
    Workbooks("prognoza Gaz PM.xlsm").Sheets("Delta 2 dni - 9").Cells(start_row + 20, 6) = objFields.Item("Wolumen_wykon!H3").Value / 1000
    Workbooks("prognoza Gaz PM.xlsm").Sheets("Delta 2 dni - 9").Cells(start_row + 21, 6) = objFields.Item("Wolumen_wykon!H4").Value / 1000
    Workbooks("prognoza Gaz PM.xlsm").Sheets("Delta 2 dni - 9").Cells(start_row + 22, 6) = objFields.Item("Wolumen_wykon!H5").Value / 1000
    Workbooks("prognoza Gaz PM.xlsm").Sheets("Delta 2 dni - 9").Cells(start_row + 23, 6) = objFields.Item("Wolumen_wykon!H6").Value / 1000
    rst.MoveNext
  Loop
  Application.ScreenUpdating = True
  Application.DisplayAlerts = True
  Set rst = Nothing
  Set objFields = Nothing
  cn.Close
  
  'deklaracja zmiennych do skalowania
       Dim zakres As Range
       Dim max, min As Integer
       
    'scalowanie wykresu
    'ustawia zakres
  Set zakres = Workbooks("prognoza Gaz PM.xlsm").Sheets("Delta 2 dni - 9").Range("F33:F56")
        min = Application.WorksheetFunction.min(zakres)
        max = Application.WorksheetFunction.max(zakres)
     ActiveSheet.ChartObjects("Wykres 23").Activate
     ActiveChart.Axes(xlValue).Select
        
       With ActiveChart.Axes(xlValue)
       .MinimumScale = Round((min - 20), 0)
       .MaximumScale = Round((max + 20), 0)
       End With
    
  Application.ScreenUpdating = True
  Workbooks("prognoza Gaz PM.xlsm").Sheets("Delta 2 dni - 9").Visible = True
End Sub

Private Sub SpinButton2_Change()
  Dim strSQL As String
  Set cn = New ADODB.Connection
  Set rst = New ADODB.Recordset
  Set objaccess = CreateObject("Access.Application")
  Dim ldate As Date
  Dim start_row As Byte
  start_row = 33
  Application.ScreenUpdating = False
  Application.DisplayAlerts = False
  cn.Open "Provider=Microsoft.ACE.OLEDB.15.0; Data Source=" & baza_gaz
  ldate = VBA.Format(Workbooks("prognoza Gaz PM.xlsm").Sheets("Delta 2 dni - 9").Range("H30"), "yyyy-mm-dd")
  Workbooks("prognoza Gaz PM.xlsm").Sheets("Delta 2 dni - 9").Range("I33:L56").ClearContents
'Wykonanie pogody
  strSQL = "SELECT Wykonanie!Temperatura, Wykonanie!Naslon FROM Wykonanie WHERE (Wykonanie!Godzina > 6 AND Wykonanie!Obszar_ID =" & Obszar_pogoda _
  & " and Wykonanie.Doba = #" & ldate & "#)  ;"
  rst.Open strSQL, cn
  Workbooks("prognoza Gaz PM.xlsm").Sheets("Delta 2 dni - 9").Cells(start_row, 9).CopyFromRecordset rst
  strSQL = "SELECT Wykonanie!Temperatura, Wykonanie!Naslon FROM Wykonanie WHERE (Wykonanie!Godzina < 7 AND Wykonanie!Obszar_ID =" & Obszar_pogoda _
  & " and Wykonanie.Doba = #" & ldate + 1 & "#);"
  Set rst = New ADODB.Recordset
  rst.Open strSQL, cn
  Workbooks("prognoza Gaz PM.xlsm").Sheets("Delta 2 dni - 9").Cells(start_row + 17, 9).CopyFromRecordset rst
  
   'Wykonanie zużycia kWh
 
'Wykonanie zuzycia
  Set rst = New ADODB.Recordset
  strSQL = "SELECT Wolumen_wykon!Doba, Wolumen_wykon!H7, Wolumen_wykon!H8, Wolumen_wykon!H9, Wolumen_wykon!H10, Wolumen_wykon!H11, Wolumen_wykon!H12, Wolumen_wykon!H13, Wolumen_wykon!H14, Wolumen_wykon!H15, Wolumen_wykon!H16, Wolumen_wykon!H17," _
  & "Wolumen_wykon!H18, Wolumen_wykon!H19, Wolumen_wykon!H20, Wolumen_wykon!H21, Wolumen_wykon!H22, Wolumen_wykon!H23, Wolumen_wykon!H0, Wolumen_wykon!H1, Wolumen_wykon!H2, Wolumen_wykon!H3, Wolumen_wykon!H4, Wolumen_wykon!H5," _
  & "Wolumen_wykon!H6 FROM Wolumen_wykon WHERE (Wolumen_wykon!ID_Klient =" & ID_Klient & " and (Wolumen_wykon.Doba = #" & ldate & "#));"
  rst.Open strSQL, cn
  Dim objFields As ADODB.Fields
  Set objFields = rst.Fields
  
  Do While (rst.EOF = False)
    Workbooks("prognoza Gaz PM.xlsm").Sheets("Delta 2 dni - 9").Cells(start_row - 0, 12) = objFields.Item("Wolumen_wykon!H7").Value / 1000
    Workbooks("prognoza Gaz PM.xlsm").Sheets("Delta 2 dni - 9").Cells(start_row + 1, 12) = objFields.Item("Wolumen_wykon!H8").Value / 1000
    Workbooks("prognoza Gaz PM.xlsm").Sheets("Delta 2 dni - 9").Cells(start_row + 2, 12) = objFields.Item("Wolumen_wykon!H9").Value / 1000
    Workbooks("prognoza Gaz PM.xlsm").Sheets("Delta 2 dni - 9").Cells(start_row + 3, 12) = objFields.Item("Wolumen_wykon!H10").Value / 1000
    Workbooks("prognoza Gaz PM.xlsm").Sheets("Delta 2 dni - 9").Cells(start_row + 4, 12) = objFields.Item("Wolumen_wykon!H11").Value / 1000
    Workbooks("prognoza Gaz PM.xlsm").Sheets("Delta 2 dni - 9").Cells(start_row + 5, 12) = objFields.Item("Wolumen_wykon!H12").Value / 1000
    Workbooks("prognoza Gaz PM.xlsm").Sheets("Delta 2 dni - 9").Cells(start_row + 6, 12) = objFields.Item("Wolumen_wykon!H13").Value / 1000
    Workbooks("prognoza Gaz PM.xlsm").Sheets("Delta 2 dni - 9").Cells(start_row + 7, 12) = objFields.Item("Wolumen_wykon!H14").Value / 1000
    Workbooks("prognoza Gaz PM.xlsm").Sheets("Delta 2 dni - 9").Cells(start_row + 8, 12) = objFields.Item("Wolumen_wykon!H15").Value / 1000
    Workbooks("prognoza Gaz PM.xlsm").Sheets("Delta 2 dni - 9").Cells(start_row + 9, 12) = objFields.Item("Wolumen_wykon!H16").Value / 1000
    Workbooks("prognoza Gaz PM.xlsm").Sheets("Delta 2 dni - 9").Cells(start_row + 10, 12) = objFields.Item("Wolumen_wykon!H17").Value / 1000
    Workbooks("prognoza Gaz PM.xlsm").Sheets("Delta 2 dni - 9").Cells(start_row + 11, 12) = objFields.Item("Wolumen_wykon!H18").Value / 1000
    Workbooks("prognoza Gaz PM.xlsm").Sheets("Delta 2 dni - 9").Cells(start_row + 12, 12) = objFields.Item("Wolumen_wykon!H19").Value / 1000
    Workbooks("prognoza Gaz PM.xlsm").Sheets("Delta 2 dni - 9").Cells(start_row + 13, 12) = objFields.Item("Wolumen_wykon!H20").Value / 1000
    Workbooks("prognoza Gaz PM.xlsm").Sheets("Delta 2 dni - 9").Cells(start_row + 14, 12) = objFields.Item("Wolumen_wykon!H21").Value / 1000
    Workbooks("prognoza Gaz PM.xlsm").Sheets("Delta 2 dni - 9").Cells(start_row + 15, 12) = objFields.Item("Wolumen_wykon!H22").Value / 1000
    Workbooks("prognoza Gaz PM.xlsm").Sheets("Delta 2 dni - 9").Cells(start_row + 16, 12) = objFields.Item("Wolumen_wykon!H23").Value / 1000
    Workbooks("prognoza Gaz PM.xlsm").Sheets("Delta 2 dni - 9").Cells(start_row + 17, 12) = objFields.Item("Wolumen_wykon!H0").Value / 1000
    Workbooks("prognoza Gaz PM.xlsm").Sheets("Delta 2 dni - 9").Cells(start_row + 18, 12) = objFields.Item("Wolumen_wykon!H1").Value / 1000
    Workbooks("prognoza Gaz PM.xlsm").Sheets("Delta 2 dni - 9").Cells(start_row + 19, 12) = objFields.Item("Wolumen_wykon!H2").Value / 1000
    Workbooks("prognoza Gaz PM.xlsm").Sheets("Delta 2 dni - 9").Cells(start_row + 20, 12) = objFields.Item("Wolumen_wykon!H3").Value / 1000
    Workbooks("prognoza Gaz PM.xlsm").Sheets("Delta 2 dni - 9").Cells(start_row + 21, 12) = objFields.Item("Wolumen_wykon!H4").Value / 1000
    Workbooks("prognoza Gaz PM.xlsm").Sheets("Delta 2 dni - 9").Cells(start_row + 22, 12) = objFields.Item("Wolumen_wykon!H5").Value / 1000
    Workbooks("prognoza Gaz PM.xlsm").Sheets("Delta 2 dni - 9").Cells(start_row + 23, 12) = objFields.Item("Wolumen_wykon!H6").Value / 1000
    rst.MoveNext
  Loop
  Application.ScreenUpdating = True
  Application.DisplayAlerts = True
  Set rst = Nothing
  Set objFields = Nothing
  cn.Close
  
    'deklaracja zmiennych do skalowania
       Dim zakres As Range
       Dim max, min As Integer
       
    'scalowanie wykresu
    'ustawia zakres
  Set zakres = Workbooks("prognoza Gaz PM.xlsm").Sheets("Delta 2 dni - 9").Range("L33:L56")
        min = Application.WorksheetFunction.min(zakres)
        max = Application.WorksheetFunction.max(zakres)
     ActiveSheet.ChartObjects("Wykres 23").Activate
     ActiveChart.Axes(xlValue).Select
        
       With ActiveChart.Axes(xlValue)
       .MinimumScale = Round((min - 20), 0)
       .MaximumScale = Round((max + 20), 0)
       End With
       
  Application.ScreenUpdating = True
  Workbooks("prognoza Gaz PM.xlsm").Sheets("Delta 2 dni - 9").Visible = True
End Sub

Private Sub SpinButton3_Change()
  Dim strSQL As String
  Set cn = New ADODB.Connection
  Set rst = New ADODB.Recordset
  Set objaccess = CreateObject("Access.Application")
  Dim ldate As Date
  Dim start_row As Byte
  start_row = 33
  Application.ScreenUpdating = False
  Application.DisplayAlerts = False
  cn.Open "Provider=Microsoft.ACE.OLEDB.15.0; Data Source=" & baza_gaz
  ldate = VBA.Format(Workbooks("prognoza Gaz PM.xlsm").Sheets("Delta 2 dni - 9").Range("N30"), "yyyy-mm-dd")
  Workbooks("prognoza Gaz PM.xlsm").Sheets("Delta 2 dni - 9").Range("o33:r56").ClearContents
'Wykonanie pogody
  strSQL = "SELECT Wykonanie!Temperatura, Wykonanie!Naslon FROM Wykonanie WHERE (Wykonanie!Godzina > 6 AND Wykonanie!Obszar_ID =" & Obszar_pogoda _
  & " and Wykonanie.Doba = #" & ldate & "#)  ;"
  rst.Open strSQL, cn
  Workbooks("prognoza Gaz PM.xlsm").Sheets("Delta 2 dni - 9").Cells(start_row, 15).CopyFromRecordset rst
  strSQL = "SELECT Wykonanie!Temperatura, Wykonanie!Naslon FROM Wykonanie WHERE (Wykonanie!Godzina < 7 AND Wykonanie!Obszar_ID =" & Obszar_pogoda _
  & " and Wykonanie.Doba = #" & ldate + 1 & "#);"
  Set rst = New ADODB.Recordset
  rst.Open strSQL, cn
  Workbooks("prognoza Gaz PM.xlsm").Sheets("Delta 2 dni - 9").Cells(start_row + 17, 15).CopyFromRecordset rst
  
   'Wykonanie zużycia kWh
 
'Wykonanie zuzycia
  Set rst = New ADODB.Recordset
  strSQL = "SELECT Wolumen_wykon!Doba, Wolumen_wykon!H7, Wolumen_wykon!H8, Wolumen_wykon!H9, Wolumen_wykon!H10, Wolumen_wykon!H11, Wolumen_wykon!H12, Wolumen_wykon!H13, Wolumen_wykon!H14, Wolumen_wykon!H15, Wolumen_wykon!H16, Wolumen_wykon!H17," _
  & "Wolumen_wykon!H18, Wolumen_wykon!H19, Wolumen_wykon!H20, Wolumen_wykon!H21, Wolumen_wykon!H22, Wolumen_wykon!H23, Wolumen_wykon!H0, Wolumen_wykon!H1, Wolumen_wykon!H2, Wolumen_wykon!H3, Wolumen_wykon!H4, Wolumen_wykon!H5," _
  & "Wolumen_wykon!H6 FROM Wolumen_wykon WHERE (Wolumen_wykon!ID_Klient =" & ID_Klient & " and (Wolumen_wykon.Doba = #" & ldate & "#));"
  rst.Open strSQL, cn
  Dim objFields As ADODB.Fields
  Set objFields = rst.Fields
  
  Do While (rst.EOF = False)
    Workbooks("prognoza Gaz PM.xlsm").Sheets("Delta 2 dni - 9").Cells(start_row - 0, 18) = objFields.Item("Wolumen_wykon!H7").Value / 1000
    Workbooks("prognoza Gaz PM.xlsm").Sheets("Delta 2 dni - 9").Cells(start_row + 1, 18) = objFields.Item("Wolumen_wykon!H8").Value / 1000
    Workbooks("prognoza Gaz PM.xlsm").Sheets("Delta 2 dni - 9").Cells(start_row + 2, 18) = objFields.Item("Wolumen_wykon!H9").Value / 1000
    Workbooks("prognoza Gaz PM.xlsm").Sheets("Delta 2 dni - 9").Cells(start_row + 3, 18) = objFields.Item("Wolumen_wykon!H10").Value / 1000
    Workbooks("prognoza Gaz PM.xlsm").Sheets("Delta 2 dni - 9").Cells(start_row + 4, 18) = objFields.Item("Wolumen_wykon!H11").Value / 1000
    Workbooks("prognoza Gaz PM.xlsm").Sheets("Delta 2 dni - 9").Cells(start_row + 5, 18) = objFields.Item("Wolumen_wykon!H12").Value / 1000
    Workbooks("prognoza Gaz PM.xlsm").Sheets("Delta 2 dni - 9").Cells(start_row + 6, 18) = objFields.Item("Wolumen_wykon!H13").Value / 1000
    Workbooks("prognoza Gaz PM.xlsm").Sheets("Delta 2 dni - 9").Cells(start_row + 7, 18) = objFields.Item("Wolumen_wykon!H14").Value / 1000
    Workbooks("prognoza Gaz PM.xlsm").Sheets("Delta 2 dni - 9").Cells(start_row + 8, 18) = objFields.Item("Wolumen_wykon!H15").Value / 1000
    Workbooks("prognoza Gaz PM.xlsm").Sheets("Delta 2 dni - 9").Cells(start_row + 9, 18) = objFields.Item("Wolumen_wykon!H16").Value / 1000
    Workbooks("prognoza Gaz PM.xlsm").Sheets("Delta 2 dni - 9").Cells(start_row + 10, 18) = objFields.Item("Wolumen_wykon!H17").Value / 1000
    Workbooks("prognoza Gaz PM.xlsm").Sheets("Delta 2 dni - 9").Cells(start_row + 11, 18) = objFields.Item("Wolumen_wykon!H18").Value / 1000
    Workbooks("prognoza Gaz PM.xlsm").Sheets("Delta 2 dni - 9").Cells(start_row + 12, 18) = objFields.Item("Wolumen_wykon!H19").Value / 1000
    Workbooks("prognoza Gaz PM.xlsm").Sheets("Delta 2 dni - 9").Cells(start_row + 13, 18) = objFields.Item("Wolumen_wykon!H20").Value / 1000
    Workbooks("prognoza Gaz PM.xlsm").Sheets("Delta 2 dni - 9").Cells(start_row + 14, 18) = objFields.Item("Wolumen_wykon!H21").Value / 1000
    Workbooks("prognoza Gaz PM.xlsm").Sheets("Delta 2 dni - 9").Cells(start_row + 15, 18) = objFields.Item("Wolumen_wykon!H22").Value / 1000
    Workbooks("prognoza Gaz PM.xlsm").Sheets("Delta 2 dni - 9").Cells(start_row + 16, 18) = objFields.Item("Wolumen_wykon!H23").Value / 1000
    Workbooks("prognoza Gaz PM.xlsm").Sheets("Delta 2 dni - 9").Cells(start_row + 17, 18) = objFields.Item("Wolumen_wykon!H0").Value / 1000
    Workbooks("prognoza Gaz PM.xlsm").Sheets("Delta 2 dni - 9").Cells(start_row + 18, 18) = objFields.Item("Wolumen_wykon!H1").Value / 1000
    Workbooks("prognoza Gaz PM.xlsm").Sheets("Delta 2 dni - 9").Cells(start_row + 19, 18) = objFields.Item("Wolumen_wykon!H2").Value / 1000
    Workbooks("prognoza Gaz PM.xlsm").Sheets("Delta 2 dni - 9").Cells(start_row + 20, 18) = objFields.Item("Wolumen_wykon!H3").Value / 1000
    Workbooks("prognoza Gaz PM.xlsm").Sheets("Delta 2 dni - 9").Cells(start_row + 21, 18) = objFields.Item("Wolumen_wykon!H4").Value / 1000
    Workbooks("prognoza Gaz PM.xlsm").Sheets("Delta 2 dni - 9").Cells(start_row + 22, 18) = objFields.Item("Wolumen_wykon!H5").Value / 1000
    Workbooks("prognoza Gaz PM.xlsm").Sheets("Delta 2 dni - 9").Cells(start_row + 23, 18) = objFields.Item("Wolumen_wykon!H6").Value / 1000
    rst.MoveNext
  Loop
  Application.ScreenUpdating = True
  Application.DisplayAlerts = True
  Set rst = Nothing
  Set objFields = Nothing
  cn.Close
  
    'deklaracja zmiennych do skalowania
       Dim zakres As Range
       Dim max, min As Integer
       
    'scalowanie wykresu
    'ustawia zakres
  Set zakres = Workbooks("prognoza Gaz PM.xlsm").Sheets("Delta 2 dni - 9").Range("r33:r56")
        min = Application.WorksheetFunction.min(zakres)
        max = Application.WorksheetFunction.max(zakres)
     ActiveSheet.ChartObjects("Wykres 29").Activate
     ActiveChart.Axes(xlValue).Select
        
       With ActiveChart.Axes(xlValue)
       .MinimumScale = Round((min - 20), 0)
       .MaximumScale = Round((max + 20), 0)
       End With
       
  Application.ScreenUpdating = True
  Workbooks("prognoza Gaz PM.xlsm").Sheets("Delta 2 dni - 9").Visible = True
End Sub

Private Sub SpinButton4_Change()
  Dim strSQL As String
  Set cn = New ADODB.Connection
  Set rst = New ADODB.Recordset
  Set objaccess = CreateObject("Access.Application")
  Dim ldate As Date
  Dim start_row As Byte
  start_row = 33
  Application.ScreenUpdating = False
  Application.DisplayAlerts = False
  cn.Open "Provider=Microsoft.ACE.OLEDB.15.0; Data Source=" & baza_gaz
  ldate = VBA.Format(Workbooks("prognoza Gaz PM.xlsm").Sheets("Delta 2 dni - 9").Range("T30"), "yyyy-mm-dd")
  Workbooks("prognoza Gaz PM.xlsm").Sheets("Delta 2 dni - 9").Range("U33:X56").ClearContents
'Wykonanie pogody
  strSQL = "SELECT Wykonanie!Temperatura, Wykonanie!Naslon FROM Wykonanie WHERE (Wykonanie!Godzina > 6 AND Wykonanie!Obszar_ID =" & Obszar_pogoda _
  & " and Wykonanie.Doba = #" & ldate & "#)  ;"
  rst.Open strSQL, cn
  Workbooks("prognoza Gaz PM.xlsm").Sheets("Delta 2 dni - 9").Cells(start_row, 21).CopyFromRecordset rst
  strSQL = "SELECT Wykonanie!Temperatura, Wykonanie!Naslon FROM Wykonanie WHERE (Wykonanie!Godzina < 7 AND Wykonanie!Obszar_ID =" & Obszar_pogoda _
  & " and Wykonanie.Doba = #" & ldate + 1 & "#);"
  Set rst = New ADODB.Recordset
  rst.Open strSQL, cn
  Workbooks("prognoza Gaz PM.xlsm").Sheets("Delta 2 dni - 9").Cells(start_row + 17, 21).CopyFromRecordset rst
  
   'Wykonanie zużycia kWh
 
'Wykonanie zuzycia
  Set rst = New ADODB.Recordset
  strSQL = "SELECT Wolumen_wykon!Doba, Wolumen_wykon!H7, Wolumen_wykon!H8, Wolumen_wykon!H9, Wolumen_wykon!H10, Wolumen_wykon!H11, Wolumen_wykon!H12, Wolumen_wykon!H13, Wolumen_wykon!H14, Wolumen_wykon!H15, Wolumen_wykon!H16, Wolumen_wykon!H17," _
  & "Wolumen_wykon!H18, Wolumen_wykon!H19, Wolumen_wykon!H20, Wolumen_wykon!H21, Wolumen_wykon!H22, Wolumen_wykon!H23, Wolumen_wykon!H0, Wolumen_wykon!H1, Wolumen_wykon!H2, Wolumen_wykon!H3, Wolumen_wykon!H4, Wolumen_wykon!H5," _
  & "Wolumen_wykon!H6 FROM Wolumen_wykon WHERE (Wolumen_wykon!ID_Klient =" & ID_Klient & " and (Wolumen_wykon.Doba = #" & ldate & "#));"
  rst.Open strSQL, cn
  Dim objFields As ADODB.Fields
  Set objFields = rst.Fields
  
  Do While (rst.EOF = False)
    Workbooks("prognoza Gaz PM.xlsm").Sheets("Delta 2 dni - 9").Cells(start_row - 0, 24) = objFields.Item("Wolumen_wykon!H7").Value / 1000
    Workbooks("prognoza Gaz PM.xlsm").Sheets("Delta 2 dni - 9").Cells(start_row + 1, 24) = objFields.Item("Wolumen_wykon!H8").Value / 1000
    Workbooks("prognoza Gaz PM.xlsm").Sheets("Delta 2 dni - 9").Cells(start_row + 2, 24) = objFields.Item("Wolumen_wykon!H9").Value / 1000
    Workbooks("prognoza Gaz PM.xlsm").Sheets("Delta 2 dni - 9").Cells(start_row + 3, 24) = objFields.Item("Wolumen_wykon!H10").Value / 1000
    Workbooks("prognoza Gaz PM.xlsm").Sheets("Delta 2 dni - 9").Cells(start_row + 4, 24) = objFields.Item("Wolumen_wykon!H11").Value / 1000
    Workbooks("prognoza Gaz PM.xlsm").Sheets("Delta 2 dni - 9").Cells(start_row + 5, 24) = objFields.Item("Wolumen_wykon!H12").Value / 1000
    Workbooks("prognoza Gaz PM.xlsm").Sheets("Delta 2 dni - 9").Cells(start_row + 6, 24) = objFields.Item("Wolumen_wykon!H13").Value / 1000
    Workbooks("prognoza Gaz PM.xlsm").Sheets("Delta 2 dni - 9").Cells(start_row + 7, 24) = objFields.Item("Wolumen_wykon!H14").Value / 1000
    Workbooks("prognoza Gaz PM.xlsm").Sheets("Delta 2 dni - 9").Cells(start_row + 8, 24) = objFields.Item("Wolumen_wykon!H15").Value / 1000
    Workbooks("prognoza Gaz PM.xlsm").Sheets("Delta 2 dni - 9").Cells(start_row + 9, 24) = objFields.Item("Wolumen_wykon!H16").Value / 1000
    Workbooks("prognoza Gaz PM.xlsm").Sheets("Delta 2 dni - 9").Cells(start_row + 10, 24) = objFields.Item("Wolumen_wykon!H17").Value / 1000
    Workbooks("prognoza Gaz PM.xlsm").Sheets("Delta 2 dni - 9").Cells(start_row + 11, 24) = objFields.Item("Wolumen_wykon!H18").Value / 1000
    Workbooks("prognoza Gaz PM.xlsm").Sheets("Delta 2 dni - 9").Cells(start_row + 12, 24) = objFields.Item("Wolumen_wykon!H19").Value / 1000
    Workbooks("prognoza Gaz PM.xlsm").Sheets("Delta 2 dni - 9").Cells(start_row + 13, 24) = objFields.Item("Wolumen_wykon!H20").Value / 1000
    Workbooks("prognoza Gaz PM.xlsm").Sheets("Delta 2 dni - 9").Cells(start_row + 14, 24) = objFields.Item("Wolumen_wykon!H21").Value / 1000
    Workbooks("prognoza Gaz PM.xlsm").Sheets("Delta 2 dni - 9").Cells(start_row + 15, 24) = objFields.Item("Wolumen_wykon!H22").Value / 1000
    Workbooks("prognoza Gaz PM.xlsm").Sheets("Delta 2 dni - 9").Cells(start_row + 16, 24) = objFields.Item("Wolumen_wykon!H23").Value / 1000
    Workbooks("prognoza Gaz PM.xlsm").Sheets("Delta 2 dni - 9").Cells(start_row + 17, 24) = objFields.Item("Wolumen_wykon!H0").Value / 1000
    Workbooks("prognoza Gaz PM.xlsm").Sheets("Delta 2 dni - 9").Cells(start_row + 18, 24) = objFields.Item("Wolumen_wykon!H1").Value / 1000
    Workbooks("prognoza Gaz PM.xlsm").Sheets("Delta 2 dni - 9").Cells(start_row + 19, 24) = objFields.Item("Wolumen_wykon!H2").Value / 1000
    Workbooks("prognoza Gaz PM.xlsm").Sheets("Delta 2 dni - 9").Cells(start_row + 20, 24) = objFields.Item("Wolumen_wykon!H3").Value / 1000
    Workbooks("prognoza Gaz PM.xlsm").Sheets("Delta 2 dni - 9").Cells(start_row + 21, 24) = objFields.Item("Wolumen_wykon!H4").Value / 1000
    Workbooks("prognoza Gaz PM.xlsm").Sheets("Delta 2 dni - 9").Cells(start_row + 22, 24) = objFields.Item("Wolumen_wykon!H5").Value / 1000
    Workbooks("prognoza Gaz PM.xlsm").Sheets("Delta 2 dni - 9").Cells(start_row + 23, 24) = objFields.Item("Wolumen_wykon!H6").Value / 1000
    rst.MoveNext
  Loop
  Application.ScreenUpdating = True
  Application.DisplayAlerts = True
  Set rst = Nothing
  Set objFields = Nothing
  cn.Close
  
    'deklaracja zmiennych do skalowania
       Dim zakres As Range
       Dim max, min As Integer
       
    'scalowanie wykresu
    'ustawia zakres
  Set zakres = Workbooks("prognoza Gaz PM.xlsm").Sheets("Delta 2 dni - 9").Range("x33:x56")
        min = Application.WorksheetFunction.min(zakres)
        max = Application.WorksheetFunction.max(zakres)
     ActiveSheet.ChartObjects("Wykres 29").Activate
     ActiveChart.Axes(xlValue).Select
        
       With ActiveChart.Axes(xlValue)
       .MinimumScale = Round((min - 20), 0)
       .MaximumScale = Round((max + 20), 0)
       End With
       
  Application.ScreenUpdating = True
  Workbooks("prognoza Gaz PM.xlsm").Sheets("Delta 2 dni - 9").Visible = True
End Sub

