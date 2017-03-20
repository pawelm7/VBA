Option Explicit

Private Sub CheckBox1_Click()
    Dim suma, i As Double
    suma = 0
    For i = 0 To 26 Step 1
         If Not Workbooks("prognoza Gaz PM.xlsm").Sheets("Wykresy - 2").Cells(46, 3 + i).Value = "" Then
          suma = suma + Workbooks("prognoza Gaz PM.xlsm").Sheets("Wykresy - 2").Cells(46, 3 + i).Value
          End If
    Next i

    If suma <> 0 Then

    'widocznoasc na wykresie seria 1
    If Not CheckBox1.Value = True Then
    ActiveSheet.ChartObjects("Wykres 1").Activate
          ActiveChart.SeriesCollection(1).Select
        With Selection.Border
            .Weight = xlThin
            .LineStyle = xlNone
        End With
        With Selection
            .MarkerBackgroundColorIndex = xlNone
            .MarkerForegroundColorIndex = xlNone
            .MarkerStyle = xlNone
            .Smooth = False
            .MarkerSize = 3
            .Shadow = False
        End With
    Else

    ActiveSheet.ChartObjects("Wykres 1").Activate
        ActiveChart.SeriesCollection(1).Select
     With Selection.Border
            .ColorIndex = 1
            .Weight = xlThin
            .LineStyle = xlContinuous
        End With
        With Selection
            .MarkerBackgroundColorIndex = xlNone
            .MarkerForegroundColorIndex = xlNone
            .MarkerStyle = xlNone
            .Smooth = False
            .MarkerSize = 3
            .Shadow = False
        End With
    End If
    Application.ScreenUpdating = False
    If CheckBox1.Value = True Then
    Workbooks("prognoza Gaz PM.xlsm").Sheets("Wykresy - 2").Range("c46:ac46").Copy
    Workbooks("prognoza Gaz PM.xlsm").Sheets("Transponuj i wykres -6 ").Range("b5:ab5").PasteSpecial xlPasteValues
    Else
    Workbooks("prognoza Gaz PM.xlsm").Sheets("Transponuj i wykres -6 ").Range("b5:ab5").ClearContents
    End If
    Application.ScreenUpdating = True

    'deklaracja zmiennych do skalowania
       Dim zakres As Range
       Dim max, min As Integer
       
    'scalowanie wykresu
    'ustawia zakres
     Set zakres = Workbooks("prognoza Gaz PM.xlsm").Sheets("Transponuj i wykres -6 ").Range("b5:ab9")
        min = Application.WorksheetFunction.min(zakres)
        max = Application.WorksheetFunction.max(zakres)
     ActiveSheet.ChartObjects("Wykres 1").Activate
     ActiveChart.Axes(xlValue).Select
        
       With ActiveChart.Axes(xlValue)
       .MinimumScale = Round((min - 10), 0)
       .MaximumScale = Round((max + 10), 0)
       End With
    End If

End Sub

Private Sub CheckBox2_Click()
    Dim suma, i As Double
    suma = 0
    For i = 0 To 26 Step 1
      If Not Workbooks("prognoza Gaz PM.xlsm").Sheets("Wykresy - 2").Cells(47, 3 + i).Value = "" Then
          suma = suma + Workbooks("prognoza Gaz PM.xlsm").Sheets("Wykresy - 2").Cells(47, 3 + i).Value
          End If
    Next i

    If suma <> 0 Then

    'widocznoasc na wykresie seria 2
    If Not CheckBox2.Value = True Then
    ActiveSheet.ChartObjects("Wykres 1").Activate
        ActiveChart.SeriesCollection(2).Select
        With Selection.Border
            .Weight = xlThin
            .LineStyle = xlNone
        End With
        With Selection
            .MarkerBackgroundColorIndex = xlNone
            .MarkerForegroundColorIndex = xlNone
            .MarkerStyle = xlNone
            .Smooth = False
            .MarkerSize = 3
            .Shadow = False
        End With
    Else
    ActiveSheet.ChartObjects("Wykres 1").Activate
        ActiveChart.SeriesCollection(2).Select
     With Selection.Border
            .ColorIndex = 7
            .Weight = xlThin
            .LineStyle = xlContinuous
        End With
        With Selection
            .MarkerBackgroundColorIndex = xlNone
            .MarkerForegroundColorIndex = xlNone
            .MarkerStyle = xlNone
            .Smooth = False
            .MarkerSize = 3
            .Shadow = False
        End With
    End If
    Application.ScreenUpdating = False
    If CheckBox2.Value = True Then
    Workbooks("prognoza Gaz PM.xlsm").Sheets("Wykresy - 2").Range("c47:ac47").Copy
    Workbooks("prognoza Gaz PM.xlsm").Sheets("Transponuj i wykres -6 ").Range("b6:ab6").PasteSpecial xlPasteValues
    Else
    Workbooks("prognoza Gaz PM.xlsm").Sheets("Transponuj i wykres -6 ").Range("b6:ab6").ClearContents
    End If
    Application.ScreenUpdating = True

    'deklaracja zmiennych do skalowania
       Dim zakres As Range
       Dim max, min As Integer
       
    'scalowanie wykresu
    'ustawia zakres
     Set zakres = Workbooks("prognoza Gaz PM.xlsm").Sheets("Transponuj i wykres -6 ").Range("b5:ab9")
        min = Application.WorksheetFunction.min(zakres)
        max = Application.WorksheetFunction.max(zakres)
     ActiveSheet.ChartObjects("Wykres 1").Activate
     ActiveChart.Axes(xlValue).Select
        
       With ActiveChart.Axes(xlValue)
       .MinimumScale = Round((min - 10), 0)
       .MaximumScale = Round((max + 10), 0)
       End With
    End If

End Sub

Private Sub CheckBox3_Click()
    Dim suma, i As Double
    suma = 0
    For i = 0 To 26 Step 1
      If Not Workbooks("prognoza Gaz PM.xlsm").Sheets("Wykresy - 2").Cells(48, 3 + i).Value = "" Then
          suma = suma + Workbooks("prognoza Gaz PM.xlsm").Sheets("Wykresy - 2").Cells(48, 3 + i).Value
          End If
    Next i

    If suma <> 0 Then

    'widocznoasc na wykresie seria 3
    If Not CheckBox3.Value = True Then
    ActiveSheet.ChartObjects("Wykres 1").Activate
        ActiveChart.SeriesCollection(3).Select
        With Selection.Border
            .Weight = xlThin
            .LineStyle = xlNone
        End With
        With Selection
            .MarkerBackgroundColorIndex = xlNone
            .MarkerForegroundColorIndex = xlNone
            .MarkerStyle = xlNone
            .Smooth = False
            .MarkerSize = 3
            .Shadow = False
        End With
    Else
    ActiveSheet.ChartObjects("Wykres 1").Activate
        ActiveChart.SeriesCollection(3).Select
     With Selection.Border
            .ColorIndex = 5
            .Weight = xlThin
            .LineStyle = xlContinuous
        End With
        With Selection
            .MarkerBackgroundColorIndex = xlNone
            .MarkerForegroundColorIndex = xlNone
            .MarkerStyle = xlNone
            .Smooth = False
            .MarkerSize = 3
            .Shadow = False
        End With
    End If
    Application.ScreenUpdating = False
    If CheckBox3.Value = True Then
    Workbooks("prognoza Gaz PM.xlsm").Sheets("Wykresy - 2").Range("c48:ac48").Copy
    Workbooks("prognoza Gaz PM.xlsm").Sheets("Transponuj i wykres -6 ").Range("b7:ab7").PasteSpecial xlPasteValues
    Else
    Workbooks("prognoza Gaz PM.xlsm").Sheets("Transponuj i wykres -6 ").Range("b7:ab7").ClearContents
    End If
    Application.ScreenUpdating = True

    'deklaracja zmiennych do skalowania
       Dim zakres As Range
       Dim max, min As Integer
       
    'scalowanie wykresu
    'ustawia zakres
     Set zakres = Workbooks("prognoza Gaz PM.xlsm").Sheets("Transponuj i wykres -6 ").Range("b5:ab9")
        min = Application.WorksheetFunction.min(zakres)
        max = Application.WorksheetFunction.max(zakres)
     ActiveSheet.ChartObjects("Wykres 1").Activate
     ActiveChart.Axes(xlValue).Select
        
       With ActiveChart.Axes(xlValue)
       .MinimumScale = Round((min - 10), 0)
       .MaximumScale = Round((max + 10), 0)
       End With
    End If
End Sub

Private Sub CheckBox4_Click()
    Dim suma, i As Double
    suma = 0
    For i = 0 To 26 Step 1
      If Not Workbooks("prognoza Gaz PM.xlsm").Sheets("Wykresy - 2").Cells(49, 3 + i).Value = "" Then
          suma = suma + Workbooks("prognoza Gaz PM.xlsm").Sheets("Wykresy - 2").Cells(49, 3 + i).Value
          End If
    Next i
    If suma <> 0 Then

    'widocznoasc na wykresie seria 4
    If Not CheckBox4.Value = True Then
    ActiveSheet.ChartObjects("Wykres 1").Activate
        ActiveChart.SeriesCollection(4).Select
        With Selection.Border
            .Weight = xlThin
            .LineStyle = xlNone
        End With
        With Selection
            .MarkerBackgroundColorIndex = xlNone
            .MarkerForegroundColorIndex = xlNone
            .MarkerStyle = xlNone
            .Smooth = False
            .MarkerSize = 3
            .Shadow = False
        End With
    Else
    ActiveSheet.ChartObjects("Wykres 1").Activate
        ActiveChart.SeriesCollection(4).Select
     With Selection.Border
            .ColorIndex = 10
            .Weight = xlThin
            .LineStyle = xlContinuous
        End With
        With Selection
            .MarkerBackgroundColorIndex = xlNone
            .MarkerForegroundColorIndex = xlNone
            .MarkerStyle = xlNone
            .Smooth = False
            .MarkerSize = 3
            .Shadow = False
        End With
    End If
    Application.ScreenUpdating = False
    If CheckBox4.Value = True Then
    Workbooks("prognoza Gaz PM.xlsm").Sheets("Wykresy - 2").Range("c49:ac49").Copy
    Workbooks("prognoza Gaz PM.xlsm").Sheets("Transponuj i wykres -6 ").Range("b8:ab8").PasteSpecial xlPasteValues
    Else
    Workbooks("prognoza Gaz PM.xlsm").Sheets("Transponuj i wykres -6 ").Range("b8:ab8").ClearContents
    End If
    Application.ScreenUpdating = True

    'deklaracja zmiennych do skalowania
       Dim zakres As Range
       Dim max, min As Integer
       
    'scalowanie wykresu
    'ustawia zakres
     Set zakres = Workbooks("prognoza Gaz PM.xlsm").Sheets("Transponuj i wykres -6 ").Range("b5:ab9")
        min = Application.WorksheetFunction.min(zakres)
        max = Application.WorksheetFunction.max(zakres)
     ActiveSheet.ChartObjects("Wykres 1").Activate
     ActiveChart.Axes(xlValue).Select
        
       With ActiveChart.Axes(xlValue)
       .MinimumScale = Round((min - 10), 0)
       .MaximumScale = Round((max + 10), 0)
       End With
    End If
End Sub

Private Sub CheckBox5_Click()
    'widocznoasc na wykresie seria 5 srednia
    If Not CheckBox5.Value = True Then
    ActiveSheet.ChartObjects("wykres 1").Activate
        ActiveChart.SeriesCollection(5).Select
        With Selection.Border
            .Weight = xlThin
            .LineStyle = xlNone
        End With
        With Selection
            .MarkerBackgroundColorIndex = xlNone
            .MarkerForegroundColorIndex = xlNone
            .MarkerStyle = xlNone
            .Smooth = False
            .MarkerSize = 3
            .Shadow = False
        End With
    Else
    ActiveSheet.ChartObjects("wykres 1").Activate
        ActiveChart.SeriesCollection(5).Select
     With Selection.Border
            .ColorIndex = 3
            .Weight = xlThick
            .LineStyle = xlDot
        End With
        With Selection
            .MarkerBackgroundColorIndex = xlNone
            .MarkerForegroundColorIndex = xlNone
            .MarkerStyle = xlNone
            .Smooth = False
            .MarkerSize = 3
            .Shadow = False
        End With
    End If
End Sub

Private Sub CommandButton1_Click()
    'Œrednia z danych dodanych do wykresu
    Dim suma As Double
    Dim liczba As Double
    Dim i As Integer
    Dim srednia As Double

    For i = 0 To 26 Step 1
        suma = 0
        liczba = 0
        If CheckBox1.Value = True Then
       If Not Workbooks("prognoza Gaz PM.xlsm").Sheets("Wykresy - 2").Cells(46, 3 + i) = "" Then
        suma = Workbooks("prognoza Gaz PM.xlsm").Sheets("Wykresy - 2").Cells(46, 3 + i) + suma
        liczba = liczba + 1
        End If
      
        End If
        
        If CheckBox2.Value = True Then
        If Not Workbooks("prognoza Gaz PM.xlsm").Sheets("Wykresy - 2").Cells(47, 3 + i) = "" Then
        suma = Workbooks("prognoza Gaz PM.xlsm").Sheets("Wykresy - 2").Cells(47, 3 + i) + suma
       liczba = liczba + 1
       End If
        End If
        
        If CheckBox3.Value = True Then
        If Not Workbooks("prognoza Gaz PM.xlsm").Sheets("Wykresy - 2").Cells(48, 3 + i) = "" Then
        suma = Workbooks("prognoza Gaz PM.xlsm").Sheets("Wykresy - 2").Cells(48, 3 + i) + suma
        liczba = liczba + 1
        End If
        End If
        
        If CheckBox4.Value = True Then
        If Not Workbooks("prognoza Gaz PM.xlsm").Sheets("Wykresy - 2").Cells(49, 3 + i) = "" Then
        suma = Workbooks("prognoza Gaz PM.xlsm").Sheets("Wykresy - 2").Cells(49, 3 + i) + suma
        liczba = liczba + 1
        
    End If
        End If

       If liczba = 0 Then
        Workbooks("prognoza Gaz PM.xlsm").Sheets("Wykresy - 2").Cells(50, 3 + i).Value = 0
     Else
       srednia = suma / liczba
       Workbooks("prognoza Gaz PM.xlsm").Sheets("Wykresy - 2").Cells(50, 3 + i) = srednia
       Workbooks("prognoza Gaz PM.xlsm").Sheets("Wykresy - 2").Cells(9, 2 + i) = srednia
       End If
    Next i

    Application.ScreenUpdating = False
    If CheckBox1.Value = True Then
    Workbooks("prognoza Gaz PM.xlsm").Sheets("Wykresy - 2").Range("c46:ac46").Copy
    Workbooks("prognoza Gaz PM.xlsm").Sheets("Transponuj i wykres -6 ").Range("b5:ab5").PasteSpecial xlPasteValues
    Else
    Workbooks("prognoza Gaz PM.xlsm").Sheets("Transponuj i wykres -6 ").Range("b5:ab5").ClearContents
    End If
    Application.ScreenUpdating = True

    Application.ScreenUpdating = False
    If CheckBox2.Value = True Then
    Workbooks("prognoza Gaz PM.xlsm").Sheets("Wykresy - 2").Range("c47:ac47").Copy
    Workbooks("prognoza Gaz PM.xlsm").Sheets("Transponuj i wykres -6 ").Range("b6:ab6").PasteSpecial xlPasteValues
    Else
    Workbooks("prognoza Gaz PM.xlsm").Sheets("Transponuj i wykres -6 ").Range("b6:ab6").ClearContents
    End If
    Application.ScreenUpdating = True

    Application.ScreenUpdating = False
    If CheckBox3.Value = True Then
    Workbooks("prognoza Gaz PM.xlsm").Sheets("Wykresy - 2").Range("c48:ac48").Copy
    Workbooks("prognoza Gaz PM.xlsm").Sheets("Transponuj i wykres -6 ").Range("b7:ab7").PasteSpecial xlPasteValues
    Else
    Workbooks("prognoza Gaz PM.xlsm").Sheets("Transponuj i wykres -6 ").Range("b7:ab7").ClearContents
    End If
    Application.ScreenUpdating = True

    Application.ScreenUpdating = False
    If CheckBox4.Value = True Then
    Workbooks("prognoza Gaz PM.xlsm").Sheets("Wykresy - 2").Range("c49:ac49").Copy
    Workbooks("prognoza Gaz PM.xlsm").Sheets("Transponuj i wykres -6 ").Range("b8:ab8").PasteSpecial xlPasteValues
    Else
    Workbooks("prognoza Gaz PM.xlsm").Sheets("Transponuj i wykres -6 ").Range("b8:ab8").ClearContents
    End If
    Application.ScreenUpdating = True

    'skopiowanie œredniej
    Workbooks("prognoza Gaz PM.xlsm").Sheets("Wykresy - 2").Range("c50:ac50").Copy
    Workbooks("prognoza Gaz PM.xlsm").Sheets("Transponuj i wykres -6 ").Range("b9:ab9").PasteSpecial xlPasteValues

    'deklaracja zmiennych do skalowania
       Dim zakres As Range
       Dim max, min As Integer
       
    'scalowanie wykresu
    'ustawia zakres
     Set zakres = Workbooks("prognoza Gaz PM.xlsm").Sheets("Transponuj i wykres -6 ").Range("b5:ab9")
        min = Application.WorksheetFunction.min(zakres)
        max = Application.WorksheetFunction.max(zakres)
     ActiveSheet.ChartObjects("Wykres 1").Activate
     ActiveChart.Axes(xlValue).Select
        
       With ActiveChart.Axes(xlValue)
       .MinimumScale = Round((min - 10), 0)
       .MaximumScale = Round((max + 10), 0)
       End With
End Sub

Private Sub CommandButton2_Click()
    'wyszukiwanie dni podobnych pogodowo
    UserForm1.Show
End Sub

Private Sub CommandButton3_Click()
	On Error GoTo ErrHandler
	Dim strChaine As String
	Dim Sql as string
	strChaine = VBA.Format(Workbooks("prognoza Gaz PM.xlsm").Sheets("Analiza - 1").Range("F4"), "yyyy-mm-dd")
	Set objaccess = CreateObject("Access.Application")
	Set cn = New ADODB.Connection
    Set rst = New ADODB.Recordset
    
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    cn.Open "Provider=Microsoft.ACE.OLEDB.15.0; Data Source=" & baza_gaz
    objaccess.OpenCurrentDatabase baza_gaz
   
    rst.Open "Select Wolumen_prognoza.Data_przygotowania from Wolumen_prognoza where Wolumen_prognoza.Data_prognozy =#" & strChaine & "#;", cn
    If rst.EOF = False Then
    MsgBox "Prognozy zu¿ycia na dzieñ " & strChaine & " jest ju¿ sporz¹dzona!"
	Else
   'zapis prognozy EE
    Dim h7 As Double
    Dim h8 As Double
    Dim h9 As Double
    Dim h10 As Double
    Dim h11 As Double
    Dim h12 As Double
    Dim h13 As Double
    Dim h14 As Double
    Dim h15 As Double
    Dim h16 As Double
    Dim h17 As Double
    Dim h18 As Double
    Dim h19 As Double
    Dim h20 As Double
    Dim h21 As Double
    Dim h22 As Double
    Dim h23 As Double
    Dim H0 As Double
    Dim h1 As Double
    Dim h2 As Double
    Dim h3 As Double
    Dim h4 As Double
    Dim h5 As Double
    Dim h6 As Double

    h7 = Sheets("Wykresy - 2").Cells(53, 6).Value
    h8 = Sheets("Wykresy - 2").Cells(53, 7).Value
    h9 = Sheets("Wykresy - 2").Cells(53, 8).Value
    h10 = Sheets("Wykresy - 2").Cells(53, 9).Value
    h11 = Sheets("Wykresy - 2").Cells(53, 10).Value
    h12 = Sheets("Wykresy - 2").Cells(53, 11).Value
    h13 = Sheets("Wykresy - 2").Cells(53, 12).Value
    h14 = Sheets("Wykresy - 2").Cells(53, 13).Value
    h15 = Sheets("Wykresy - 2").Cells(53, 14).Value
    h16 = Sheets("Wykresy - 2").Cells(53, 15).Value
    h17 = Sheets("Wykresy - 2").Cells(53, 16).Value
    h18 = Sheets("Wykresy - 2").Cells(53, 17).Value
    h19 = Sheets("Wykresy - 2").Cells(53, 18).Value
    h20 = Sheets("Wykresy - 2").Cells(53, 19).Value
    h21 = Sheets("Wykresy - 2").Cells(53, 20).Value
    h22 = Sheets("Wykresy - 2").Cells(53, 21).Value
    h23 = Sheets("Wykresy - 2").Cells(53, 22).Value
    H0 = Sheets("Wykresy - 2").Cells(53, 23).Value
    h1 = Sheets("Wykresy - 2").Cells(53, 24).Value
    h2 = Sheets("Wykresy - 2").Cells(53, 25).Value
    h3 = Sheets("Wykresy - 2").Cells(53, 26).Value
    h4 = Sheets("Wykresy - 2").Cells(53, 27).Value
    h5 = Sheets("Wykresy - 2").Cells(53, 28).Value
    h6 = Sheets("Wykresy - 2").Cells(53, 29).Value
    Sql = "Insert Into Wolumen_prognoza (Data_przygotowania, time, ID_User, ID_klient, Doba_prognozy, H7, H8, H9, H10, H11, H12, H13, H14, H15, H16, H17, H18, H19, H20, H21, H22, H23, H0, H1, H2, H3, H4, H5, H6) Values (" _
    & Date() & ", " & Now() & ", " & userID & ", " & ID_Klient & ", '" & strChaine & "', '" & h7 & "', '" & h8 & "', '" & h9 & "', '" & h10 & "', '" & h11 & "', '" & h12 & "', '" & h13 & "', '" & h14 & "', '" & h15 & "', '" & h16 & "', '" & h17 & "', '" _
    & h18 & "', '" & h19 & "', '" & h20 & "', '" & h21 & "', '" & h22 & "', '" & h23 & "', '" & H0 & "', '" & h1 & "', '" & h2 & "', '" & h3 & "', '" & h4 & "', '" & h5 & "', '" & h6 & "')"
    objaccess.OpenCurrentDatabase baza_gaz
    objaccess.DoCmd.RunSQL Sql
            

       MsgBox "Prognoza zu¿ycia na dzieñ " & strChaine & " zapisana poprawnie."
    End If
    
    Set objaccess = Nothing
	ErrHandler:
    'If Err.Number = 91 Then
     '   MsgBox "Nie ma nastêpnego arkusza."
    'Else
    'inny b³¹d
        MsgBox "Wyst¹pi³ b³¹d nr: " & Err.Number & Chr(10) & Err.Description
    'End If
    Err.Clear	
End Sub

Private Sub SpinButton1_SpinUp()
    Workbooks("prognoza Gaz PM.xlsm").Sheets("Wykresy - 2").Range("f50").Value = Workbooks("prognoza Gaz PM.xlsm").Sheets("Wykresy - 2").Range("f50").Value + 1
End Sub

Private Sub SpinButton1_SpinDown()
    Workbooks("prognoza Gaz PM.xlsm").Sheets("Wykresy - 2").Range("f50").Value = Workbooks("prognoza Gaz PM.xlsm").Sheets("Wykresy - 2").Range("f50").Value - 1
End Sub

Private Sub SpinButton2_SpinUp()
    Workbooks("prognoza Gaz PM.xlsm").Sheets("Wykresy - 2").Range("g50").Value = Workbooks("prognoza Gaz PM.xlsm").Sheets("Wykresy - 2").Range("g50").Value + 1
End Sub

Private Sub SpinButton2_SpinDown()
    Workbooks("prognoza Gaz PM.xlsm").Sheets("Wykresy - 2").Range("g50").Value = Workbooks("prognoza Gaz PM.xlsm").Sheets("Wykresy - 2").Range("g50").Value - 1
End Sub

Private Sub SpinButton3_SpinUp()
    Workbooks("prognoza Gaz PM.xlsm").Sheets("Wykresy - 2").Range("h50").Value = Workbooks("prognoza Gaz PM.xlsm").Sheets("Wykresy - 2").Range("h50").Value + 1
End Sub

Private Sub SpinButton3_SpinDown()
    Workbooks("prognoza Gaz PM.xlsm").Sheets("Wykresy - 2").Range("h50").Value = Workbooks("prognoza Gaz PM.xlsm").Sheets("Wykresy - 2").Range("h50").Value - 1
End Sub

Private Sub SpinButton4_SpinUp()
    Workbooks("prognoza Gaz PM.xlsm").Sheets("Wykresy - 2").Range("i50").Value = Workbooks("prognoza Gaz PM.xlsm").Sheets("Wykresy - 2").Range("i50").Value + 1
End Sub

Private Sub SpinButton4_SpinDown()
    Workbooks("prognoza Gaz PM.xlsm").Sheets("Wykresy - 2").Range("i50").Value = Workbooks("prognoza Gaz PM.xlsm").Sheets("Wykresy - 2").Range("i50").Value - 1
End Sub

Private Sub SpinButton5_SpinUp()
    Workbooks("prognoza Gaz PM.xlsm").Sheets("Wykresy - 2").Range("j50").Value = Workbooks("prognoza Gaz PM.xlsm").Sheets("Wykresy - 2").Range("j50").Value + 1
End Sub

Private Sub SpinButton5_SpinDown()
    Workbooks("prognoza Gaz PM.xlsm").Sheets("Wykresy - 2").Range("j50").Value = Workbooks("prognoza Gaz PM.xlsm").Sheets("Wykresy - 2").Range("j50").Value - 1
End Sub

Private Sub SpinButton6_SpinUp()
    Workbooks("prognoza Gaz PM.xlsm").Sheets("Wykresy - 2").Range("k50").Value = Workbooks("prognoza Gaz PM.xlsm").Sheets("Wykresy - 2").Range("k50").Value + 1
End Sub

Private Sub SpinButton6_SpinDown()
    Workbooks("prognoza Gaz PM.xlsm").Sheets("Wykresy - 2").Range("k50").Value = Workbooks("prognoza Gaz PM.xlsm").Sheets("Wykresy - 2").Range("k50").Value - 1
End Sub

Private Sub SpinButton7_SpinUp()
    Workbooks("prognoza Gaz PM.xlsm").Sheets("Wykresy - 2").Range("l50").Value = Workbooks("prognoza Gaz PM.xlsm").Sheets("Wykresy - 2").Range("l50").Value + 1
End Sub

Private Sub SpinButton7_SpinDown()
    Workbooks("prognoza Gaz PM.xlsm").Sheets("Wykresy - 2").Range("l50").Value = Workbooks("prognoza Gaz PM.xlsm").Sheets("Wykresy - 2").Range("l50").Value - 1
End Sub

Private Sub SpinButton8_SpinUp()
    Workbooks("prognoza Gaz PM.xlsm").Sheets("Wykresy - 2").Range("m50").Value = Workbooks("prognoza Gaz PM.xlsm").Sheets("Wykresy - 2").Range("m50").Value + 1
End Sub

Private Sub SpinButton8_SpinDown()
    Workbooks("prognoza Gaz PM.xlsm").Sheets("Wykresy - 2").Range("m50").Value = Workbooks("prognoza Gaz PM.xlsm").Sheets("Wykresy - 2").Range("m50").Value - 1
End Sub

Private Sub SpinButton9_SpinUp()
    Workbooks("prognoza Gaz PM.xlsm").Sheets("Wykresy - 2").Range("n50").Value = Workbooks("prognoza Gaz PM.xlsm").Sheets("Wykresy - 2").Range("n50").Value + 1
End Sub

Private Sub SpinButton9_SpinDown()
    Workbooks("prognoza Gaz PM.xlsm").Sheets("Wykresy - 2").Range("n50").Value = Workbooks("prognoza Gaz PM.xlsm").Sheets("Wykresy - 2").Range("n50").Value - 1
End Sub

Private Sub SpinButton10_SpinUp()
    Workbooks("prognoza Gaz PM.xlsm").Sheets("Wykresy - 2").Range("o50").Value = Workbooks("prognoza Gaz PM.xlsm").Sheets("Wykresy - 2").Range("o50").Value + 1
End Sub

Private Sub SpinButton10_SpinDown()
    Workbooks("prognoza Gaz PM.xlsm").Sheets("Wykresy - 2").Range("o50").Value = Workbooks("prognoza Gaz PM.xlsm").Sheets("Wykresy - 2").Range("o50").Value - 1
End Sub

Private Sub SpinButton11_SpinUp()
    Workbooks("prognoza Gaz PM.xlsm").Sheets("Wykresy - 2").Range("p50").Value = Workbooks("prognoza Gaz PM.xlsm").Sheets("Wykresy - 2").Range("p50").Value + 1
End Sub

Private Sub SpinButton11_SpinDown()
    Workbooks("prognoza Gaz PM.xlsm").Sheets("Wykresy - 2").Range("p50").Value = Workbooks("prognoza Gaz PM.xlsm").Sheets("Wykresy - 2").Range("p50").Value - 1
End Sub

Private Sub SpinButton12_SpinUp()
    Workbooks("prognoza Gaz PM.xlsm").Sheets("Wykresy - 2").Range("q50").Value = Workbooks("prognoza Gaz PM.xlsm").Sheets("Wykresy - 2").Range("q50").Value + 1
End Sub

Private Sub SpinButton12_SpinDown()
    Workbooks("prognoza Gaz PM.xlsm").Sheets("Wykresy - 2").Range("q50").Value = Workbooks("prognoza Gaz PM.xlsm").Sheets("Wykresy - 2").Range("q50").Value - 1
End Sub

Private Sub SpinButton13_SpinUp()
    Workbooks("prognoza Gaz PM.xlsm").Sheets("Wykresy - 2").Range("r50").Value = Workbooks("prognoza Gaz PM.xlsm").Sheets("Wykresy - 2").Range("r50").Value + 1
End Sub

Private Sub SpinButton13_SpinDown()
    Workbooks("prognoza Gaz PM.xlsm").Sheets("Wykresy - 2").Range("r50").Value = Workbooks("prognoza Gaz PM.xlsm").Sheets("Wykresy - 2").Range("r50").Value - 1
End Sub

Private Sub SpinButton14_SpinUp()
    Workbooks("prognoza Gaz PM.xlsm").Sheets("Wykresy - 2").Range("s50").Value = Workbooks("prognoza Gaz PM.xlsm").Sheets("Wykresy - 2").Range("s50").Value + 1
End Sub

Private Sub SpinButton14_SpinDown()
    Workbooks("prognoza Gaz PM.xlsm").Sheets("Wykresy - 2").Range("s50").Value = Workbooks("prognoza Gaz PM.xlsm").Sheets("Wykresy - 2").Range("s50").Value - 1
End Sub

Private Sub SpinButton15_SpinUp()
    Workbooks("prognoza Gaz PM.xlsm").Sheets("Wykresy - 2").Range("t50").Value = Workbooks("prognoza Gaz PM.xlsm").Sheets("Wykresy - 2").Range("t50").Value + 1
End Sub

Private Sub SpinButton15_SpinDown()
    Workbooks("prognoza Gaz PM.xlsm").Sheets("Wykresy - 2").Range("t50").Value = Workbooks("prognoza Gaz PM.xlsm").Sheets("Wykresy - 2").Range("t50").Value - 1
End Sub

Private Sub SpinButton16_SpinUp()
    Workbooks("prognoza Gaz PM.xlsm").Sheets("Wykresy - 2").Range("u50").Value = Workbooks("prognoza Gaz PM.xlsm").Sheets("Wykresy - 2").Range("u50").Value + 1
End Sub

Private Sub SpinButton16_SpinDown()
    Workbooks("prognoza Gaz PM.xlsm").Sheets("Wykresy - 2").Range("u50").Value = Workbooks("prognoza Gaz PM.xlsm").Sheets("Wykresy - 2").Range("u50").Value - 1
End Sub

Private Sub SpinButton17_SpinUp()
    Workbooks("prognoza Gaz PM.xlsm").Sheets("Wykresy - 2").Range("v50").Value = Workbooks("prognoza Gaz PM.xlsm").Sheets("Wykresy - 2").Range("v50").Value + 1
End Sub

Private Sub SpinButton17_SpinDown()
    Workbooks("prognoza Gaz PM.xlsm").Sheets("Wykresy - 2").Range("v50").Value = Workbooks("prognoza Gaz PM.xlsm").Sheets("Wykresy - 2").Range("v50").Value - 1
End Sub

Private Sub SpinButton18_SpinUp()
    Workbooks("prognoza Gaz PM.xlsm").Sheets("Wykresy - 2").Range("w50").Value = Workbooks("prognoza Gaz PM.xlsm").Sheets("Wykresy - 2").Range("w50").Value + 1
End Sub

Private Sub SpinButton18_SpinDown()
    Workbooks("prognoza Gaz PM.xlsm").Sheets("Wykresy - 2").Range("w50").Value = Workbooks("prognoza Gaz PM.xlsm").Sheets("Wykresy - 2").Range("w50").Value - 1
End Sub

Private Sub SpinButton19_SpinUp()
    Workbooks("prognoza Gaz PM.xlsm").Sheets("Wykresy - 2").Range("x50").Value = Workbooks("prognoza Gaz PM.xlsm").Sheets("Wykresy - 2").Range("x50").Value + 1
End Sub

Private Sub SpinButton19_SpinDown()
    Workbooks("prognoza Gaz PM.xlsm").Sheets("Wykresy - 2").Range("x50").Value = Workbooks("prognoza Gaz PM.xlsm").Sheets("Wykresy - 2").Range("x50").Value - 1
End Sub

Private Sub SpinButton20_SpinUp()
    Workbooks("prognoza Gaz PM.xlsm").Sheets("Wykresy - 2").Range("y50").Value = Workbooks("prognoza Gaz PM.xlsm").Sheets("Wykresy - 2").Range("y50").Value + 1
End Sub

Private Sub SpinButton20_SpinDown()
    Workbooks("prognoza Gaz PM.xlsm").Sheets("Wykresy - 2").Range("y50").Value = Workbooks("prognoza Gaz PM.xlsm").Sheets("Wykresy - 2").Range("y50").Value - 1
End Sub

Private Sub SpinButton21_SpinUp()
    Workbooks("prognoza Gaz PM.xlsm").Sheets("Wykresy - 2").Range("z50").Value = Workbooks("prognoza Gaz PM.xlsm").Sheets("Wykresy - 2").Range("z50").Value + 1
End Sub

Private Sub SpinButton21_SpinDown()
    Workbooks("prognoza Gaz PM.xlsm").Sheets("Wykresy - 2").Range("z50").Value = Workbooks("prognoza Gaz PM.xlsm").Sheets("Wykresy - 2").Range("z50").Value - 1
End Sub

Private Sub SpinButton22_SpinUp()
    Workbooks("prognoza Gaz PM.xlsm").Sheets("Wykresy - 2").Range("aa50").Value = Workbooks("prognoza Gaz PM.xlsm").Sheets("Wykresy - 2").Range("aa50").Value + 1
End Sub

Private Sub SpinButton22_SpinDown()
    Workbooks("prognoza Gaz PM.xlsm").Sheets("Wykresy - 2").Range("aa50").Value = Workbooks("prognoza Gaz PM.xlsm").Sheets("Wykresy - 2").Range("aa50").Value - 1
End Sub

Private Sub SpinButton23_SpinUp()
    Workbooks("prognoza Gaz PM.xlsm").Sheets("Wykresy - 2").Range("ab50").Value = Workbooks("prognoza Gaz PM.xlsm").Sheets("Wykresy - 2").Range("ab50").Value + 1
End Sub

Private Sub SpinButton23_SpinDown()
    Workbooks("prognoza Gaz PM.xlsm").Sheets("Wykresy - 2").Range("ab50").Value = Workbooks("prognoza Gaz PM.xlsm").Sheets("Wykresy - 2").Range("ab50").Value - 1
End Sub

Private Sub SpinButton24_SpinUp()
    Workbooks("prognoza Gaz PM.xlsm").Sheets("Wykresy - 2").Range("ac50").Value = Workbooks("prognoza Gaz PM.xlsm").Sheets("Wykresy - 2").Range("ac50").Value + 1
End Sub

Private Sub SpinButton24_SpinDown()
    Workbooks("prognoza Gaz PM.xlsm").Sheets("Wykresy - 2").Range("ac50").Value = Workbooks("prognoza Gaz PM.xlsm").Sheets("Wykresy - 2").Range("ac50").Value - 1
End Sub

