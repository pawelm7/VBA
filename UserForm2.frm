Option Explicit

Private Sub CommandButton1_Click()
        Dim rngTrouve As Range
        Dim strChaine As Date
        Dim wiersz As Integer
        Dim z, i, k As Integer
        Dim lastrow, startrow As Double
        Dim dzien, dzien1 As Double
        Dim suma As Double
        Dim godzina As Integer
        Dim m As Integer
        Dim zakres1 As Range
        Dim Answer As Double
        Dim h1, h2, h3, h4, h5, h6, h7, h8, h9, h10, h11, h12, h13, h14, h15, h16, h17, h18, h19, h20, h21, h22, h23, h24 As Integer
        Dim typ_param As String
        Dim strSQL As String
        Set cn = New ADODB.Connection
        Set objaccess = CreateObject("Access.Application")
        cn.Open "Provider=Microsoft.ACE.OLEDB.15.0; Data Source=" & baza_gaz
   
    strChaine = VBA.Format(Workbooks("prognoza Gaz PM.xlsm").Sheets("Analiza - 1").Range("F4"), "yyyy-mm-dd")
    Workbooks("prognoza Gaz PM.xlsm").Sheets("podobne").Range("a2:B43").ClearContents
    'ustawienie iloœci dnia cofniêcia
    If OptionButton4 = True Then
    'cofniecie daty poczatkowej o tydzien
        startrow = 9
    End If
    If OptionButton5 = True Then
    'cofniecie daty poczatkowej o 2 tyg
        startrow = 16
    End If
    If OptionButton6 = True Then
    'cofniecie daty poczatkowej o 3 tyg
        startrow = 23
    End If
    If OptionButton7 = True Then
    'cofniecie daty poczatkowej o 4tyg
        startrow = 32
    End If
      'macierz godzinowa
           If CheckBox1.Value = True Then
             h7 = 1
           Else
             h7 = 0
           End If
              
           If CheckBox2.Value = True Then
             h8 = 1
           Else
             h8 = 0
           End If
             
           If CheckBox3.Value = True Then
             h9 = 1
           Else
             h9 = 0
           End If
            
           If CheckBox4.Value = True Then
             h10 = 1
           Else
             h10 = 0
           End If
            
           If CheckBox5.Value = True Then
             h11 = 1
           Else
             h11 = 0
           End If
             
           If CheckBox6.Value = True Then
             h12 = 1
           Else
             h12 = 0
           End If
           
           If CheckBox7.Value = True Then
             h13 = 1
           Else
             h13 = 0
           End If
            
           If CheckBox8.Value = True Then
             h14 = 1
           Else
             h14 = 0
           End If
            
           If CheckBox9.Value = True Then
             h15 = 1
           Else
             h15 = 0
           End If
            
           If CheckBox10.Value = True Then
             h16 = 1
           Else
             h16 = 0
           End If
            
           If CheckBox11.Value = True Then
             h17 = 1
           Else
             h17 = 0
           End If
            
           If CheckBox12.Value = True Then
             h18 = 1
           Else
             h18 = 0
           End If
            
           If CheckBox13.Value = True Then
             h19 = 1
           Else
             h19 = 0
           End If
            
           If CheckBox14.Value = True Then
             h20 = 1
           Else
             h20 = 0
           End If
            
           If CheckBox15.Value = True Then
             h21 = 1
           Else
             h21 = 0
           End If
            
           If CheckBox16.Value = True Then
             h22 = 1
           Else
             h22 = 0
           End If
            
           If CheckBox17.Value = True Then
             h23 = 1
           Else
             h23 = 0
           End If
            
           If CheckBox18.Value = True Then
             H0 = 1
           Else
             H0 = 0
           End If
            
           If CheckBox19.Value = True Then
             h1 = 1
           Else
             h1 = 0
           End If
            
           If CheckBox20.Value = True Then
             h2 = 1
           Else
             h2 = 0
           End If
            
           If CheckBox21.Value = True Then
             h3 = 1
           Else
             h3 = 0
           End If
            
           If CheckBox22.Value = True Then
             h4 = 1
           Else
             h4 = 0
           End If
            
           If CheckBox23.Value = True Then
             h5 = 1
           Else
             h5 = 0
           End If
            
           If CheckBox24.Value = True Then
             h6 = 1
           Else
             h6 = 0
           End If
    If OptionButton1 = True Then
       ' TEMPERATURA
       typ_param = "Temperatura"
     End If
            
    If OptionButton2 = True Then
     ' Nas³on
       typ_param = "Naslon"
    End If
     'data_od
      Year_od = VBA.Format((strChaine - 1), "YYYY")
      month_od = VBA.Format((strChaine - 1), "MM")
      day_od = VBA.Format((strChaine - 1), "DD")
      
    ' zaczynam od godz 7 dla prognozy pogody
    Set rst = New ADODB.Recordset
    Set rst2 = New ADODB.Recordset
    strSQL = "SELECT Prognoza!Godzina, Prognoza!Doba, Prognoza!" & typ_param & " FROM Prognoza WHERE (Prognoza!Obszar_ID =" & Obszar_pogoda _
    & " and Prognoza.data_prognozy = '" & Year_od & month_od & day_od & "') and Prognoza.Doba Between #" & strChaine & "# and #" & strChaine + 1 & "# ORDER BY Prognoza.Doba, Prognoza.Godzina;"
    rst.Open strSQL, cn
    
    'Wykonania pogody
    wykon = "SELECT Wykonanie!Godzina, Wykonanie!Doba, Wykonanie!" & typ_param & " FROM Wykonanie WHERE Wykonanie!Obszar_ID =" & Obszar_pogoda _
    & " and Wykonanie!Doba Between #" & strChaine - startrow - 1 & "# and #" & strChaine - 1 & "# ORDER BY Wykonanie.Doba, Wykonanie.Godzina;"
    rst2.Open wykon, cn
    
    'Msgbox rst.getstring
      
    Dim objFields As ADODB.Fields
    Set objFields = rst.Fields
    rst.MoveFirst
    Dim objFields2 As ADODB.Fields
    Set objFields2 = rst2.Fields
    rst2.MoveFirst

' Muszê pomin¹æ pierwsze 6 godzin bo doba zaczyna siê od 7
    For i = 1 To 7
    rst.MoveNext
    rst2.MoveNext
    Next i
    'wypisuje prognozê i przypisuje do zmiennych bo wykorzystam je x razy
            progh7 = objFields.Item("Prognoza!" & typ_param).Value
            rst.MoveNext
            progh8 = objFields.Item("Prognoza!" & typ_param).Value
            rst.MoveNext
            progh9 = objFields.Item("Prognoza!" & typ_param).Value
            rst.MoveNext
            progh10 = objFields.Item("Prognoza!" & typ_param).Value
            rst.MoveNext
            progh11 = objFields.Item("Prognoza!" & typ_param).Value
            rst.MoveNext
            progh12 = objFields.Item("Prognoza!" & typ_param).Value
            rst.MoveNext
            progh13 = objFields.Item("Prognoza!" & typ_param).Value
            rst.MoveNext
            progh14 = objFields.Item("Prognoza!" & typ_param).Value
            rst.MoveNext
            progh15 = objFields.Item("Prognoza!" & typ_param).Value
            rst.MoveNext
            progh16 = objFields.Item("Prognoza!" & typ_param).Value
            rst.MoveNext
            progh17 = objFields.Item("Prognoza!" & typ_param).Value
            rst.MoveNext
            progh18 = objFields.Item("Prognoza!" & typ_param).Value
            rst.MoveNext
            progh19 = objFields.Item("Prognoza!" & typ_param).Value
            rst.MoveNext
            progh20 = objFields.Item("Prognoza!" & typ_param).Value
            rst.MoveNext
            progh21 = objFields.Item("Prognoza!" & typ_param).Value
            rst.MoveNext
            progh22 = objFields.Item("Prognoza!" & typ_param).Value
            rst.MoveNext
            progh23 = objFields.Item("Prognoza!" & typ_param).Value
            rst.MoveNext
            progh0 = objFields.Item("Prognoza!" & typ_param).Value
            rst.MoveNext
            progh1 = objFields.Item("Prognoza!" & typ_param).Value
            rst.MoveNext
            progh2 = objFields.Item("Prognoza!" & typ_param).Value
            rst.MoveNext
            progh3 = objFields.Item("Prognoza!" & typ_param).Value
            rst.MoveNext
            progh4 = objFields.Item("Prognoza!" & typ_param).Value
            rst.MoveNext
            progh5 = objFields.Item("Prognoza!" & typ_param).Value
            rst.MoveNext
            progh6 = objFields.Item("Prognoza!" & typ_param).Value
            rst.MoveNext

    'licze podobienstwo
                For i = 1 To startrow Step 1
                    pph7 = Abs(progh7 - objFields2.Item("Wykonanie!" & typ_param).Value) * h7
                    rst2.MoveNext
                    pph8 = Abs(progh8 - objFields2.Item("Wykonanie!" & typ_param).Value) * h8
                    rst2.MoveNext
                    pph9 = Abs(progh9 - objFields2.Item("Wykonanie!" & typ_param).Value) * h9
                    rst2.MoveNext
                    pph10 = Abs(progh10 - objFields2.Item("Wykonanie!" & typ_param).Value) * h10
                    rst2.MoveNext
                    pph11 = Abs(progh11 - objFields2.Item("Wykonanie!" & typ_param).Value) * h11
                    rst2.MoveNext
                    pph12 = Abs(progh12 - objFields2.Item("Wykonanie!" & typ_param).Value) * h12
                    rst2.MoveNext
                    pph13 = Abs(progh13 - objFields2.Item("Wykonanie!" & typ_param).Value) * h13
                    rst2.MoveNext
                    pph14 = Abs(progh14 - objFields2.Item("Wykonanie!" & typ_param).Value) * h14
                    rst2.MoveNext
                    pph15 = Abs(progh15 - objFields2.Item("Wykonanie!" & typ_param).Value) * h15
                    rst2.MoveNext
                    pph16 = Abs(progh16 - objFields2.Item("Wykonanie!" & typ_param).Value) * h16
                    rst2.MoveNext
                    pph17 = Abs(progh17 - objFields2.Item("Wykonanie!" & typ_param).Value) * h17
                    rst2.MoveNext
                    pph18 = Abs(progh18 - objFields2.Item("Wykonanie!" & typ_param).Value) * h18
                    rst2.MoveNext
                    pph19 = Abs(progh19 - objFields2.Item("Wykonanie!" & typ_param).Value) * h19
                    rst2.MoveNext
                    pph20 = Abs(progh20 - objFields2.Item("Wykonanie!" & typ_param).Value) * h20
                    rst2.MoveNext
                    pph21 = Abs(progh21 - objFields2.Item("Wykonanie!" & typ_param).Value) * h21
                    rst2.MoveNext
                    pph22 = Abs(progh22 - objFields2.Item("Wykonanie!" & typ_param).Value) * h22
                    rst2.MoveNext
                    pph23 = Abs(progh23 - objFields2.Item("Wykonanie!" & typ_param).Value) * h23
                    rst2.MoveNext
                    pph0 = Abs(progh0 - objFields2.Item("Wykonanie!" & typ_param).Value) * H0
                    rst2.MoveNext
                    pph1 = Abs(progh1 - objFields2.Item("Wykonanie!" & typ_param).Value) * h1
                    rst2.MoveNext
                    pph2 = Abs(progh2 - objFields2.Item("Wykonanie!" & typ_param).Value) * h2
                    rst2.MoveNext
                    pph3 = Abs(progh3 - objFields2.Item("Wykonanie!" & typ_param).Value) * h3
                    rst2.MoveNext
                    pph4 = Abs(progh4 - objFields2.Item("Wykonanie!" & typ_param).Value) * h4
                    rst2.MoveNext
                    pph5 = Abs(progh5 - objFields2.Item("Wykonanie!" & typ_param).Value) * h5
                    rst2.MoveNext
                    pph6 = Abs(progh6 - objFields2.Item("Wykonanie!" & typ_param).Value) * h6
                    rst2.MoveNext
                    
                    suma = pph7 + pph8 + pph9 + pph10 + pph11 + pph12 + pph13 + pph14 + pph15 + pph16 + pph17 + pph18 + pph19 + pph20 + pph21 + pph22 + pph23 + pph0 + pph1 + pph2 + pph3 + pph4 + pph5 + pph6
                    Workbooks("prognoza Gaz PM.xlsm").Sheets("podobne").Cells(1 + i, 1) = suma
                    Workbooks("prognoza Gaz PM.xlsm").Sheets("podobne").Cells(1 + i, 2) = objFields2.Item("Wykonanie!Doba").Value - 1
                Next i
            Workbooks("prognoza Gaz PM.xlsm").Sheets("podobne").Sort.SortFields.Clear
            Workbooks("prognoza Gaz PM.xlsm").Sheets("podobne").Sort.SortFields.Add Key:=Range("A2"), _
            SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
            With Workbooks("prognoza Gaz PM.xlsm").Sheets("podobne").Sort
                .SetRange Range("A1:B43")
                .Header = xlYes
                .MatchCase = False
                .Orientation = xlTopToBottom
                .SortMethod = xlPinYin
                .Apply
            End With
            
            Workbooks("prognoza Gaz PM.xlsm").Sheets("podobne").Range("A2:B8").Copy
            If OptionButton1 = True Then
            ' TEMPERATURA
            Workbooks("prognoza Gaz PM.xlsm").Sheets("Analiza - 1").Range("O11:P17").PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
            End If
                   
            If OptionButton2 = True Then
            ' ZACHMURZENIE
               Workbooks("prognoza Gaz PM.xlsm").Sheets("Analiza - 1").Range("O21:P27").PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
            End If
            
Workbooks("prognoza Gaz PM.xlsm").Sheets("Analiza - 1").Cells(1, 1).Select
Application.ScreenUpdating = True
Unload UserForm2
End Sub

Private Sub CommandButton2_Click()
    Unload UserForm2
End Sub

Private Sub CommandButton3_Click()
    CheckBox1.Value = True
    CheckBox2.Value = True
    CheckBox3.Value = True
    CheckBox4.Value = True
    CheckBox5.Value = True
    CheckBox6.Value = True
    CheckBox7.Value = True
    CheckBox8.Value = True
    CheckBox9.Value = True
    CheckBox10.Value = True
    CheckBox11.Value = True
    CheckBox12.Value = True
    CheckBox13.Value = True
    CheckBox14.Value = True
    CheckBox15.Value = True
    CheckBox16.Value = True
    CheckBox17.Value = True
    CheckBox18.Value = True
    CheckBox19.Value = True
    CheckBox20.Value = True
    CheckBox21.Value = True
    CheckBox22.Value = True
    CheckBox23.Value = True
    CheckBox24.Value = True
End Sub

Private Sub CommandButton4_Click()
    CheckBox1.Value = False
    CheckBox2.Value = False
    CheckBox3.Value = False
    CheckBox4.Value = False
    CheckBox5.Value = False
    CheckBox6.Value = False
    CheckBox7.Value = False
    CheckBox8.Value = False
    CheckBox9.Value = False
    CheckBox10.Value = False
    CheckBox11.Value = False
    CheckBox12.Value = False
    CheckBox13.Value = False
    CheckBox14.Value = False
    CheckBox15.Value = False
    CheckBox16.Value = False
    CheckBox17.Value = False
    CheckBox18.Value = False
    CheckBox19.Value = False
    CheckBox20.Value = False
    CheckBox21.Value = False
    CheckBox22.Value = False
    CheckBox23.Value = False
    CheckBox24.Value = False
End Sub

Private Sub CommandButton5_Click()
    Dim m, z As Integer
    For m = 1 To 300 Step 168
        'zaczynam od niedzieli
        
                For z = 0 To 23 Step 1
                If CheckBox25.Value = True Then
                Arkusz5.Range("i" & m + 2 + z).Value = 1
            
                Else
                Arkusz5.Range("i" & m + 2 + z).Value = 0
                End If
                Next z
       
                
             ' poniedzialek
               For z = 0 To 23 Step 1
                If CheckBox26.Value = True Then
                Arkusz5.Range("i" & m + 26 + z).Value = 1
                Else
                Arkusz5.Range("i" & m + 26 + z).Value = 0
                End If
                Next z
           
                'wtorek
                For z = 0 To 23 Step 1
                If CheckBox27.Value = True Then
                Arkusz5.Range("i" & m + 50 + z).Value = 1
                Else
                Arkusz5.Range("i" & m + 50 + z).Value = 0
                End If
                Next z
              
                'sroda
                For z = 0 To 23 Step 1
                If CheckBox28.Value = True Then
                Arkusz5.Range("i" & m + 74 + z).Value = 1
                Else
                Arkusz5.Range("i" & m + 74 + z).Value = 0
                End If
             Next z
                
                 
                'czwartek
                For z = 0 To 23 Step 1
                If CheckBox29.Value = True Then
                Arkusz5.Range("i" & m + 98 + z).Value = 1
                Else
               Arkusz5.Range("i" & m + 98 + z).Value = 0
                End If
               Next z
               
               'piatek
                For z = 0 To 23 Step 1
                If CheckBox30.Value = True Then
                Arkusz5.Range("i" & m + 122 + z).Value = 1
                Else
                Arkusz5.Range("i" & m + 122 + z).Value = 0
                End If
                Next z
               
                 
              'sobota
                For z = 0 To 23 Step 1
                If CheckBox31.Value = True Then
                Arkusz5.Range("i" & m + 146 + z).Value = 1
                Else
                Arkusz5.Range("i" & m + 146 + z).Value = 0
                End If
                Next z
              
    Next m


End Sub









