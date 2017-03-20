Option Explicit

Private Sub CommandButton1_Click()
        Dim pom As Double
        Dim memory As Double
        'deklaracja zmiennych do skalowania
            Dim zakres, zakres1 As Range
            Dim max, min, max1, min1 As Integer

            If TextBox1.Value = "" Then
            MsgBox "Wpisz wartoœæ wspó³czynnika", vbInformation, "Informacja"
            TextBox1.Value = 0
            Else
            End If

        pom = TextBox1.Value


    'godzina 22
    If CheckBox25.Value = True Then
        If OptionButton1 = True Then
           Arkusz2.Range("c50").Value = Arkusz2.Range("c50").Value + pom
       End If
       If OptionButton2 = True Then
            Arkusz2.Range("c50").Value = Arkusz2.Range("c50").Value * pom
        End If
        If OptionButton3 = True Then
            Arkusz2.Range("c50").Value = (Arkusz2.Range("c50").Value * (pom / 100) + Arkusz2.Range("c50").Value)
        End If
    End If

    'godzina 23
    If CheckBox26.Value = True Then
        If OptionButton1 = True Then
           Arkusz2.Range("d50").Value = Arkusz2.Range("d50").Value + pom
       End If
       If OptionButton2 = True Then
            Arkusz2.Range("d50").Value = Arkusz2.Range("d50").Value * pom
        End If
        If OptionButton3 = True Then
            Arkusz2.Range("d50").Value = (Arkusz2.Range("d50").Value * (pom / 100) + Arkusz2.Range("d50").Value)
        End If
    End If

    'godzina 24
    If CheckBox27.Value = True Then
        If OptionButton1 = True Then
           Arkusz2.Range("e50").Value = Arkusz2.Range("e50").Value + pom
       End If
       If OptionButton2 = True Then
            Arkusz2.Range("e50").Value = Arkusz2.Range("e50").Value * pom
        End If
        If OptionButton3 = True Then
            Arkusz2.Range("e50").Value = (Arkusz2.Range("e50").Value * (pom / 100) + Arkusz2.Range("e50").Value)
        End If
    End If

    'godzina 1
    If CheckBox1.Value = True Then
        If OptionButton1 = True Then
           Arkusz2.Range("f50").Value = Arkusz2.Range("f50").Value + pom
       End If
       If OptionButton2 = True Then
            Arkusz2.Range("f50").Value = Arkusz2.Range("f50").Value * pom
        End If
        If OptionButton3 = True Then
            Arkusz2.Range("f50").Value = (Arkusz2.Range("f50").Value * (pom / 100) + Arkusz2.Range("f50").Value)
        End If
    End If


    'godzina 2
    If CheckBox2.Value = True Then
        If OptionButton1 = True Then
           Arkusz2.Range("g50").Value = Arkusz2.Range("g50").Value + pom
        
       End If
       If OptionButton2 = True Then
            Arkusz2.Range("g50").Value = Arkusz2.Range("g50").Value * pom
        End If
         If OptionButton3 = True Then
            Arkusz2.Range("g50").Value = (Arkusz2.Range("g50").Value * (pom / 100) + Arkusz2.Range("g50").Value)
        End If
    End If

    'godzina 3
    If CheckBox3.Value = True Then
        If OptionButton1 = True Then
           Arkusz2.Range("h50").Value = Arkusz2.Range("h50").Value + pom
       End If
       If OptionButton2 = True Then
            Arkusz2.Range("h50").Value = Arkusz2.Range("h50").Value * pom
        End If
         If OptionButton3 = True Then
            Arkusz2.Range("h50").Value = (Arkusz2.Range("h50").Value * (pom / 100) + Arkusz2.Range("h50").Value)
        End If
    End If

    'godzina 4
    If CheckBox4.Value = True Then
        If OptionButton1 = True Then
           Arkusz2.Range("i50").Value = Arkusz2.Range("i50").Value + pom
       End If
       If OptionButton2 = True Then
            Arkusz2.Range("i50").Value = Arkusz2.Range("i50").Value * pom
        End If
         If OptionButton3 = True Then
            Arkusz2.Range("i50").Value = (Arkusz2.Range("i50").Value * (pom / 100) + Arkusz2.Range("i50").Value)
        End If
    End If

    'godzina 5
    If CheckBox5.Value = True Then
        If OptionButton1 = True Then
           Arkusz2.Range("j50").Value = Arkusz2.Range("j50").Value + pom
       End If
       If OptionButton2 = True Then
            Arkusz2.Range("j50").Value = Arkusz2.Range("j50").Value * pom
        End If
         If OptionButton3 = True Then
            Arkusz2.Range("j50").Value = (Arkusz2.Range("j50").Value * (pom / 100) + Arkusz2.Range("j50").Value)
        End If
     End If

    'godzina 6
    If CheckBox6.Value = True Then
        If OptionButton1 = True Then
           Arkusz2.Range("k50").Value = Arkusz2.Range("k50").Value + pom
       End If
       If OptionButton2 = True Then
            Arkusz2.Range("k50").Value = Arkusz2.Range("k50").Value * pom
        End If
         If OptionButton3 = True Then
            Arkusz2.Range("k50").Value = (Arkusz2.Range("k50").Value * (pom / 100) + Arkusz2.Range("k50").Value)
        End If
    End If

    'godzina 7
    If CheckBox7.Value = True Then
        If OptionButton1 = True Then
           Arkusz2.Range("l50").Value = Arkusz2.Range("l50").Value + pom
       End If
       If OptionButton2 = True Then
            Arkusz2.Range("l50").Value = Arkusz2.Range("l50").Value * pom
        End If
         If OptionButton3 = True Then
            Arkusz2.Range("l50").Value = (Arkusz2.Range("l50").Value * (pom / 100) + Arkusz2.Range("l50").Value)
        End If
    End If

    'godzina 8
    If CheckBox8.Value = True Then
        If OptionButton1 = True Then
           Arkusz2.Range("m50").Value = Arkusz2.Range("m50").Value + pom
       End If
       If OptionButton2 = True Then
            Arkusz2.Range("m50").Value = Arkusz2.Range("m50").Value * pom
        End If
         If OptionButton3 = True Then
            Arkusz2.Range("m50").Value = (Arkusz2.Range("m50").Value * (pom / 100) + Arkusz2.Range("m50").Value)
        End If
    End If

    'godzina 9
    If CheckBox9.Value = True Then
        If OptionButton1 = True Then
           Arkusz2.Range("n50").Value = Arkusz2.Range("n50").Value + pom
       End If
       If OptionButton2 = True Then
            Arkusz2.Range("n50").Value = Arkusz2.Range("n50").Value * pom
        End If
         If OptionButton3 = True Then
            Arkusz2.Range("n50").Value = (Arkusz2.Range("n50").Value * (pom / 100) + Arkusz2.Range("n50").Value)
        End If
    End If

    'godzina 10
    If CheckBox10.Value = True Then
        If OptionButton1 = True Then
           Arkusz2.Range("o50").Value = Arkusz2.Range("o50").Value + pom
       End If
       If OptionButton2 = True Then
            Arkusz2.Range("o50").Value = Arkusz2.Range("o50").Value * pom
        End If
         If OptionButton3 = True Then
            Arkusz2.Range("o50").Value = (Arkusz2.Range("o50").Value * (pom / 100) + Arkusz2.Range("o50").Value)
        End If
    End If

    'godzina 11
    If CheckBox11.Value = True Then
        If OptionButton1 = True Then
           Arkusz2.Range("p50").Value = Arkusz2.Range("p50").Value + pom
       End If
       If OptionButton2 = True Then
            Arkusz2.Range("p50").Value = Arkusz2.Range("p50").Value * pom
        End If
         If OptionButton3 = True Then
            Arkusz2.Range("p50").Value = (Arkusz2.Range("p50").Value * (pom / 100) + Arkusz2.Range("p50").Value)
        End If
    End If

    'godzina 12
    If CheckBox12.Value = True Then
        If OptionButton1 = True Then
           Arkusz2.Range("q50").Value = Arkusz2.Range("q50").Value + pom
       End If
       If OptionButton2 = True Then
            Arkusz2.Range("q50").Value = Arkusz2.Range("q50").Value * pom
        End If
         If OptionButton3 = True Then
            Arkusz2.Range("q50").Value = (Arkusz2.Range("q50").Value * (pom / 100) + Arkusz2.Range("q50").Value)
        End If
    End If

    'godzina 13
    If CheckBox13.Value = True Then
        If OptionButton1 = True Then
           Arkusz2.Range("r50").Value = Arkusz2.Range("r50").Value + pom
       End If
       If OptionButton2 = True Then
            Arkusz2.Range("r50").Value = Arkusz2.Range("r50").Value * pom
        End If
         If OptionButton3 = True Then
            Arkusz2.Range("r50").Value = (Arkusz2.Range("r50").Value * (pom / 100) + Arkusz2.Range("r50").Value)
        End If
    End If

    'godzina 14
    If CheckBox14.Value = True Then
        If OptionButton1 = True Then
           Arkusz2.Range("s50").Value = Arkusz2.Range("s50").Value + pom
       End If
       If OptionButton2 = True Then
            Arkusz2.Range("s50").Value = Arkusz2.Range("s50").Value * pom
        End If
         If OptionButton3 = True Then
            Arkusz2.Range("s50").Value = (Arkusz2.Range("s50").Value * (pom / 100) + Arkusz2.Range("s50").Value)
        End If
    End If

    'godzina 15
    If CheckBox15.Value = True Then
        If OptionButton1 = True Then
           Arkusz2.Range("t50").Value = Arkusz2.Range("t50").Value + pom
       End If
       If OptionButton2 = True Then
            Arkusz2.Range("t50").Value = Arkusz2.Range("t50").Value * pom
        End If
         If OptionButton3 = True Then
            Arkusz2.Range("t50").Value = (Arkusz2.Range("t50").Value * (pom / 100) + Arkusz2.Range("t50").Value)
        End If
    End If

    'godzina 16
    If CheckBox16.Value = True Then
        If OptionButton1 = True Then
           Arkusz2.Range("u50").Value = Arkusz2.Range("u50").Value + pom
       End If
       If OptionButton2 = True Then
            Arkusz2.Range("u50").Value = Arkusz2.Range("u50").Value * pom
        End If
         If OptionButton3 = True Then
            Arkusz2.Range("u50").Value = (Arkusz2.Range("u50").Value * (pom / 100) + Arkusz2.Range("u50").Value)
        End If
    End If

    'godzina 17
    If CheckBox17.Value = True Then
        If OptionButton1 = True Then
           Arkusz2.Range("v50").Value = Arkusz2.Range("v50").Value + pom
       End If
       If OptionButton2 = True Then
            Arkusz2.Range("v50").Value = Arkusz2.Range("v50").Value * pom
        End If
         If OptionButton3 = True Then
            Arkusz2.Range("v50").Value = (Arkusz2.Range("v50").Value * (pom / 100) + Arkusz2.Range("v50").Value)
        End If
    End If

    'godzina 18
    If CheckBox18.Value = True Then
        If OptionButton1 = True Then
           Arkusz2.Range("w50").Value = Arkusz2.Range("w50").Value + pom
       End If
       If OptionButton2 = True Then
            Arkusz2.Range("w50").Value = Arkusz2.Range("w50").Value * pom
        End If
         If OptionButton3 = True Then
            Arkusz2.Range("w50").Value = (Arkusz2.Range("w50").Value * (pom / 100) + Arkusz2.Range("w50").Value)
        End If
    End If

    'godzina 19
    If CheckBox19.Value = True Then
        If OptionButton1 = True Then
           Arkusz2.Range("x50").Value = Arkusz2.Range("x50").Value + pom
       End If
       If OptionButton2 = True Then
            Arkusz2.Range("x50").Value = Arkusz2.Range("x50").Value * pom
        End If
         If OptionButton3 = True Then
            Arkusz2.Range("x50").Value = (Arkusz2.Range("x50").Value * (pom / 100) + Arkusz2.Range("x50").Value)
        End If
    End If

    'godzina 20
    If CheckBox20.Value = True Then
        If OptionButton1 = True Then
           Arkusz2.Range("y50").Value = Arkusz2.Range("y50").Value + pom
       End If
       If OptionButton2 = True Then
            Arkusz2.Range("y50").Value = Arkusz2.Range("y50").Value * pom
        End If
         If OptionButton3 = True Then
            Arkusz2.Range("y50").Value = (Arkusz2.Range("y50").Value * (pom / 100) + Arkusz2.Range("y50").Value)
        End If
    End If

    'godzina 21
    If CheckBox21.Value = True Then
        If OptionButton1 = True Then
           Arkusz2.Range("z50").Value = Arkusz2.Range("z50").Value + pom
       End If
       If OptionButton2 = True Then
            Arkusz2.Range("z50").Value = Arkusz2.Range("z50").Value * pom
        End If
         If OptionButton3 = True Then
            Arkusz2.Range("z50").Value = (Arkusz2.Range("z50").Value * (pom / 100) + Arkusz2.Range("z50").Value)
        End If
    End If

    'godzina 22
    If CheckBox22.Value = True Then
        If OptionButton1 = True Then
           Arkusz2.Range("aa50").Value = Arkusz2.Range("aa50").Value + pom
       End If
       If OptionButton2 = True Then
            Arkusz2.Range("aa50").Value = Arkusz2.Range("aa50").Value * pom
        End If
         If OptionButton3 = True Then
            Arkusz2.Range("aa50").Value = (Arkusz2.Range("aa50").Value * (pom / 100) + Arkusz2.Range("aa50").Value)
        End If
    End If

    'godzina 23
    If CheckBox23.Value = True Then
        If OptionButton1 = True Then
           Arkusz2.Range("ab50").Value = Arkusz2.Range("ab50").Value + pom
       End If
       If OptionButton2 = True Then
            Arkusz2.Range("ab50").Value = Arkusz2.Range("ab50").Value * pom
        End If
         If OptionButton3 = True Then
            Arkusz2.Range("ab50").Value = (Arkusz2.Range("ab50").Value * (pom / 100) + Arkusz2.Range("ab50").Value)
        End If
    End If

    'godzina 24
    If CheckBox24.Value = True Then
        If OptionButton1 = True Then
           Arkusz2.Range("ac50").Value = Arkusz2.Range("ac50").Value + pom
       End If
       If OptionButton2 = True Then
            Arkusz2.Range("ac50").Value = Arkusz2.Range("ac50").Value * pom
        End If
         If OptionButton3 = True Then
            Arkusz2.Range("ac50").Value = (Arkusz2.Range("ac50").Value * (pom / 100) + Arkusz2.Range("ac50").Value)
        End If
    End If


      Application.ScreenUpdating = False
      Arkusz2.Range("c50:ac50").Copy
      Arkusz4.Range("b9:ab9").PasteSpecial xlPasteValues
      Application.ScreenUpdating = True
      
      'wykres1
      Set zakres = Arkusz4.Range("b5:ab9")
         min = Application.WorksheetFunction.min(zakres)
         max = Application.WorksheetFunction.max(zakres)
      
      Arkusz2.ChartObjects("Wykres 1").Activate
        ActiveChart.Axes(xlValue).Select
        'Application.CutCopyMode = False
        With ActiveChart.Axes(xlValue)
        .MinimumScale = Round((min - 20), 0)
        .MaximumScale = Round((max + 20), 0)
        End With
        
End Sub

Private Sub CommandButton2_Click()
    UserForm1.Hide
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
    CheckBox25.Value = True
    CheckBox26.Value = True
    CheckBox27.Value = True
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
    CheckBox25.Value = False
    CheckBox26.Value = False
    CheckBox27.Value = False
End Sub


