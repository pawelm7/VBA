Option Explicit

Private Sub dzien_Change()
dzien = Start.dzien.Value
End Sub

Private Sub miesiac_Change()
miesiac = Start.miesiac.Value
End Sub

Private Sub rok_Change()
rok = Start.rok.Value
End Sub

Private Sub Userform_Initialize()
'Za³adowanie listy argumentów do wyœwietlenia
dzien.List = Array("01", "02", "03", "04", "05", "06", "07", "08", "09", "10", "11", "12", "13", "14", "15", "16", "17", "18", "19", "20", "21", "22", "23", "24", "25", "26", "27", "28", "29", "30", "31")
miesiac.List = Array("01", "02", "03", "04", "05", "06", "07", "08", "09", "10", "11", "12")
rok.List = Array("2018", "2017", "2016", "2015")
'Ustawienie wartoœci wyœwietlanych domyœlnie w okienku listy przewijania
 dzien.ListIndex = Day(Date) - 1
 miesiac.ListIndex = Month(Date) - 1
 rok.ListIndex = 1
End Sub

Private Sub Ok_Click()
    MsgBox rok & miesiac & dzien
    Unload Start
End Sub

Private Sub Anuluj_Click()
    Unload Start
End Sub


Private Sub CommandButton1_Click()
	On Error GoTo ErrHandler
    ' import prognozy pogody do bazy danych
    Dim rok2 As String
    Dim miesiac2 As String
    Dim dzien2 As String
    Dim data As Date
    Set cn = New ADODB.Connection
    Set rst = New ADODB.Recordset
    Set objaccess = CreateObject("Access.Application")
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    cn.Open "Provider=Microsoft.ACE.OLEDB.15.0; Data Source=" & baza_gaz
    objaccess.OpenCurrentDatabase baza_gaz
    data = DateSerial(rok, miesiac, dzien)

    rok = Left(data, 4)
    miesiac = Left(Right(data, 5), 2)
    dzien = Right(data, 2)
    rok2 = Left(data - 1, 4)
    miesiac2 = Left(Right(data - 1, 5), 2)
    dzien2 = Right(data - 1, 2)

    rst.Open "Select Prognoza.Data_prognozy from Prognoza where Prognoza.Data_prognozy ='" & rok & miesiac & dzien & "';", cn
    If rst.EOF = True Then
    'import prognozy pogody z pliku TXT
    'O1
    objaccess.DoCmd.TransferText , SpecificationName:="Prognoza_O1", TableName:="Prognoza_O", Filename:=lokal & "prognoza" & O1 & rok & miesiac & dzien & ".txt", HasFieldNames:=False
    objaccess.DoCmd.RunSQL "UPDATE Prognoza_O SET Prognoza_O.obszar_ID = 1;"
    objaccess.DoCmd.RunSQL "Delete Prognoza_O.Doba FROM Prognoza_O WHERE (((Prognoza_O.Doba) Is Null));"
    objaccess.DoCmd.RunSQL "UPDATE Prognoza_O SET Prognoza_O.Data_prognozy = " & rok & miesiac & dzien & ";"
    objaccess.DoCmd.RunSQL "INSERT INTO Prognoza ( Doba, Godzina, Temperatura, Naslon, Data_prognozy, Obszar_ID )" _
    & "SELECT Prognoza_O.Doba, Prognoza_O.Godzina, Prognoza_O.Temperatura, Prognoza_O.Naslon, Prognoza_O.Data_prognozy, Prognoza_O.Obszar_ID FROM Prognoza_O;"
    objaccess.DoCmd.RunSQL "DELETE Prognoza_O.* FROM Prognoza_O;"

    'O2
    objaccess.DoCmd.TransferText , SpecificationName:="Prognoza_O1", TableName:="Prognoza_O", Filename:=lokal & "prognoza" & O2 & rok & miesiac & dzien & ".txt", HasFieldNames:=False
    objaccess.DoCmd.RunSQL "UPDATE Prognoza_O SET Prognoza_O.obszar_ID = 2;"
    objaccess.DoCmd.RunSQL "Delete Prognoza_O.Doba FROM Prognoza_O WHERE (((Prognoza_O.Doba) Is Null));"
    objaccess.DoCmd.RunSQL "UPDATE Prognoza_O SET Prognoza_O.Data_prognozy = " & rok & miesiac & dzien & ";"
    objaccess.DoCmd.RunSQL "INSERT INTO Prognoza ( Doba, Godzina, Temperatura, Naslon, Data_prognozy, Obszar_ID )" _
    & "SELECT Prognoza_O.Doba, Prognoza_O.Godzina, Prognoza_O.Temperatura, Prognoza_O.Naslon, Prognoza_O.Data_prognozy, Prognoza_O.Obszar_ID FROM Prognoza_O;"
    objaccess.DoCmd.RunSQL "DELETE Prognoza_O.* FROM Prognoza_O;"

    'O3
    objaccess.DoCmd.TransferText , SpecificationName:="Prognoza_O1", TableName:="Prognoza_O", Filename:=lokal & "prognoza" & O3 & rok & miesiac & dzien & ".txt", HasFieldNames:=False
    objaccess.DoCmd.RunSQL "UPDATE Prognoza_O SET Prognoza_O.obszar_ID = 3;"
    objaccess.DoCmd.RunSQL "Delete Prognoza_O.Doba FROM Prognoza_O WHERE (((Prognoza_O.Doba) Is Null));"
    objaccess.DoCmd.RunSQL "UPDATE Prognoza_O SET Prognoza_O.Data_prognozy = " & rok & miesiac & dzien & ";"
    objaccess.DoCmd.RunSQL "INSERT INTO Prognoza ( Doba, Godzina, Temperatura, Naslon, Data_prognozy, Obszar_ID )" _
    & "SELECT Prognoza_O.Doba, Prognoza_O.Godzina, Prognoza_O.Temperatura, Prognoza_O.Naslon, Prognoza_O.Data_prognozy, Prognoza_O.Obszar_ID FROM Prognoza_O;"
    objaccess.DoCmd.RunSQL "DELETE Prognoza_O.* FROM Prognoza_O;"

    'O4
    objaccess.DoCmd.TransferText , SpecificationName:="Prognoza_O1", TableName:="Prognoza_O", Filename:=lokal & "prognoza" & O4 & rok & miesiac & dzien & ".txt", HasFieldNames:=False
    objaccess.DoCmd.RunSQL "UPDATE Prognoza_O SET Prognoza_O.obszar_ID = 4;"
    objaccess.DoCmd.RunSQL "Delete Prognoza_O.Doba FROM Prognoza_O WHERE (((Prognoza_O.Doba) Is Null));"
    objaccess.DoCmd.RunSQL "UPDATE Prognoza_O SET Prognoza_O.Data_prognozy = " & rok & miesiac & dzien & ";"
    objaccess.DoCmd.RunSQL "INSERT INTO Prognoza ( Doba, Godzina, Temperatura, Naslon, Data_prognozy, Obszar_ID )" _
    & "SELECT Prognoza_O.Doba, Prognoza_O.Godzina, Prognoza_O.Temperatura, Prognoza_O.Naslon, Prognoza_O.Data_prognozy, Prognoza_O.Obszar_ID FROM Prognoza_O;"
    objaccess.DoCmd.RunSQL "DELETE Prognoza_O.* FROM Prognoza_O;"

    'O5
    objaccess.DoCmd.TransferText , SpecificationName:="Prognoza_O1", TableName:="Prognoza_O", Filename:=lokal & "prognoza" & O5 & rok & miesiac & dzien & ".txt", HasFieldNames:=False
    objaccess.DoCmd.RunSQL "UPDATE Prognoza_O SET Prognoza_O.obszar_ID = 5;"
    objaccess.DoCmd.RunSQL "Delete Prognoza_O.Doba FROM Prognoza_O WHERE (((Prognoza_O.Doba) Is Null));"
    objaccess.DoCmd.RunSQL "UPDATE Prognoza_O SET Prognoza_O.Data_prognozy = " & rok & miesiac & dzien & ";"
    objaccess.DoCmd.RunSQL "INSERT INTO Prognoza ( Doba, Godzina, Temperatura, Naslon, Data_prognozy, Obszar_ID )" _
    & "SELECT Prognoza_O.Doba, Prognoza_O.Godzina, Prognoza_O.Temperatura, Prognoza_O.Naslon, Prognoza_O.Data_prognozy, Prognoza_O.Obszar_ID FROM Prognoza_O;"
    objaccess.DoCmd.RunSQL "DELETE Prognoza_O.* FROM Prognoza_O;"

    'O6
    objaccess.DoCmd.TransferText , SpecificationName:="Prognoza_O1", TableName:="Prognoza_O", Filename:=lokal & "prognoza" & O6 & rok & miesiac & dzien & ".txt", HasFieldNames:=False
    objaccess.DoCmd.RunSQL "UPDATE Prognoza_O SET Prognoza_O.obszar_ID = 6;"
    objaccess.DoCmd.RunSQL "Delete Prognoza_O.Doba FROM Prognoza_O WHERE (((Prognoza_O.Doba) Is Null));"
    objaccess.DoCmd.RunSQL "UPDATE Prognoza_O SET Prognoza_O.Data_prognozy = " & rok & miesiac & dzien & ";"
    objaccess.DoCmd.RunSQL "INSERT INTO Prognoza ( Doba, Godzina, Temperatura, Naslon, Data_prognozy, Obszar_ID )" _
    & "SELECT Prognoza_O.Doba, Prognoza_O.Godzina, Prognoza_O.Temperatura, Prognoza_O.Naslon, Prognoza_O.Data_prognozy, Prognoza_O.Obszar_ID FROM Prognoza_O;"
    objaccess.DoCmd.RunSQL "DELETE Prognoza_O.* FROM Prognoza_O;"

    'O7
    objaccess.DoCmd.TransferText , SpecificationName:="Prognoza_O1", TableName:="Prognoza_O", Filename:=lokal & "prognoza" & O7 & rok & miesiac & dzien & ".txt", HasFieldNames:=False
    objaccess.DoCmd.RunSQL "UPDATE Prognoza_O SET Prognoza_O.obszar_ID = 7;"
    objaccess.DoCmd.RunSQL "Delete Prognoza_O.Doba FROM Prognoza_O WHERE (((Prognoza_O.Doba) Is Null));"
    objaccess.DoCmd.RunSQL "UPDATE Prognoza_O SET Prognoza_O.Data_prognozy = " & rok & miesiac & dzien & ";"
    objaccess.DoCmd.RunSQL "INSERT INTO Prognoza ( Doba, Godzina, Temperatura, Naslon, Data_prognozy, Obszar_ID )" _
    & "SELECT Prognoza_O.Doba, Prognoza_O.Godzina, Prognoza_O.Temperatura, Prognoza_O.Naslon, Prognoza_O.Data_prognozy, Prognoza_O.Obszar_ID FROM Prognoza_O;"
    objaccess.DoCmd.RunSQL "DELETE Prognoza_O.* FROM Prognoza_O;"

    'O8
    objaccess.DoCmd.TransferText , SpecificationName:="Prognoza_O1", TableName:="Prognoza_O", Filename:=lokal & "prognoza" & O8 & rok & miesiac & dzien & ".txt", HasFieldNames:=False
    objaccess.DoCmd.RunSQL "UPDATE Prognoza_O SET Prognoza_O.obszar_ID = 8;"
    objaccess.DoCmd.RunSQL "Delete Prognoza_O.Doba FROM Prognoza_O WHERE (((Prognoza_O.Doba) Is Null));"
    objaccess.DoCmd.RunSQL "UPDATE Prognoza_O SET Prognoza_O.Data_prognozy = " & rok & miesiac & dzien & ";"
    objaccess.DoCmd.RunSQL "INSERT INTO Prognoza ( Doba, Godzina, Temperatura, Naslon, Data_prognozy, Obszar_ID )" _
    & "SELECT Prognoza_O.Doba, Prognoza_O.Godzina, Prognoza_O.Temperatura, Prognoza_O.Naslon, Prognoza_O.Data_prognozy, Prognoza_O.Obszar_ID FROM Prognoza_O;"
    objaccess.DoCmd.RunSQL "DELETE Prognoza_O.* FROM Prognoza_O;"

    'O9
    objaccess.DoCmd.TransferText , SpecificationName:="Prognoza_O1", TableName:="Prognoza_O", Filename:=lokal & "prognoza" & O9 & rok & miesiac & dzien & ".txt", HasFieldNames:=False
    objaccess.DoCmd.RunSQL "UPDATE Prognoza_O SET Prognoza_O.obszar_ID = 9;"
    objaccess.DoCmd.RunSQL "Delete Prognoza_O.Doba FROM Prognoza_O WHERE (((Prognoza_O.Doba) Is Null));"
    objaccess.DoCmd.RunSQL "UPDATE Prognoza_O SET Prognoza_O.Data_prognozy = " & rok & miesiac & dzien & ";"
    objaccess.DoCmd.RunSQL "INSERT INTO Prognoza ( Doba, Godzina, Temperatura, Naslon, Data_prognozy, Obszar_ID )" _
    & "SELECT Prognoza_O.Doba, Prognoza_O.Godzina, Prognoza_O.Temperatura, Prognoza_O.Naslon, Prognoza_O.Data_prognozy, Prognoza_O.Obszar_ID FROM Prognoza_O;"
    objaccess.DoCmd.RunSQL "DELETE Prognoza_O.* FROM Prognoza_O;"

    'O10
    objaccess.DoCmd.TransferText , SpecificationName:="Prognoza_O1", TableName:="Prognoza_O", Filename:=lokal & "prognoza" & O10 & rok & miesiac & dzien & ".txt", HasFieldNames:=False
    objaccess.DoCmd.RunSQL "UPDATE Prognoza_O SET Prognoza_O.obszar_ID = 10;"
    objaccess.DoCmd.RunSQL "Delete Prognoza_O.Doba FROM Prognoza_O WHERE (((Prognoza_O.Doba) Is Null));"
    objaccess.DoCmd.RunSQL "UPDATE Prognoza_O SET Prognoza_O.Data_prognozy = " & rok & miesiac & dzien & ";"
    objaccess.DoCmd.RunSQL "INSERT INTO Prognoza ( Doba, Godzina, Temperatura, Naslon, Data_prognozy, Obszar_ID )" _
    & "SELECT Prognoza_O.Doba, Prognoza_O.Godzina, Prognoza_O.Temperatura, Prognoza_O.Naslon, Prognoza_O.Data_prognozy, Prognoza_O.Obszar_ID FROM Prognoza_O;"
    objaccess.DoCmd.RunSQL "DELETE Prognoza_O.* FROM Prognoza_O;"

    'O11
    objaccess.DoCmd.TransferText , SpecificationName:="Prognoza_O1", TableName:="Prognoza_O", Filename:=lokal & "prognoza" & O11 & rok & miesiac & dzien & ".txt", HasFieldNames:=False
    objaccess.DoCmd.RunSQL "UPDATE Prognoza_O SET Prognoza_O.obszar_ID = 11;"
    objaccess.DoCmd.RunSQL "Delete Prognoza_O.Doba FROM Prognoza_O WHERE (((Prognoza_O.Doba) Is Null));"
    objaccess.DoCmd.RunSQL "UPDATE Prognoza_O SET Prognoza_O.Data_prognozy = " & rok & miesiac & dzien & ";"
    objaccess.DoCmd.RunSQL "INSERT INTO Prognoza ( Doba, Godzina, Temperatura, Naslon, Data_prognozy, Obszar_ID )" _
    & "SELECT Prognoza_O.Doba, Prognoza_O.Godzina, Prognoza_O.Temperatura, Prognoza_O.Naslon, Prognoza_O.Data_prognozy, Prognoza_O.Obszar_ID FROM Prognoza_O;"
    objaccess.DoCmd.RunSQL "DELETE Prognoza_O.* FROM Prognoza_O;"

    'O12
    objaccess.DoCmd.TransferText , SpecificationName:="Prognoza_O1", TableName:="Prognoza_O", Filename:=lokal & "prognoza" & O12 & rok & miesiac & dzien & ".txt", HasFieldNames:=False
    objaccess.DoCmd.RunSQL "UPDATE Prognoza_O SET Prognoza_O.obszar_ID = 12;"
    objaccess.DoCmd.RunSQL "Delete Prognoza_O.Doba FROM Prognoza_O WHERE (((Prognoza_O.Doba) Is Null));"
    objaccess.DoCmd.RunSQL "UPDATE Prognoza_O SET Prognoza_O.Data_prognozy = " & rok & miesiac & dzien & ";"
    objaccess.DoCmd.RunSQL "INSERT INTO Prognoza ( Doba, Godzina, Temperatura, Naslon, Data_prognozy, Obszar_ID )" _
    & "SELECT Prognoza_O.Doba, Prognoza_O.Godzina, Prognoza_O.Temperatura, Prognoza_O.Naslon, Prognoza_O.Data_prognozy, Prognoza_O.Obszar_ID FROM Prognoza_O;"
    objaccess.DoCmd.RunSQL "DELETE Prognoza_O.* FROM Prognoza_O;"

    'O13
    objaccess.DoCmd.TransferText , SpecificationName:="Prognoza_O1", TableName:="Prognoza_O", Filename:=lokal & "prognoza" & O13 & rok & miesiac & dzien & ".txt", HasFieldNames:=False
    objaccess.DoCmd.RunSQL "UPDATE Prognoza_O SET Prognoza_O.obszar_ID = 13;"
    objaccess.DoCmd.RunSQL "Delete Prognoza_O.Doba FROM Prognoza_O WHERE (((Prognoza_O.Doba) Is Null));"
    objaccess.DoCmd.RunSQL "UPDATE Prognoza_O SET Prognoza_O.Data_prognozy = " & rok & miesiac & dzien & ";"
    objaccess.DoCmd.RunSQL "INSERT INTO Prognoza ( Doba, Godzina, Temperatura, Naslon, Data_prognozy, Obszar_ID )" _
    & "SELECT Prognoza_O.Doba, Prognoza_O.Godzina, Prognoza_O.Temperatura, Prognoza_O.Naslon, Prognoza_O.Data_prognozy, Prognoza_O.Obszar_ID FROM Prognoza_O;"
    objaccess.DoCmd.RunSQL "DELETE Prognoza_O.* FROM Prognoza_O;"

    'O14
    objaccess.DoCmd.TransferText , SpecificationName:="Prognoza_O1", TableName:="Prognoza_O", Filename:=lokal & "prognoza" & O14 & rok & miesiac & dzien & ".txt", HasFieldNames:=False
    objaccess.DoCmd.RunSQL "UPDATE Prognoza_O SET Prognoza_O.obszar_ID = 14;"
    objaccess.DoCmd.RunSQL "Delete Prognoza_O.Doba FROM Prognoza_O WHERE (((Prognoza_O.Doba) Is Null));"
    objaccess.DoCmd.RunSQL "UPDATE Prognoza_O SET Prognoza_O.Data_prognozy = " & rok & miesiac & dzien & ";"
    objaccess.DoCmd.RunSQL "INSERT INTO Prognoza ( Doba, Godzina, Temperatura, Naslon, Data_prognozy, Obszar_ID )" _
    & "SELECT Prognoza_O.Doba, Prognoza_O.Godzina, Prognoza_O.Temperatura, Prognoza_O.Naslon, Prognoza_O.Data_prognozy, Prognoza_O.Obszar_ID FROM Prognoza_O;"
    objaccess.DoCmd.RunSQL "DELETE Prognoza_O.* FROM Prognoza_O;"

    'O15
    objaccess.DoCmd.TransferText , SpecificationName:="Prognoza_O1", TableName:="Prognoza_O", Filename:=lokal & "prognoza" & O15 & rok & miesiac & dzien & ".txt", HasFieldNames:=False
    objaccess.DoCmd.RunSQL "UPDATE Prognoza_O SET Prognoza_O.obszar_ID = 15;"
    objaccess.DoCmd.RunSQL "Delete Prognoza_O.Doba FROM Prognoza_O WHERE (((Prognoza_O.Doba) Is Null));"
    objaccess.DoCmd.RunSQL "UPDATE Prognoza_O SET Prognoza_O.Data_prognozy = " & rok & miesiac & dzien & ";"
    objaccess.DoCmd.RunSQL "INSERT INTO Prognoza ( Doba, Godzina, Temperatura, Naslon, Data_prognozy, Obszar_ID )" _
    & "SELECT Prognoza_O.Doba, Prognoza_O.Godzina, Prognoza_O.Temperatura, Prognoza_O.Naslon, Prognoza_O.Data_prognozy, Prognoza_O.Obszar_ID FROM Prognoza_O;"
    objaccess.DoCmd.RunSQL "DELETE Prognoza_O.* FROM Prognoza_O;"

    'O16
    objaccess.DoCmd.TransferText , SpecificationName:="Prognoza_O1", TableName:="Prognoza_O", Filename:=lokal & "prognoza" & O16 & rok & miesiac & dzien & ".txt", HasFieldNames:=False
    objaccess.DoCmd.RunSQL "UPDATE Prognoza_O SET Prognoza_O.obszar_ID = 16;"
    objaccess.DoCmd.RunSQL "Delete Prognoza_O.Doba FROM Prognoza_O WHERE (((Prognoza_O.Doba) Is Null));"
    objaccess.DoCmd.RunSQL "UPDATE Prognoza_O SET Prognoza_O.Data_prognozy = " & rok & miesiac & dzien & ";"
    objaccess.DoCmd.RunSQL "INSERT INTO Prognoza ( Doba, Godzina, Temperatura, Naslon, Data_prognozy, Obszar_ID )" _
    & "SELECT Prognoza_O.Doba, Prognoza_O.Godzina, Prognoza_O.Temperatura, Prognoza_O.Naslon, Prognoza_O.Data_prognozy, Prognoza_O.Obszar_ID FROM Prognoza_O;"
    objaccess.DoCmd.RunSQL "DELETE Prognoza_O.* FROM Prognoza_O;"

    'O17
    objaccess.DoCmd.TransferText , SpecificationName:="Prognoza_O1", TableName:="Prognoza_O", Filename:=lokal & "prognoza" & O17 & rok & miesiac & dzien & ".txt", HasFieldNames:=False
    objaccess.DoCmd.RunSQL "UPDATE Prognoza_O SET Prognoza_O.obszar_ID = 17;"
    objaccess.DoCmd.RunSQL "Delete Prognoza_O.Doba FROM Prognoza_O WHERE (((Prognoza_O.Doba) Is Null));"
    objaccess.DoCmd.RunSQL "UPDATE Prognoza_O SET Prognoza_O.Data_prognozy = " & rok & miesiac & dzien & ";"
    objaccess.DoCmd.RunSQL "INSERT INTO Prognoza ( Doba, Godzina, Temperatura, Naslon, Data_prognozy, Obszar_ID )" _
    & "SELECT Prognoza_O.Doba, Prognoza_O.Godzina, Prognoza_O.Temperatura, Prognoza_O.Naslon, Prognoza_O.Data_prognozy, Prognoza_O.Obszar_ID FROM Prognoza_O;"
    objaccess.DoCmd.RunSQL "DELETE Prognoza_O.* FROM Prognoza_O;"

    'O18
    objaccess.DoCmd.TransferText , SpecificationName:="Prognoza_O1", TableName:="Prognoza_O", Filename:=lokal & "prognoza" & O18 & rok & miesiac & dzien & ".txt", HasFieldNames:=False
    objaccess.DoCmd.RunSQL "UPDATE Prognoza_O SET Prognoza_O.obszar_ID = 18;"
    objaccess.DoCmd.RunSQL "Delete Prognoza_O.Doba FROM Prognoza_O WHERE (((Prognoza_O.Doba) Is Null));"
    objaccess.DoCmd.RunSQL "UPDATE Prognoza_O SET Prognoza_O.Data_prognozy = " & rok & miesiac & dzien & ";"
    objaccess.DoCmd.RunSQL "INSERT INTO Prognoza ( Doba, Godzina, Temperatura, Naslon, Data_prognozy, Obszar_ID )" _
    & "SELECT Prognoza_O.Doba, Prognoza_O.Godzina, Prognoza_O.Temperatura, Prognoza_O.Naslon, Prognoza_O.Data_prognozy, Prognoza_O.Obszar_ID FROM Prognoza_O;"
    objaccess.DoCmd.RunSQL "DELETE Prognoza_O.* FROM Prognoza_O;"

    'O19
    objaccess.DoCmd.TransferText , SpecificationName:="Prognoza_O1", TableName:="Prognoza_O", Filename:=lokal & "prognoza" & O19 & rok & miesiac & dzien & ".txt", HasFieldNames:=False
    objaccess.DoCmd.RunSQL "UPDATE Prognoza_O SET Prognoza_O.obszar_ID = 19;"
    objaccess.DoCmd.RunSQL "Delete Prognoza_O.Doba FROM Prognoza_O WHERE (((Prognoza_O.Doba) Is Null));"
    objaccess.DoCmd.RunSQL "UPDATE Prognoza_O SET Prognoza_O.Data_prognozy = " & rok & miesiac & dzien & ";"
    objaccess.DoCmd.RunSQL "INSERT INTO Prognoza ( Doba, Godzina, Temperatura, Naslon, Data_prognozy, Obszar_ID )" _
    & "SELECT Prognoza_O.Doba, Prognoza_O.Godzina, Prognoza_O.Temperatura, Prognoza_O.Naslon, Prognoza_O.Data_prognozy, Prognoza_O.Obszar_ID FROM Prognoza_O;"
    objaccess.DoCmd.RunSQL "DELETE Prognoza_O.* FROM Prognoza_O;"

    'O20
    objaccess.DoCmd.TransferText , SpecificationName:="Prognoza_O1", TableName:="Prognoza_O", Filename:=lokal & "prognoza" & O20 & rok & miesiac & dzien & ".txt", HasFieldNames:=False
    objaccess.DoCmd.RunSQL "UPDATE Prognoza_O SET Prognoza_O.obszar_ID = 20;"
    objaccess.DoCmd.RunSQL "Delete Prognoza_O.Doba FROM Prognoza_O WHERE (((Prognoza_O.Doba) Is Null));"
    objaccess.DoCmd.RunSQL "UPDATE Prognoza_O SET Prognoza_O.Data_prognozy = " & rok & miesiac & dzien & ";"
    objaccess.DoCmd.RunSQL "INSERT INTO Prognoza ( Doba, Godzina, Temperatura, Naslon, Data_prognozy, Obszar_ID )" _
    & "SELECT Prognoza_O.Doba, Prognoza_O.Godzina, Prognoza_O.Temperatura, Prognoza_O.Naslon, Prognoza_O.Data_prognozy, Prognoza_O.Obszar_ID FROM Prognoza_O;"
    objaccess.DoCmd.RunSQL "DELETE Prognoza_O.* FROM Prognoza_O;"

        '#################### Wykonanie
        'import wykonania pogody z pliku TXT
        'O1
        objaccess.DoCmd.TransferText , SpecificationName:="Prognoza_O1", TableName:="Wykonanie_O", Filename:=lokal & "wykonanie" & O1 & rok2 & miesiac2 & dzien2 & ".txt", HasFieldNames:=False
        objaccess.DoCmd.RunSQL "UPDATE Wykonanie_O SET Wykonanie_O.obszar_ID = 1;"
        objaccess.DoCmd.RunSQL "Delete Wykonanie_O.Doba FROM Wykonanie_O WHERE (((Wykonanie_O.Doba) Is Null));"
        objaccess.DoCmd.RunSQL "INSERT INTO Wykonanie ( Doba, Godzina, Temperatura, Naslon, Obszar_ID )" _
        & "SELECT Wykonanie_O.Doba, Wykonanie_O.Godzina, Wykonanie_O.Temperatura, Wykonanie_O.Naslon, Wykonanie_O.Obszar_ID FROM Wykonanie_O;"
        objaccess.DoCmd.RunSQL "DELETE Wykonanie_O.* FROM Wykonanie_O;"

        'O2
        objaccess.DoCmd.TransferText , SpecificationName:="Prognoza_O1", TableName:="Wykonanie_O", Filename:=lokal & "wykonanie" & O2 & rok2 & miesiac2 & dzien2 & ".txt", HasFieldNames:=False
        objaccess.DoCmd.RunSQL "UPDATE Wykonanie_O SET Wykonanie_O.obszar_ID = 2;"
        objaccess.DoCmd.RunSQL "Delete Wykonanie_O.Doba FROM Wykonanie_O WHERE (((Wykonanie_O.Doba) Is Null));"
        objaccess.DoCmd.RunSQL "INSERT INTO Wykonanie ( Doba, Godzina, Temperatura, Naslon, Obszar_ID )" _
        & "SELECT Wykonanie_O.Doba, Wykonanie_O.Godzina, Wykonanie_O.Temperatura, Wykonanie_O.Naslon, Wykonanie_O.Obszar_ID FROM Wykonanie_O;"
        objaccess.DoCmd.RunSQL "DELETE Wykonanie_O.* FROM Wykonanie_O;"

        'O3
        objaccess.DoCmd.TransferText , SpecificationName:="Prognoza_O1", TableName:="Wykonanie_O", Filename:=lokal & "wykonanie" & O3 & rok2 & miesiac2 & dzien2 & ".txt", HasFieldNames:=False
        objaccess.DoCmd.RunSQL "UPDATE Wykonanie_O SET Wykonanie_O.obszar_ID = 3;"
        objaccess.DoCmd.RunSQL "Delete Wykonanie_O.Doba FROM Wykonanie_O WHERE (((Wykonanie_O.Doba) Is Null));"
        objaccess.DoCmd.RunSQL "INSERT INTO Wykonanie ( Doba, Godzina, Temperatura, Naslon, Obszar_ID )" _
        & "SELECT Wykonanie_O.Doba, Wykonanie_O.Godzina, Wykonanie_O.Temperatura, Wykonanie_O.Naslon, Wykonanie_O.Obszar_ID FROM Wykonanie_O;"
        objaccess.DoCmd.RunSQL "DELETE Wykonanie_O.* FROM Wykonanie_O;"

        'O4
        objaccess.DoCmd.TransferText , SpecificationName:="Prognoza_O1", TableName:="Wykonanie_O", Filename:=lokal & "wykonanie" & O4 & rok2 & miesiac2 & dzien2 & ".txt", HasFieldNames:=False
        objaccess.DoCmd.RunSQL "UPDATE Wykonanie_O SET Wykonanie_O.obszar_ID = 4;"
        objaccess.DoCmd.RunSQL "Delete Wykonanie_O.Doba FROM Wykonanie_O WHERE (((Wykonanie_O.Doba) Is Null));"

        objaccess.DoCmd.RunSQL "INSERT INTO Wykonanie ( Doba, Godzina, Temperatura, Naslon, Obszar_ID )" _
        & "SELECT Wykonanie_O.Doba, Wykonanie_O.Godzina, Wykonanie_O.Temperatura, Wykonanie_O.Naslon, Wykonanie_O.Obszar_ID FROM Wykonanie_O;"
        objaccess.DoCmd.RunSQL "DELETE Wykonanie_O.* FROM Wykonanie_O;"

        'O5
        objaccess.DoCmd.TransferText , SpecificationName:="Prognoza_O1", TableName:="Wykonanie_O", Filename:=lokal & "wykonanie" & O5 & rok2 & miesiac2 & dzien2 & ".txt", HasFieldNames:=False
        objaccess.DoCmd.RunSQL "UPDATE Wykonanie_O SET Wykonanie_O.obszar_ID = 5;"
        objaccess.DoCmd.RunSQL "Delete Wykonanie_O.Doba FROM Wykonanie_O WHERE (((Wykonanie_O.Doba) Is Null));"

        objaccess.DoCmd.RunSQL "INSERT INTO Wykonanie ( Doba, Godzina, Temperatura, Naslon, Obszar_ID )" _
        & "SELECT Wykonanie_O.Doba, Wykonanie_O.Godzina, Wykonanie_O.Temperatura, Wykonanie_O.Naslon, Wykonanie_O.Obszar_ID FROM Wykonanie_O;"
        objaccess.DoCmd.RunSQL "DELETE Wykonanie_O.* FROM Wykonanie_O;"

        'O6
        objaccess.DoCmd.TransferText , SpecificationName:="Prognoza_O1", TableName:="Wykonanie_O", Filename:=lokal & "wykonanie" & O6 & rok2 & miesiac2 & dzien2 & ".txt", HasFieldNames:=False
        objaccess.DoCmd.RunSQL "UPDATE Wykonanie_O SET Wykonanie_O.obszar_ID = 6;"
        objaccess.DoCmd.RunSQL "Delete Wykonanie_O.Doba FROM Wykonanie_O WHERE (((Wykonanie_O.Doba) Is Null));"

        objaccess.DoCmd.RunSQL "INSERT INTO Wykonanie ( Doba, Godzina, Temperatura, Naslon, Obszar_ID )" _
        & "SELECT Wykonanie_O.Doba, Wykonanie_O.Godzina, Wykonanie_O.Temperatura, Wykonanie_O.Naslon, Wykonanie_O.Obszar_ID FROM Wykonanie_O;"
        objaccess.DoCmd.RunSQL "DELETE Wykonanie_O.* FROM Wykonanie_O;"

        'O7
        objaccess.DoCmd.TransferText , SpecificationName:="Prognoza_O1", TableName:="Wykonanie_O", Filename:=lokal & "wykonanie" & O7 & rok2 & miesiac2 & dzien2 & ".txt", HasFieldNames:=False
        objaccess.DoCmd.RunSQL "UPDATE Wykonanie_O SET Wykonanie_O.obszar_ID = 7;"
        objaccess.DoCmd.RunSQL "Delete Wykonanie_O.Doba FROM Wykonanie_O WHERE (((Wykonanie_O.Doba) Is Null));"

        objaccess.DoCmd.RunSQL "INSERT INTO Wykonanie ( Doba, Godzina, Temperatura, Naslon, Obszar_ID )" _
        & "SELECT Wykonanie_O.Doba, Wykonanie_O.Godzina, Wykonanie_O.Temperatura, Wykonanie_O.Naslon, Wykonanie_O.Obszar_ID FROM Wykonanie_O;"
        objaccess.DoCmd.RunSQL "DELETE Wykonanie_O.* FROM Wykonanie_O;"

        'O8
        objaccess.DoCmd.TransferText , SpecificationName:="Prognoza_O1", TableName:="Wykonanie_O", Filename:=lokal & "wykonanie" & O8 & rok2 & miesiac2 & dzien2 & ".txt", HasFieldNames:=False
        objaccess.DoCmd.RunSQL "UPDATE Wykonanie_O SET Wykonanie_O.obszar_ID = 8;"
        objaccess.DoCmd.RunSQL "Delete Wykonanie_O.Doba FROM Wykonanie_O WHERE (((Wykonanie_O.Doba) Is Null));"

        objaccess.DoCmd.RunSQL "INSERT INTO Wykonanie ( Doba, Godzina, Temperatura, Naslon, Obszar_ID )" _
        & "SELECT Wykonanie_O.Doba, Wykonanie_O.Godzina, Wykonanie_O.Temperatura, Wykonanie_O.Naslon, Wykonanie_O.Obszar_ID FROM Wykonanie_O;"
        objaccess.DoCmd.RunSQL "DELETE Wykonanie_O.* FROM Wykonanie_O;"

        'O9
        objaccess.DoCmd.TransferText , SpecificationName:="Prognoza_O1", TableName:="Wykonanie_O", Filename:=lokal & "wykonanie" & O9 & rok2 & miesiac2 & dzien2 & ".txt", HasFieldNames:=False
        objaccess.DoCmd.RunSQL "UPDATE Wykonanie_O SET Wykonanie_O.obszar_ID = 9;"
        objaccess.DoCmd.RunSQL "Delete Wykonanie_O.Doba FROM Wykonanie_O WHERE (((Wykonanie_O.Doba) Is Null));"

        objaccess.DoCmd.RunSQL "INSERT INTO Wykonanie ( Doba, Godzina, Temperatura, Naslon, Obszar_ID )" _
        & "SELECT Wykonanie_O.Doba, Wykonanie_O.Godzina, Wykonanie_O.Temperatura, Wykonanie_O.Naslon, Wykonanie_O.Obszar_ID FROM Wykonanie_O;"
        objaccess.DoCmd.RunSQL "DELETE Wykonanie_O.* FROM Wykonanie_O;"

        'O10
        objaccess.DoCmd.TransferText , SpecificationName:="Prognoza_O1", TableName:="Wykonanie_O", Filename:=lokal & "wykonanie" & O10 & rok2 & miesiac2 & dzien2 & ".txt", HasFieldNames:=False
        objaccess.DoCmd.RunSQL "UPDATE Wykonanie_O SET Wykonanie_O.obszar_ID = 10;"
        objaccess.DoCmd.RunSQL "Delete Wykonanie_O.Doba FROM Wykonanie_O WHERE (((Wykonanie_O.Doba) Is Null));"

        objaccess.DoCmd.RunSQL "INSERT INTO Wykonanie ( Doba, Godzina, Temperatura, Naslon, Obszar_ID )" _
        & "SELECT Wykonanie_O.Doba, Wykonanie_O.Godzina, Wykonanie_O.Temperatura, Wykonanie_O.Naslon, Wykonanie_O.Obszar_ID FROM Wykonanie_O;"
        objaccess.DoCmd.RunSQL "DELETE Wykonanie_O.* FROM Wykonanie_O;"

        'O11
        objaccess.DoCmd.TransferText , SpecificationName:="Prognoza_O1", TableName:="Wykonanie_O", Filename:=lokal & "wykonanie" & O11 & rok2 & miesiac2 & dzien2 & ".txt", HasFieldNames:=False
        objaccess.DoCmd.RunSQL "UPDATE Wykonanie_O SET Wykonanie_O.obszar_ID = 11;"
        objaccess.DoCmd.RunSQL "Delete Wykonanie_O.Doba FROM Wykonanie_O WHERE (((Wykonanie_O.Doba) Is Null));"

        objaccess.DoCmd.RunSQL "INSERT INTO Wykonanie ( Doba, Godzina, Temperatura, Naslon, Obszar_ID )" _
        & "SELECT Wykonanie_O.Doba, Wykonanie_O.Godzina, Wykonanie_O.Temperatura, Wykonanie_O.Naslon, Wykonanie_O.Obszar_ID FROM Wykonanie_O;"
        objaccess.DoCmd.RunSQL "DELETE Wykonanie_O.* FROM Wykonanie_O;"

        'O12
        objaccess.DoCmd.TransferText , SpecificationName:="Prognoza_O1", TableName:="Wykonanie_O", Filename:=lokal & "wykonanie" & O12 & rok2 & miesiac2 & dzien2 & ".txt", HasFieldNames:=False
        objaccess.DoCmd.RunSQL "UPDATE Wykonanie_O SET Wykonanie_O.obszar_ID = 12;"
        objaccess.DoCmd.RunSQL "Delete Wykonanie_O.Doba FROM Wykonanie_O WHERE (((Wykonanie_O.Doba) Is Null));"

        objaccess.DoCmd.RunSQL "INSERT INTO Wykonanie ( Doba, Godzina, Temperatura, Naslon, Obszar_ID )" _
        & "SELECT Wykonanie_O.Doba, Wykonanie_O.Godzina, Wykonanie_O.Temperatura, Wykonanie_O.Naslon, Wykonanie_O.Obszar_ID FROM Wykonanie_O;"
        objaccess.DoCmd.RunSQL "DELETE Wykonanie_O.* FROM Wykonanie_O;"

        'O13
        objaccess.DoCmd.TransferText , SpecificationName:="Prognoza_O1", TableName:="Wykonanie_O", Filename:=lokal & "wykonanie" & O13 & rok2 & miesiac2 & dzien2 & ".txt", HasFieldNames:=False
        objaccess.DoCmd.RunSQL "UPDATE Wykonanie_O SET Wykonanie_O.obszar_ID = 13;"
        objaccess.DoCmd.RunSQL "Delete Wykonanie_O.Doba FROM Wykonanie_O WHERE (((Wykonanie_O.Doba) Is Null));"

        objaccess.DoCmd.RunSQL "INSERT INTO Wykonanie ( Doba, Godzina, Temperatura, Naslon, Obszar_ID )" _
        & "SELECT Wykonanie_O.Doba, Wykonanie_O.Godzina, Wykonanie_O.Temperatura, Wykonanie_O.Naslon, Wykonanie_O.Obszar_ID FROM Wykonanie_O;"
        objaccess.DoCmd.RunSQL "DELETE Wykonanie_O.* FROM Wykonanie_O;"

        'O14
        objaccess.DoCmd.TransferText , SpecificationName:="Prognoza_O1", TableName:="Wykonanie_O", Filename:=lokal & "wykonanie" & O14 & rok2 & miesiac2 & dzien2 & ".txt", HasFieldNames:=False
        objaccess.DoCmd.RunSQL "UPDATE Wykonanie_O SET Wykonanie_O.obszar_ID = 14;"
        objaccess.DoCmd.RunSQL "Delete Wykonanie_O.Doba FROM Wykonanie_O WHERE (((Wykonanie_O.Doba) Is Null));"

        objaccess.DoCmd.RunSQL "INSERT INTO Wykonanie ( Doba, Godzina, Temperatura, Naslon, Obszar_ID )" _
        & "SELECT Wykonanie_O.Doba, Wykonanie_O.Godzina, Wykonanie_O.Temperatura, Wykonanie_O.Naslon, Wykonanie_O.Obszar_ID FROM Wykonanie_O;"
        objaccess.DoCmd.RunSQL "DELETE Wykonanie_O.* FROM Wykonanie_O;"

        'O15
        objaccess.DoCmd.TransferText , SpecificationName:="Prognoza_O1", TableName:="Wykonanie_O", Filename:=lokal & "wykonanie" & O15 & rok2 & miesiac2 & dzien2 & ".txt", HasFieldNames:=False
        objaccess.DoCmd.RunSQL "UPDATE Wykonanie_O SET Wykonanie_O.obszar_ID = 15;"
        objaccess.DoCmd.RunSQL "Delete Wykonanie_O.Doba FROM Wykonanie_O WHERE (((Wykonanie_O.Doba) Is Null));"

        objaccess.DoCmd.RunSQL "INSERT INTO Wykonanie ( Doba, Godzina, Temperatura, Naslon, Obszar_ID )" _
        & "SELECT Wykonanie_O.Doba, Wykonanie_O.Godzina, Wykonanie_O.Temperatura, Wykonanie_O.Naslon, Wykonanie_O.Obszar_ID FROM Wykonanie_O;"
        objaccess.DoCmd.RunSQL "DELETE Wykonanie_O.* FROM Wykonanie_O;"

        'O16
        objaccess.DoCmd.TransferText , SpecificationName:="Prognoza_O1", TableName:="Wykonanie_O", Filename:=lokal & "wykonanie" & O16 & rok2 & miesiac2 & dzien2 & ".txt", HasFieldNames:=False
        objaccess.DoCmd.RunSQL "UPDATE Wykonanie_O SET Wykonanie_O.obszar_ID = 16;"
        objaccess.DoCmd.RunSQL "Delete Wykonanie_O.Doba FROM Wykonanie_O WHERE (((Wykonanie_O.Doba) Is Null));"

        objaccess.DoCmd.RunSQL "INSERT INTO Wykonanie ( Doba, Godzina, Temperatura, Naslon, Obszar_ID )" _
        & "SELECT Wykonanie_O.Doba, Wykonanie_O.Godzina, Wykonanie_O.Temperatura, Wykonanie_O.Naslon, Wykonanie_O.Obszar_ID FROM Wykonanie_O;"
        objaccess.DoCmd.RunSQL "DELETE Wykonanie_O.* FROM Wykonanie_O;"

        'O17
        objaccess.DoCmd.TransferText , SpecificationName:="Prognoza_O1", TableName:="Wykonanie_O", Filename:=lokal & "wykonanie" & O17 & rok2 & miesiac2 & dzien2 & ".txt", HasFieldNames:=False
        objaccess.DoCmd.RunSQL "UPDATE Wykonanie_O SET Wykonanie_O.obszar_ID = 17;"
        objaccess.DoCmd.RunSQL "Delete Wykonanie_O.Doba FROM Wykonanie_O WHERE (((Wykonanie_O.Doba) Is Null));"

        objaccess.DoCmd.RunSQL "INSERT INTO Wykonanie ( Doba, Godzina, Temperatura, Naslon, Obszar_ID )" _
        & "SELECT Wykonanie_O.Doba, Wykonanie_O.Godzina, Wykonanie_O.Temperatura, Wykonanie_O.Naslon, Wykonanie_O.Obszar_ID FROM Wykonanie_O;"
        objaccess.DoCmd.RunSQL "DELETE Wykonanie_O.* FROM Wykonanie_O;"

        'O18
        objaccess.DoCmd.TransferText , SpecificationName:="Prognoza_O1", TableName:="Wykonanie_O", Filename:=lokal & "wykonanie" & O18 & rok2 & miesiac2 & dzien2 & ".txt", HasFieldNames:=False
        objaccess.DoCmd.RunSQL "UPDATE Wykonanie_O SET Wykonanie_O.obszar_ID = 18;"
        objaccess.DoCmd.RunSQL "Delete Wykonanie_O.Doba FROM Wykonanie_O WHERE (((Wykonanie_O.Doba) Is Null));"

        objaccess.DoCmd.RunSQL "INSERT INTO Wykonanie ( Doba, Godzina, Temperatura, Naslon, Obszar_ID )" _
        & "SELECT Wykonanie_O.Doba, Wykonanie_O.Godzina, Wykonanie_O.Temperatura, Wykonanie_O.Naslon, Wykonanie_O.Obszar_ID FROM Wykonanie_O;"
        objaccess.DoCmd.RunSQL "DELETE Wykonanie_O.* FROM Wykonanie_O;"

        'O19
        objaccess.DoCmd.TransferText , SpecificationName:="Prognoza_O1", TableName:="Wykonanie_O", Filename:=lokal & "wykonanie" & O19 & rok2 & miesiac2 & dzien2 & ".txt", HasFieldNames:=False
        objaccess.DoCmd.RunSQL "UPDATE Wykonanie_O SET Wykonanie_O.obszar_ID = 19;"
        objaccess.DoCmd.RunSQL "Delete Wykonanie_O.Doba FROM Wykonanie_O WHERE (((Wykonanie_O.Doba) Is Null));"

        objaccess.DoCmd.RunSQL "INSERT INTO Wykonanie ( Doba, Godzina, Temperatura, Naslon, Obszar_ID )" _
        & "SELECT Wykonanie_O.Doba, Wykonanie_O.Godzina, Wykonanie_O.Temperatura, Wykonanie_O.Naslon, Wykonanie_O.Obszar_ID FROM Wykonanie_O;"
        objaccess.DoCmd.RunSQL "DELETE Wykonanie_O.* FROM Wykonanie_O;"

        'O20
        objaccess.DoCmd.TransferText , SpecificationName:="Prognoza_O1", TableName:="Wykonanie_O", Filename:=lokal & "wykonanie" & O20 & rok2 & miesiac2 & dzien2 & ".txt", HasFieldNames:=False
        objaccess.DoCmd.RunSQL "UPDATE Wykonanie_O SET Wykonanie_O.obszar_ID = 20;"
        objaccess.DoCmd.RunSQL "Delete Wykonanie_O.Doba FROM Wykonanie_O WHERE (((Wykonanie_O.Doba) Is Null));"

        objaccess.DoCmd.RunSQL "INSERT INTO Wykonanie ( Doba, Godzina, Temperatura, Naslon, Obszar_ID )" _
        & "SELECT Wykonanie_O.Doba, Wykonanie_O.Godzina, Wykonanie_O.Temperatura, Wykonanie_O.Naslon, Wykonanie_O.Obszar_ID FROM Wykonanie_O;"
        objaccess.DoCmd.RunSQL "DELETE Wykonanie_O.* FROM Wykonanie_O;"
        
    objaccess.DoCmd.DeleteObject acTable, "prognoza" & O1 & rok & miesiac & dzien & "_B³êdyImportu"
    objaccess.DoCmd.DeleteObject acTable, "prognoza" & O2 & rok & miesiac & dzien & "_B³êdyImportu"
    objaccess.DoCmd.DeleteObject acTable, "prognoza" & O3 & rok & miesiac & dzien & "_B³êdyImportu"
    objaccess.DoCmd.DeleteObject acTable, "prognoza" & O4 & rok & miesiac & dzien & "_B³êdyImportu"
    objaccess.DoCmd.DeleteObject acTable, "prognoza" & O5 & rok & miesiac & dzien & "_B³êdyImportu"
    objaccess.DoCmd.DeleteObject acTable, "prognoza" & O6 & rok & miesiac & dzien & "_B³êdyImportu"
    objaccess.DoCmd.DeleteObject acTable, "prognoza" & O7 & rok & miesiac & dzien & "_B³êdyImportu"
    objaccess.DoCmd.DeleteObject acTable, "prognoza" & O8 & rok & miesiac & dzien & "_B³êdyImportu"
    objaccess.DoCmd.DeleteObject acTable, "prognoza" & O9 & rok & miesiac & dzien & "_B³êdyImportu"
    objaccess.DoCmd.DeleteObject acTable, "prognoza" & O10 & rok & miesiac & dzien & "_B³êdyImportu"
    objaccess.DoCmd.DeleteObject acTable, "prognoza" & O11 & rok & miesiac & dzien & "_B³êdyImportu"
    objaccess.DoCmd.DeleteObject acTable, "prognoza" & O12 & rok & miesiac & dzien & "_B³êdyImportu"
    objaccess.DoCmd.DeleteObject acTable, "prognoza" & O13 & rok & miesiac & dzien & "_B³êdyImportu"
    objaccess.DoCmd.DeleteObject acTable, "prognoza" & O14 & rok & miesiac & dzien & "_B³êdyImportu"
    objaccess.DoCmd.DeleteObject acTable, "prognoza" & O15 & rok & miesiac & dzien & "_B³êdyImportu"
    objaccess.DoCmd.DeleteObject acTable, "prognoza" & O16 & rok & miesiac & dzien & "_B³êdyImportu"
    objaccess.DoCmd.DeleteObject acTable, "prognoza" & O17 & rok & miesiac & dzien & "_B³êdyImportu"
    objaccess.DoCmd.DeleteObject acTable, "prognoza" & O18 & rok & miesiac & dzien & "_B³êdyImportu"
    objaccess.DoCmd.DeleteObject acTable, "prognoza" & O19 & rok & miesiac & dzien & "_B³êdyImportu"
    objaccess.DoCmd.DeleteObject acTable, "prognoza" & O20 & rok & miesiac & dzien & "_B³êdyImportu"
    
    objaccess.DoCmd.DeleteObject acTable, "wykonanie" & O1 & rok2 & miesiac2 & dzien2 & "_B³êdyImportu"
    objaccess.DoCmd.DeleteObject acTable, "wykonanie" & O2 & rok2 & miesiac2 & dzien2 & "_B³êdyImportu"
    objaccess.DoCmd.DeleteObject acTable, "wykonanie" & O3 & rok2 & miesiac2 & dzien2 & "_B³êdyImportu"
    objaccess.DoCmd.DeleteObject acTable, "wykonanie" & O4 & rok2 & miesiac2 & dzien2 & "_B³êdyImportu"
    objaccess.DoCmd.DeleteObject acTable, "wykonanie" & O5 & rok2 & miesiac2 & dzien2 & "_B³êdyImportu"
    objaccess.DoCmd.DeleteObject acTable, "wykonanie" & O6 & rok2 & miesiac2 & dzien2 & "_B³êdyImportu"
    objaccess.DoCmd.DeleteObject acTable, "wykonanie" & O7 & rok2 & miesiac2 & dzien2 & "_B³êdyImportu"
    objaccess.DoCmd.DeleteObject acTable, "wykonanie" & O8 & rok2 & miesiac2 & dzien2 & "_B³êdyImportu"
    objaccess.DoCmd.DeleteObject acTable, "wykonanie" & O9 & rok2 & miesiac2 & dzien2 & "_B³êdyImportu"
    objaccess.DoCmd.DeleteObject acTable, "wykonanie" & O10 & rok2 & miesiac2 & dzien2 & "_B³êdyImportu"
    objaccess.DoCmd.DeleteObject acTable, "wykonanie" & O11 & rok2 & miesiac2 & dzien2 & "_B³êdyImportu"
    objaccess.DoCmd.DeleteObject acTable, "wykonanie" & O12 & rok2 & miesiac2 & dzien2 & "_B³êdyImportu"
    objaccess.DoCmd.DeleteObject acTable, "wykonanie" & O13 & rok2 & miesiac2 & dzien2 & "_B³êdyImportu"
    objaccess.DoCmd.DeleteObject acTable, "wykonanie" & O14 & rok2 & miesiac2 & dzien2 & "_B³êdyImportu"
    objaccess.DoCmd.DeleteObject acTable, "wykonanie" & O15 & rok2 & miesiac2 & dzien2 & "_B³êdyImportu"
    objaccess.DoCmd.DeleteObject acTable, "wykonanie" & O16 & rok2 & miesiac2 & dzien2 & "_B³êdyImportu"
    objaccess.DoCmd.DeleteObject acTable, "wykonanie" & O17 & rok2 & miesiac2 & dzien2 & "_B³êdyImportu"
    objaccess.DoCmd.DeleteObject acTable, "wykonanie" & O18 & rok2 & miesiac2 & dzien2 & "_B³êdyImportu"
    objaccess.DoCmd.DeleteObject acTable, "wykonanie" & O19 & rok2 & miesiac2 & dzien2 & "_B³êdyImportu"
    objaccess.DoCmd.DeleteObject acTable, "wykonanie" & O20 & rok2 & miesiac2 & dzien2 & "_B³êdyImportu"
    
    MsgBox "Import prognozy pogody z dnia " & data & " i wykonania z dnia " & data - 1 & " zakonczyl sie powodzeniem!"
    Else
    MsgBox "W bazie jest ju¿ prognoza pogody! Proszê wybraæ inn¹ datê"
    End If
    Set rst = Nothing
    cn.Close
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
	
	ErrHandler:
    'If Err.Number = 91 Then
     '   MsgBox "Nie ma nastêpnego arkusza."
    'Else
    'inny b³¹d
        MsgBox "Wyst¹pi³ b³¹d nr: " & Err.Number & Chr(10) & Err.Description
    'End If
    Err.Clear

End Sub




