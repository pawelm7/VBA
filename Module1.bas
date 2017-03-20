Option Explicit
	'Zadeklarowanie sta³ych
    Public Const baza_gaz = "C:\Users\pmachal\Desktop\gaz.accdb"
    Public Const lokal = "\\tpe\tpce-public$\TPE\ZHP\Biuro HPW\Pawe³\Meteo\"
    Public objaccess As Object
    Public cn As ADODB.Connection
	Public rst As Recordset
    Public Const acImport = 0
	Public Const acTable = 0
    Public Obszar_pogoda as integer
    Public ID_Klient as integer
    Public userID as integer

    
    Public Const O1 = "_Gdynia_"
    Public Const O2 = "_Grudziadz_"
    Public Const O3 = "_Warszawa_"
    Public Const O4 = "_Suwalki_"
    Public Const O5 = "_Bialystok_"
    Public Const O6 = "_Olsztyn_"
    Public Const O7 = "_Koszalin_"
    Public Const O8 = "_Szczecin_"
    Public Const O9 = "_Zielona Gora_"
    Public Const O10 = "_Poznan_"
    Public Const O11 = "_Kalisz_"
    Public Const O12 = "_Lodz_"
    Public Const O13 = "_Lublin_"
    Public Const O14 = "_Rzeszow_"
    Public Const O15 = "_Czestochowa_"
    Public Const O16 = "_Katowice_"
    Public Const O17 = "_Gliwice_"
    Public Const O18 = "_Krakow_"
    Public Const O19 = "_Wroclaw_"
    Public Const O20 = "_Bielsko-Biala_"
    
'ID_Klient - PSG nienominowani
