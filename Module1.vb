Imports System.Data
Imports System.Console
Imports System.Data.SqlClient
Imports System.Data.Common

Imports System
Imports System.IO
Imports System.Collections

'***********************************************************************
'*		Programmbeschreibung ConsoleApplication2
'*				
'*		Das Programm liest vom SQL Server 192.168.15.115 (Webserver) aus der Datenbank tbLeitsystem die Tabellen
'*			- tbMitarbeiter
'*			- tbAuftrag
'*			- tbEinsatz
'*			- tbSoll
'*			- tbTermin
'*			- tbKalender
'*		verarbeitet/berechnet die Daten nach Vorgabe und schreibt die erhaltenen Resultate in die Tabelle tbEinsatzWeb. 
'*		Bei jedem Run wird der Inhalt der Tabelle tbEinsatzWeb gelöscht. Die tbEinsatzWeb dient als Datenbasis für die
'*		Anzeige es Dietsche Leitsystem auf den Monitoren.
'*
'***********************************************************************

Module Module1

    Public SQLSERVER As String
    Public RueckgabeDatum As DatumWert

    Public F01 As Integer
    Public F02 As Integer
    Public F03 As Integer
    Public F04 As Integer
    Public F05 As Integer
    Public F06 As Integer
    Public F07 As Integer
    Public F08 As Integer
    Public F09 As Integer
	Public F10 As Integer

	Public F01_Doppelt As Integer
	Public F02_Doppelt As Integer
	Public F03_Doppelt As Integer
	Public F04_Doppelt As Integer
	Public F05_Doppelt As Integer
	Public F06_Doppelt As Integer
	Public F07_Doppelt As Integer
	Public F08_Doppelt As Integer
	Public F09_Doppelt As Integer


    Public Suchdatum As Date

    Public FeldId As Integer

    Public Moddate As Date = Now()

	Sub Main()

		Readtxt()

		Console.WriteLine("------------------------Programm Start " & Now())

		RunLoeschen()

		RueckgabeDatum = KalenderLesen(0)
		Suchdatum = RueckgabeDatum.D1
		run()

		RueckgabeDatum = KalenderLesen(1)
		Suchdatum = RueckgabeDatum.D1
		run()

		RueckgabeDatum = KalenderLesen(2)
		Suchdatum = RueckgabeDatum.D1
		run()

		RueckgabeDatum = KalenderLesen(3)
		Suchdatum = RueckgabeDatum.D1
		run()

		Console.WriteLine("------------------------Programm Ende " & Now())

	End Sub

    Private Function Readtxt()

        Dim objReader As New StreamReader("c:\Leitsystem\sql.txt")
        Dim sLine As String = ""
        Dim arrText As New ArrayList()

        Do
            sLine = objReader.ReadLine()
            If Not sLine Is Nothing Then
                arrText.Add(sLine)
            End If
        Loop Until sLine Is Nothing
        objReader.Close()

        SQLSERVER = arrText(0)

    End Function

    Private Sub RunLoeschen()

        Dim con As New SqlConnection
        Dim cmd As New SqlCommand
        Dim reader As DbDataReader

        con.ConnectionString = SQLSERVER

        Dim w As Boolean = True

        Try

            con.Open()
            cmd.Connection = con

            'cmd = New SqlCommand("delete from tbEinsatzWeb where Archiv = '" & w & "'", con)
            'cmd = New SqlCommand("delete from tbEinsatzWeb where Archiv = 0", con)
            cmd = New SqlCommand("delete from tbEinsatzWeb", con)
            reader = cmd.ExecuteReader

            reader.Close()

        Catch ex As Exception

            Console.WriteLine("Invalid" & ex.Message)

        Finally
            con.Close()
        End Try

    End Sub

    Private Sub run()

        Dim Total1 As Integer = 0
		Dim Total2 As Integer = 0

        '(1)------------------------------------------------------------------
        'Mitarbeiter Total angestellt im Betrieb

        FeldId = 1

		F01 = get_tbMitarbeiter_Total("SG")
		F02 = get_tbMitarbeiter_Total("ZH")
		F03 = get_tbMitarbeiter_Total("ZG")
		F04 = get_tbMitarbeiter_Total("FR")
		F05 = get_tbMitarbeiter_Total("TG")

        F10 = F01 + F02 + F03 + F04 + F05

		tbEinsatzWebEinfuegen()


        '(2)------------------------------------------------------------------
        'Mitarbeiter Verfügbar

        FeldId = 2

		F01 = get_tbMitarbeiter_Total("SG") - get_tbTermin_Total("F", "SG") - get_tbTermin_Total("M", "SG") - get_tbTermin_Total("Z", "SG") - get_tbTermin_Total("U", "SG") - get_tbTermin_Total("K", "SG") - get_tbTermin_Total("D", "SG")
		F02 = get_tbMitarbeiter_Total("ZH") - get_tbTermin_Total("F", "ZH") - get_tbTermin_Total("M", "ZH") - get_tbTermin_Total("Z", "ZH") - get_tbTermin_Total("U", "ZH") - get_tbTermin_Total("K", "ZH") - get_tbTermin_Total("D", "ZH")
		F03 = get_tbMitarbeiter_Total("ZG") - get_tbTermin_Total("F", "ZG") - get_tbTermin_Total("M", "ZG") - get_tbTermin_Total("Z", "ZG") - get_tbTermin_Total("U", "ZG") - get_tbTermin_Total("K", "ZG") - get_tbTermin_Total("D", "ZG")
		F04 = get_tbMitarbeiter_Total("FR") - get_tbTermin_Total("F", "FR") - get_tbTermin_Total("M", "FR") - get_tbTermin_Total("Z", "FR") - get_tbTermin_Total("U", "FR") - get_tbTermin_Total("K", "FR") - get_tbTermin_Total("D", "FR")
		F05 = get_tbMitarbeiter_Total("TG") - get_tbTermin_Total("F", "TG") - get_tbTermin_Total("M", "TG") - get_tbTermin_Total("Z", "TG") - get_tbTermin_Total("U", "TG") - get_tbTermin_Total("K", "TG") - get_tbTermin_Total("D", "TG")

        F10 = F01 + F02 + F03 + F04 + F05

		tbEinsatzWebEinfuegen()


        '(3)------------------------------------------------------------------
        'Mitarbeiter SOLL

        FeldId = 3

		F01 = get_tbSoll_Total(RueckgabeDatum.D3, "SG")
		F02 = get_tbSoll_Total(RueckgabeDatum.D3, "ZH")
		F03 = get_tbSoll_Total(RueckgabeDatum.D3, "ZG")
		F04 = get_tbSoll_Total(RueckgabeDatum.D3, "FR")
		F05 = get_tbSoll_Total(RueckgabeDatum.D3, "TG")

        F10 = F01 + F02 + F03 + F04 + F05

		tbEinsatzWebEinfuegen()


        '(4)------------------------------------------------------------------
        'Mitarbeiter IST

		FeldId = 4

		'Anpassung 11.01.2018, RMätzler: Ab 2018 kann es sein, dass Mitarbeiter an einem Tag 2 Einsätze haben.
		'Für das Leitsystem ist aber nur ein Einsatz zu zählen

		F01 = get_tbEinsatz_Total("SG")
		F01_Doppelt = get_tbEinsatz_Total_Doppelt("SG")
		F01 = F01 - F01_Doppelt

		F02 = get_tbEinsatz_Total("ZH")
		F02_Doppelt = get_tbEinsatz_Total_Doppelt("ZH")
		F02 = F02 - F02_Doppelt

		F03 = get_tbEinsatz_Total("ZG")
		F03_Doppelt = get_tbEinsatz_Total_Doppelt("ZG")
		F03 = F03 - F03_Doppelt

		F04 = get_tbEinsatz_Total("FR")
		F04_Doppelt = get_tbEinsatz_Total_Doppelt("FR")
		F04 = F04 - F04_Doppelt

		F05 = get_tbEinsatz_Total("TG")
		F05_Doppelt = get_tbEinsatz_Total_Doppelt("TG")
		F05 = F05 - F05_Doppelt


		F10 = F01 + F02 + F03 + F04 + F05

        tbEinsatzWebEinfuegen()

        '(5)------------------------------------------------------------------
        'Mitarbeiter Arbeitsfähig aber Ohne Einsatz

        FeldId = 5

        F01 = getMitarbeiter_OhneEinsatz("SG")
		F02 = getMitarbeiter_OhneEinsatz("ZH")
		F03 = getMitarbeiter_OhneEinsatz("ZG")
		F04 = getMitarbeiter_OhneEinsatz("FR")
        F05 = getMitarbeiter_OhneEinsatz("TG")

        F10 = F01 + F02 + F03 + F04 + F05

        tbEinsatzWebEinfuegen()


        '(6)------------------------------------------------------------------
        'Termin F/M/Z

        FeldId = 6

        F01 = get_tbTermin("F", "SG") + get_tbTermin("M", "SG") + get_tbTermin("Z", "SG")
        F02 = get_tbTermin("F", "ZH") + get_tbTermin("M", "ZH") + get_tbTermin("Z", "ZH")
        F03 = get_tbTermin("F", "ZG") + get_tbTermin("M", "ZG") + get_tbTermin("Z", "ZG")
        F04 = get_tbTermin("F", "FR") + get_tbTermin("M", "FR") + get_tbTermin("Z", "FR")
        F05 = get_tbTermin("F", "TG") + get_tbTermin("M", "TG") + get_tbTermin("Z", "TG")

        F10 = F01 + F02 + F03 + F04 + F05

        tbEinsatzWebEinfuegen()


        '(7)------------------------------------------------------------------
        'Termin U/K

        FeldId = 7

        F01 = get_tbTermin("U", "SG") + get_tbTermin("K", "SG")
        F02 = get_tbTermin("U", "ZH") + get_tbTermin("K", "ZH")
        F03 = get_tbTermin("U", "ZG") + get_tbTermin("K", "ZG")
        F04 = get_tbTermin("U", "FR") + get_tbTermin("K", "FR")
        F05 = get_tbTermin("U", "TG") + get_tbTermin("K", "TG")

        F10 = F01 + F02 + F03 + F04 + F05

        tbEinsatzWebEinfuegen()

        '(8)------------------------------------------------------------------
        'Termin D

        FeldId = 8

        F01 = get_tbTermin("D", "SG")
        F02 = get_tbTermin("D", "ZH")
        F03 = get_tbTermin("D", "ZG")
        F04 = get_tbTermin("D", "FR")
        F05 = get_tbTermin("D", "TG")

        F10 = F01 + F02 + F03 + F04 + F05

        tbEinsatzWebEinfuegen()


        '(9)------------------------------------------------------------------
        'Mitarbeiter verfügbar nach Sparte

        FeldId = 9

        Total1 = get_tbMitarbeiter("INN", "SG") + get_tbMitarbeiter("INN", "ZH") + get_tbMitarbeiter("INN", "ZG") + get_tbMitarbeiter("INN", "FR") + get_tbMitarbeiter("INN", "TG")
        Total2 = get_tbMitarbeiter_nichtVerfuegbar("INN", "SG") + get_tbMitarbeiter_nichtVerfuegbar("INN", "ZH") + get_tbMitarbeiter_nichtVerfuegbar("INN", "ZG") + get_tbMitarbeiter_nichtVerfuegbar("INN", "FR") + get_tbMitarbeiter_nichtVerfuegbar("INN", "TG")
        F01 = Total1 - Total2
        Total1 = Total2 = 0

        Total1 = get_tbMitarbeiter("LAD", "SG") + get_tbMitarbeiter("LAD", "ZH") + get_tbMitarbeiter("LAD", "ZG") + get_tbMitarbeiter("LAD", "FR") + get_tbMitarbeiter("LAD", "TG")
        Total2 = get_tbMitarbeiter_nichtVerfuegbar("LAD", "SG") + get_tbMitarbeiter_nichtVerfuegbar("LAD", "ZH") + get_tbMitarbeiter_nichtVerfuegbar("LAD", "ZG") + get_tbMitarbeiter_nichtVerfuegbar("LAD", "FR") + get_tbMitarbeiter_nichtVerfuegbar("LAD", "TG")
        F02 = Total1 - Total2
        Total1 = Total2 = 0


        Total1 = get_tbMitarbeiter("KÜC", "SG") + get_tbMitarbeiter("KÜC", "ZH") + get_tbMitarbeiter("KÜC", "ZG") + get_tbMitarbeiter("KÜC", "FR") + get_tbMitarbeiter("KÜC", "TG")
        Total2 = get_tbMitarbeiter_nichtVerfuegbar("KÜC", "SG") + get_tbMitarbeiter_nichtVerfuegbar("KÜC", "ZH") + get_tbMitarbeiter_nichtVerfuegbar("KÜC", "ZG") + get_tbMitarbeiter_nichtVerfuegbar("KÜC", "FR") + get_tbMitarbeiter_nichtVerfuegbar("KÜC", "TG")
        F03 = Total1 - Total2
        Total1 = Total2 = 0


        Total1 = get_tbMitarbeiter("T-B", "SG") + get_tbMitarbeiter("T-B", "ZH") + get_tbMitarbeiter("T-B", "ZG") + get_tbMitarbeiter("T-B", "FR") + get_tbMitarbeiter("T-B", "TG")
        Total2 = get_tbMitarbeiter_nichtVerfuegbar("T-B", "SG") + get_tbMitarbeiter_nichtVerfuegbar("T-B", "ZH") + get_tbMitarbeiter_nichtVerfuegbar("T-B", "ZG") + get_tbMitarbeiter_nichtVerfuegbar("T-B", "FR") + get_tbMitarbeiter_nichtVerfuegbar("T-B", "TG")
        F04 = Total1 - Total2
        Total1 = Total2 = 0


        Total1 = get_tbMitarbeiter("GLA", "SG") + get_tbMitarbeiter("GLA", "ZH") + get_tbMitarbeiter("GLA", "ZG") + get_tbMitarbeiter("GLA", "FR") + get_tbMitarbeiter("GLA", "TG")
        Total2 = get_tbMitarbeiter_nichtVerfuegbar("GLA", "SG") + get_tbMitarbeiter_nichtVerfuegbar("GLA", "ZH") + get_tbMitarbeiter_nichtVerfuegbar("GLA", "ZG") + get_tbMitarbeiter_nichtVerfuegbar("GLA", "FR") + get_tbMitarbeiter_nichtVerfuegbar("GLA", "TG")
        F05 = Total1 - Total2
        Total1 = Total2 = 0

        Total1 = get_tbMitarbeiter("FEN", "SG") + get_tbMitarbeiter("FEN", "ZH") + get_tbMitarbeiter("FEN", "ZG") + get_tbMitarbeiter("FEN", "FR") + get_tbMitarbeiter("FEN", "TG")
        Total2 = get_tbMitarbeiter_nichtVerfuegbar("FEN", "SG") + get_tbMitarbeiter_nichtVerfuegbar("FEN", "ZH") + get_tbMitarbeiter_nichtVerfuegbar("FEN", "ZG") + get_tbMitarbeiter_nichtVerfuegbar("FEN", "FR") + get_tbMitarbeiter_nichtVerfuegbar("FEN", "TG")
        F06 = Total1 - Total2
        Total1 = Total2 = 0

        Total1 = get_tbMitarbeiter("MES", "SG") + get_tbMitarbeiter("MES", "ZH") + get_tbMitarbeiter("MES", "ZG") + get_tbMitarbeiter("MES", "FR") + get_tbMitarbeiter("MES", "TG")
        Total2 = get_tbMitarbeiter_nichtVerfuegbar("MES", "SG") + get_tbMitarbeiter_nichtVerfuegbar("MES", "ZH") + get_tbMitarbeiter_nichtVerfuegbar("MES", "ZG") + get_tbMitarbeiter_nichtVerfuegbar("MES", "FR") + get_tbMitarbeiter_nichtVerfuegbar("MES", "TG")
        F07 = Total1 - Total2
        Total1 = Total2 = 0


        Total1 = get_tbMitarbeiter("MAL", "SG") + get_tbMitarbeiter("MAL", "ZH") + get_tbMitarbeiter("MAL", "ZG") + get_tbMitarbeiter("MAL", "FR") + get_tbMitarbeiter("MAL", "TG")
        Total2 = get_tbMitarbeiter_nichtVerfuegbar("MAL", "SG") + get_tbMitarbeiter_nichtVerfuegbar("MAL", "ZH") + get_tbMitarbeiter_nichtVerfuegbar("MAL", "ZG") + get_tbMitarbeiter_nichtVerfuegbar("MAL", "FR") + get_tbMitarbeiter_nichtVerfuegbar("MAL", "TG")
        F08 = Total1 - Total2
        Total1 = Total2 = 0


        Total1 = get_tbMitarbeiter("STA", "SG") + get_tbMitarbeiter("STA", "ZH") + get_tbMitarbeiter("STA", "ZG") + get_tbMitarbeiter("STA", "FR") + get_tbMitarbeiter("STA", "TG")
        Total2 = get_tbMitarbeiter_nichtVerfuegbar("STA", "SG") + get_tbMitarbeiter_nichtVerfuegbar("STA", "ZH") + get_tbMitarbeiter_nichtVerfuegbar("STA", "ZG") + get_tbMitarbeiter_nichtVerfuegbar("STA", "FR") + get_tbMitarbeiter_nichtVerfuegbar("STA", "TG")
        F09 = Total1 - Total2
        Total1 = Total2 = 0


        F10 = F01 + F02 + F03 + F04 + F05 + F06 + F07 + F08 + F09

        tbEinsatzWebEinfuegen()

        '(10)------------------------------------------------------------------
        'Mitarbeiter SOLL nach Sparte

        FeldId = 10

        F01 = get_tbSoll(RueckgabeDatum.D3, "INN", "SG") + get_tbSoll(RueckgabeDatum.D3, "INN", "ZH") + get_tbSoll(RueckgabeDatum.D3, "INN", "ZG") + get_tbSoll(RueckgabeDatum.D3, "INN", "FR") + get_tbSoll(RueckgabeDatum.D3, "INN", "TG")
        F02 = get_tbSoll(RueckgabeDatum.D3, "LAD", "SG") + get_tbSoll(RueckgabeDatum.D3, "LAD", "ZH") + get_tbSoll(RueckgabeDatum.D3, "LAD", "ZG") + get_tbSoll(RueckgabeDatum.D3, "LAD", "FR") + get_tbSoll(RueckgabeDatum.D3, "LAD", "TG")
        F03 = get_tbSoll(RueckgabeDatum.D3, "KÜC", "SG") + get_tbSoll(RueckgabeDatum.D3, "KÜC", "ZH") + get_tbSoll(RueckgabeDatum.D3, "KÜC", "ZG") + get_tbSoll(RueckgabeDatum.D3, "KÜC", "FR") + get_tbSoll(RueckgabeDatum.D3, "KÜC", "TG")
        F04 = get_tbSoll(RueckgabeDatum.D3, "T-B", "SG") + get_tbSoll(RueckgabeDatum.D3, "T-B", "ZH") + get_tbSoll(RueckgabeDatum.D3, "T-B", "ZG") + get_tbSoll(RueckgabeDatum.D3, "T-B", "FR") + get_tbSoll(RueckgabeDatum.D3, "T-B", "TG")
        F05 = get_tbSoll(RueckgabeDatum.D3, "GLA", "SG") + get_tbSoll(RueckgabeDatum.D3, "GLA", "ZH") + get_tbSoll(RueckgabeDatum.D3, "GLA", "ZG") + get_tbSoll(RueckgabeDatum.D3, "GLA", "FR") + get_tbSoll(RueckgabeDatum.D3, "GLA", "TG")
        F06 = get_tbSoll(RueckgabeDatum.D3, "FEN", "SG") + get_tbSoll(RueckgabeDatum.D3, "FEN", "ZH") + get_tbSoll(RueckgabeDatum.D3, "FEN", "ZG") + get_tbSoll(RueckgabeDatum.D3, "FEN", "FR") + get_tbSoll(RueckgabeDatum.D3, "FEN", "TG")
        F07 = get_tbSoll(RueckgabeDatum.D3, "MES", "SG") + get_tbSoll(RueckgabeDatum.D3, "MES", "ZH") + get_tbSoll(RueckgabeDatum.D3, "MES", "ZG") + get_tbSoll(RueckgabeDatum.D3, "MES", "FR") + get_tbSoll(RueckgabeDatum.D3, "MES", "TG")
        F08 = get_tbSoll(RueckgabeDatum.D3, "MAL", "SG") + get_tbSoll(RueckgabeDatum.D3, "MAL", "ZH") + get_tbSoll(RueckgabeDatum.D3, "MAL", "ZG") + get_tbSoll(RueckgabeDatum.D3, "MAL", "FR") + get_tbSoll(RueckgabeDatum.D3, "MAL", "TG")
        F09 = get_tbSoll(RueckgabeDatum.D3, "STA", "SG") + get_tbSoll(RueckgabeDatum.D3, "STA", "ZH") + get_tbSoll(RueckgabeDatum.D3, "STA", "ZG") + get_tbSoll(RueckgabeDatum.D3, "STA", "FR") + get_tbSoll(RueckgabeDatum.D3, "STA", "TG")

        F10 = F01 + F02 + F03 + F04 + F05 + F06 + F07 + F08 + F09

        tbEinsatzWebEinfuegen()

        '(11)------------------------------------------------------------------
        'Mitarbeiter IST nach Sparte

        FeldId = 11

		'Anpassung 11.01.2018, RMätzler: Ab 2018 kann es sein, dass Mitarbeiter an einem Tag 2 Einsätze haben.
		'Für das Leitsystem ist aber nur ein Einsatz zu zählen


		F01 = get_tbEinsatz_mit_MA("INN")
		F01_Doppelt = get_tbEinsatz_mit_MA_Doppelt("INN")
		F01 = F01 - F01_Doppelt

		F02 = get_tbEinsatz_mit_MA("LAD")
		F02_Doppelt = get_tbEinsatz_mit_MA_Doppelt("LAD")
		F02 = F02 - F02_Doppelt

		F03 = get_tbEinsatz_mit_MA("KÜC")
		F03_Doppelt = get_tbEinsatz_mit_MA_Doppelt("KÜC")
		F03 = F03 - F03_Doppelt

		F04 = get_tbEinsatz_mit_MA("T-B")
		F04_Doppelt = get_tbEinsatz_mit_MA_Doppelt("T-B")
		F04 = F04 - F04_Doppelt

		F05 = get_tbEinsatz_mit_MA("GLA")
		F05_Doppelt = get_tbEinsatz_mit_MA_Doppelt("GLA")
		F05 = F05 - F05_Doppelt

		F06 = get_tbEinsatz_mit_MA("FEN")
		F06_Doppelt = get_tbEinsatz_mit_MA_Doppelt("FEN")
		F06 = F06 - F06_Doppelt

		F07 = get_tbEinsatz_mit_MA("MES")
		F07_Doppelt = get_tbEinsatz_mit_MA_Doppelt("MES")
		F07 = F07 - F07_Doppelt

		F08 = get_tbEinsatz_mit_MA("MAL")
		F08_Doppelt = get_tbEinsatz_mit_MA_Doppelt("MAL")
		F08 = F08 - F08_Doppelt

		F09 = get_tbEinsatz_mit_MA("STA")
		F09_Doppelt = get_tbEinsatz_mit_MA_Doppelt("STA")
		F09 = F09 - F09_Doppelt


        F10 = F01 + F02 + F03 + F04 + F05 + F06 + F07 + F08 + F09

        tbEinsatzWebEinfuegen()

        '------------------------------------------------------------------
        'Einsätze ohne Mitarbeiter Zugeteilt

        FeldId = 12

        F01 = get_Einsatz_Ohne_MA("INN")
        F02 = get_Einsatz_Ohne_MA("LAD")
        F03 = get_Einsatz_Ohne_MA("KÜC")
        F04 = get_Einsatz_Ohne_MA("T-B")
        F05 = get_Einsatz_Ohne_MA("GLA")
        F06 = get_Einsatz_Ohne_MA("FEN")
        F07 = get_Einsatz_Ohne_MA("MES")
        F08 = get_Einsatz_Ohne_MA("MAL")
        F09 = get_Einsatz_Ohne_MA("STA")

        F10 = F01 + F02 + F03 + F04 + F05 + F06 + F07 + F08 + F09

        tbEinsatzWebEinfuegen()

    End Sub

    Private Function get_tbMitarbeiter(strZuteilungCode, strStandort)

        Dim con As New SqlConnection
        Dim cmd As New SqlCommand
        Dim count As Integer = 0

        con.ConnectionString = SQLSERVER

        Try

            con.Open()
            cmd.Connection = con

            'cmd = New SqlCommand("SELECT Count(*) FROM tbMitarbeiter where Eintritt <= '" & Suchdatum & "' and ZuteilungCode = '" & strZuteilungCode & "' and Standort = '" & strStandort & "'", con)
            cmd = New SqlCommand("SELECT Count(*) FROM tbMitarbeiter where CheckDatum = '" & Suchdatum & "' and ZuteilungCode = '" & strZuteilungCode & "' and Standort = '" & strStandort & "'", con)

            count = cmd.ExecuteScalar()


        Catch ex As Exception

            Console.WriteLine("Invalid" & ex.Message)

        Finally
            con.Close()
        End Try

        Return count

    End Function


    Public Function get_tbMitarbeiter_nichtVerfuegbar(strZuteilungCode, strStandort)


        Dim con As New SqlConnection
        Dim cmd As New SqlCommand
        Dim count As Integer = 0

        con.ConnectionString = SQLSERVER

        Try

            con.Open()
            cmd.Connection = con

            'cmd = New SqlCommand("SELECT Count(*) FROM tbMitarbeiter INNER JOIN tbTermin ON tbMitarbeiter.AdressNummer = tbTermin.Adressnummer where tbTermin.Datum = '" & Suchdatum & "' and tbMitarbeiter.ZuteilungCode = '" & strZuteilungCode & "' and tbMitarbeiter.Standort = '" & strStandort & "'", con)

            cmd = New SqlCommand("SELECT Count(*) FROM tbMitarbeiter INNER JOIN tbTermin ON tbMitarbeiter.AdressNummer = tbTermin.Adressnummer where tbTermin.Datum = '" & Suchdatum & "'" _
                                 & "and tbMitarbeiter.ZuteilungCode = '" & strZuteilungCode & "'" _
                                 & "and tbMitarbeiter.CheckDatum = '" & Suchdatum & "'" _
                                 & "and tbMitarbeiter.Standort = '" & strStandort & "'", con)

            count = cmd.ExecuteScalar()

        Catch ex As Exception

            Console.WriteLine("Invalid" & ex.Message)

        Finally
            con.Close()
        End Try

        Return count

    End Function


    Private Sub tbEinsatzWebEinfuegen()

        Dim con As New SqlConnection
        Dim cmd As New SqlCommand

        Try

            con.ConnectionString = SQLSERVER
            con.Open()
            cmd.Connection = con

            cmd = New SqlCommand("insert into tbEinsatzWeb ([Datum], [KW], [FeldId], [F01], [F02], [F03], [F04], [F05], [F06], [F07], [F08], [F09], [F10], [Moddate]) " _
                                 & " values ('" & RueckgabeDatum.D1 & "'," & RueckgabeDatum.D3 & ",'" & FeldId & "'" _
                                 & "," & F01 & "" _
                                 & "," & F02 & "" _
                                 & "," & F03 & "" _
                                 & "," & F04 & "" _
                                 & "," & F05 & "" _
                                 & "," & F06 & "" _
                                 & "," & F07 & "" _
                                 & "," & F08 & "" _
                                 & "," & F09 & "" _
                                 & "," & F10 & ", '" & Moddate & "')", con)


			cmd.ExecuteNonQuery()

        Catch ex As Exception

            Console.WriteLine("Invalid" & ex.Message)
        Finally
            con.Close()
        End Try

        reset()

    End Sub

    Public Function get_tbSoll(intKW, strZuteilungCode, strStandort)

        Dim con As New SqlConnection
        Dim cmd As New SqlCommand
        Dim Summe As Integer = 0

        con.ConnectionString = SQLSERVER

        Try

            con.Open()
            cmd.Connection = con

			cmd = New SqlCommand("SELECT Sum(Wert) FROM tbSoll where Jahr = '" & RueckgabeDatum.D4 & "' and ZuteilungCode = '" & strZuteilungCode & "' and Standort = '" & strStandort & "' and KW = " & intKW & "", con)

            If IsDBNull(cmd.ExecuteScalar) <> True Then
                Summe = cmd.ExecuteScalar()
            End If

        Catch ex As Exception

            Console.WriteLine("Invalid" & ex.Message)

        Finally
            con.Close()
        End Try

        Return Summe

    End Function

    Public Function get_tbTermin(strArt, strStandort)

        Dim con As New SqlConnection
        Dim cmd As New SqlCommand
        Dim Count As Integer = 0

        con.ConnectionString = SQLSERVER

        Try

            con.Open()
            cmd.Connection = con

            'cmd = New SqlCommand("SELECT Count(*) FROM tbTermin INNER JOIN tbMitarbeiter ON tbTermin.AdressNummer = tbMitarbeiter.Adressnummer" _
            '       & " where tbTermin.Datum = '" & Suchdatum & "' and tbTermin.Art = '" & strArt & "' and tbMitarbeiter.Standort = '" & strStandort & "'", con)

            cmd = New SqlCommand("SELECT Count(*) FROM tbTermin INNER JOIN tbMitarbeiter ON tbTermin.AdressNummer = tbMitarbeiter.Adressnummer" _
                    & " where tbTermin.Datum = '" & Suchdatum & "' and tbTermin.Art = '" & strArt & "' and tbMitarbeiter.Standort = '" & strStandort & "' and tbMitarbeiter.CheckDatum = '" & Suchdatum & "'", con)

            Count = cmd.ExecuteScalar()

        Catch ex As Exception

            Console.WriteLine("Invalid" & ex.Message)

        Finally
            con.Close()
        End Try

        Return Count

    End Function

    Public Function get_tbMitarbeiter_Total(strStandort)

        Dim con As New SqlConnection
        Dim cmd As New SqlCommand
        Dim Count As Integer = 0

        con.ConnectionString = SQLSERVER

        Try

            con.Open()
            cmd.Connection = con

            'cmd = New SqlCommand("SELECT Count(*) FROM tbMitarbeiter where Eintritt <= '" & Suchdatum & "' and Standort = '" & strStandort & "'", con)
            cmd = New SqlCommand("SELECT Count(*) FROM tbMitarbeiter where CheckDatum = '" & Suchdatum & "' and Standort = '" & strStandort & "'", con)

            Count = cmd.ExecuteScalar()

        Catch ex As Exception

            Console.WriteLine("Invalid" & ex.Message)

        Finally
            con.Close()
        End Try

        Return Count

    End Function

    Public Function get_tbTermin_Total(strArt, strStandort)

        Dim con As New SqlConnection
        Dim cmd As New SqlCommand
        Dim Count As Integer = 0

        con.ConnectionString = SQLSERVER

        Try

            con.Open()
            cmd.Connection = con

            'cmd = New SqlCommand("SELECT Count(*) FROM tbTermin INNER JOIN tbMitarbeiter ON tbTermin.AdressNummer = tbMitarbeiter.Adressnummer" _
            '    & " where tbTermin.Datum = '" & Suchdatum & "' and tbTermin.Art = '" & strArt & "' and tbMitarbeiter.Standort = '" & strStandort & "'", con)

            cmd = New SqlCommand("SELECT Count(*) FROM tbTermin INNER JOIN tbMitarbeiter ON tbTermin.AdressNummer = tbMitarbeiter.Adressnummer" _
                & " where tbTermin.Datum = '" & Suchdatum & "' and tbTermin.Art = '" & strArt & "' and tbMitarbeiter.Standort = '" & strStandort & "' and tbMitarbeiter.CheckDatum = '" & Suchdatum & "'", con)


            Count = cmd.ExecuteScalar()

        Catch ex As Exception

            Console.WriteLine("Invalid" & ex.Message)

        Finally
            con.Close()
        End Try

        Return Count

    End Function

    Public Function get_tbSoll_Total(intKW, strStandort)

        Dim con As New SqlConnection
        Dim cmd As New SqlCommand
        Dim Summe As Integer = 0

        con.ConnectionString = SQLSERVER

        Try

            con.Open()
            cmd.Connection = con

			cmd = New SqlCommand("SELECT Sum(Wert) FROM tbSoll where Jahr = '" & RueckgabeDatum.D4 & "' and Standort = '" & strStandort & "' and KW = " & intKW & "", con)

            If IsDBNull(cmd.ExecuteScalar) = False Then
                Summe = cmd.ExecuteScalar()
            End If



        Catch ex As Exception

            Console.WriteLine("Invalid" & ex.Message)

        Finally
            con.Close()
        End Try

        Return Summe

    End Function

    Public Function get_tbEinsatz_Total(strStandort)

        Dim con As New SqlConnection
        Dim cmd As New SqlCommand
        Dim Count As Integer = 0

        con.ConnectionString = SQLSERVER

        Try

            con.Open()
            cmd.Connection = con

            'cmd = New SqlCommand("SELECT Count(*) FROM tbEinsatz INNER JOIN tbMitarbeiter ON tbEinsatz.AdressNummer = tbMitarbeiter.Adressnummer" _
            '    & " where tbEinsatz.Datum = '" & Suchdatum & "' and tbMitarbeiter.Standort = '" & strStandort & "'", con)

			cmd = New SqlCommand("SELECT Count(*) FROM tbEinsatz INNER JOIN tbMitarbeiter ON tbEinsatz.AdressNummer = tbMitarbeiter.Adressnummer" _
				& " where tbEinsatz.Datum = '" & Suchdatum & "' and tbMitarbeiter.Standort = '" & strStandort & "' and tbMitarbeiter.CheckDatum = '" & Suchdatum & "'", con)

            Count = cmd.ExecuteScalar()

        Catch ex As Exception

            Console.WriteLine("Invalid" & ex.Message)

        Finally
            con.Close()
        End Try

        Return Count

    End Function


	Public Function get_tbEinsatz_Total_Doppelt(strStandort)

	'Ein Mitarbeiter kann an einem Tag zwei Einsätze haben (Vormittag/Nachmittag). Für das Leitsystem ist aber nur ein Einsatz relevant.
	' Die doppelten Einträge pro Mitarbeiter werden gezählt und dann vom Total abgezogen.

		Dim con As New SqlConnection
		Dim cmd As New SqlCommand
		Dim reader As DbDataReader

		Dim Count As Integer = 0

		con.ConnectionString = SQLSERVER

		Dim AdressAlt As String
		Dim FirstRun As Integer = 0
		

		Try
			con.Open()
			cmd.Connection = con
		
			cmd = New SqlCommand("SELECT * FROM tbEinsatz INNER JOIN tbMitarbeiter ON tbEinsatz.AdressNummer = tbMitarbeiter.Adressnummer" _
				& " where tbEinsatz.Datum = '" & Suchdatum & "' and tbMitarbeiter.Standort = '" & strStandort & "'" _
				& " and tbMitarbeiter.CheckDatum = '" & Suchdatum & "' and tbEinsatz.AdressNummer <> '' Order by tbEinsatz.Datum, tbEinsatz.AdressNummer", con)

			reader = cmd.ExecuteReader

			While reader.Read

				If FirstRun = 0 Then
					FirstRun = 1
					AdressAlt = reader(1).ToString
				Else

					If AdressAlt = reader(1).ToString Then
						Count = Count + 1
					Else
						AdressAlt = reader(1).ToString
					End If

				End If

				
			End While

			reader.Close()

		Catch ex As Exception

			Console.WriteLine("Invalid" & ex.Message)

		Finally
			con.Close()
		End Try

	Return Count

	End Function


	Public Function get_tbEinsatz_mit_MA(strZuteilungCode)

		Dim con As New SqlConnection
		Dim cmd As New SqlCommand
		Dim Count As Integer = 0

		con.ConnectionString = SQLSERVER

		Try

			con.Open()
			cmd.Connection = con

			cmd = New SqlCommand("SELECT Count(*) FROM tbEinsatz INNER JOIN tbAuftrag ON tbEinsatz.AuftragsNummer = tbAuftrag.AuftragsNummer" _
								 & " WHERE tbAuftrag.ZuteilungCode = '" & strZuteilungCode & "' and tbEinsatz.Datum = '" & Suchdatum & "' AND tbEinsatz.AdressNummer <> '" & "'", con)

			Count = cmd.ExecuteScalar()

		Catch ex As Exception

			Console.WriteLine("Invalid" & ex.Message)

		Finally
			con.Close()
		End Try

		Return Count

	End Function

	Public Function get_tbEinsatz_mit_MA_Doppelt(strZuteilungCode)

	'Ein Mitarbeiter kann an einem Tag zwei Einsätze haben (Vormittag/Nachmittag). Für das Leitsystem ist aber nur ein Einsatz relevant.
	' Die doppelten Einträge pro Mitarbeiter werden gezählt und dann vom Total abgezogen.

		Dim con As New SqlConnection
		Dim cmd As New SqlCommand
		Dim reader As DbDataReader

		Dim Count As Integer = 0

		con.ConnectionString = SQLSERVER

		Dim AdressAlt As String
		Dim FirstRun As Integer = 0


		Try
			con.Open()
			cmd.Connection = con

			cmd = New SqlCommand("SELECT * FROM tbEinsatz INNER JOIN tbAuftrag ON tbEinsatz.AuftragsNummer = tbAuftrag.AuftragsNummer" _
								 & " WHERE tbAuftrag.ZuteilungCode = '" & strZuteilungCode & "'" _
								 & " and tbEinsatz.Datum = '" & Suchdatum & "' AND tbEinsatz.AdressNummer <> '" & "' Order by tbEinsatz.Datum,tbEinsatz.Adressnummer", con)

			reader = cmd.ExecuteReader

			While reader.Read

				If FirstRun = 0 Then
					FirstRun = 1
					AdressAlt = reader(1).ToString
				Else

					If AdressAlt = reader(1).ToString Then
						Count = Count + 1
					Else
						AdressAlt = reader(1).ToString
					End If

				End If


			End While

			reader.Close()

		Catch ex As Exception

			Console.WriteLine("Invalid" & ex.Message)

		Finally
			con.Close()
		End Try

	Return Count

	End Function

	Public Function get_Einsatz_Ohne_MA(strZuteilungCode)

		Dim con As New SqlConnection
		Dim cmd As New SqlCommand
		Dim Count As Integer = 0

		con.ConnectionString = SQLSERVER

		Try

			con.Open()
			cmd.Connection = con


			cmd = New SqlCommand("SELECT Count(*) FROM tbEinsatz INNER JOIN tbAuftrag ON tbEinsatz.AuftragsNummer = tbAuftrag.AuftragsNummer" _
								 & " WHERE tbAuftrag.ZuteilungCode = '" & strZuteilungCode & "' and tbEinsatz.Datum = '" & Suchdatum & "' AND tbEinsatz.AdressNummer = '" & "'", con)


			Count = cmd.ExecuteScalar()

		Catch ex As Exception

			Console.WriteLine("Invalid" & ex.Message)

		Finally
			con.Close()
		End Try

		Return Count

	End Function

	Function KalenderLesen(Montag As Integer)

		Dim con As New SqlConnection
		Dim cmd As New SqlCommand
		Dim reader As DbDataReader

		Dim Suchdatum As System.DateTime

		If DateTime.Now.ToString("dddd") = "Freitag" Then
			Suchdatum = DateTime.Now.Date
		Else
			Suchdatum = Today.AddDays(1)

		End If

		con.ConnectionString = SQLSERVER

		Try
			con.Open()
			cmd.Connection = con


			If Montag = 0 Then
				cmd = New SqlCommand("Select * from tbKalender where Datum = '" & Suchdatum & "'", con)
			Else
				cmd = New SqlCommand("Select * from tbKalender where Datum > '" & Suchdatum & "'", con)
			End If

			reader = cmd.ExecuteReader

			Dim CountMontag As Integer = 0

			While reader.Read


				If (reader(3).ToString) = "MO" Then
					CountMontag = CountMontag + 1
				End If

				If Montag = 0 Then
					RueckgabeDatum.D1 = (reader(1).ToString)
					RueckgabeDatum.D2 = (reader(4).ToString)
					RueckgabeDatum.D3 = (reader(2).ToString)
					RueckgabeDatum.D4 = (reader(5).ToString)
					Exit While
				End If

				If Montag = 1 And CountMontag = 1 Then
					RueckgabeDatum.D1 = (reader(1).ToString)
					RueckgabeDatum.D2 = (reader(4).ToString)
					RueckgabeDatum.D3 = (reader(2).ToString)
					RueckgabeDatum.D4 = (reader(5).ToString)
					Exit While
				End If

				If Montag = 2 And CountMontag = 2 Then
					RueckgabeDatum.D1 = (reader(1).ToString)
					RueckgabeDatum.D2 = (reader(4).ToString)
					RueckgabeDatum.D3 = (reader(2).ToString)
					RueckgabeDatum.D4 = (reader(5).ToString)
					Exit While
				End If

				If Montag = 3 And CountMontag = 3 Then
					RueckgabeDatum.D1 = (reader(1).ToString)
					RueckgabeDatum.D2 = (reader(4).ToString)
					RueckgabeDatum.D3 = (reader(2).ToString)
					RueckgabeDatum.D4 = (reader(5).ToString)
					Exit While
				End If


			End While

			reader.Close()

		Catch ex As Exception

			Console.WriteLine("Invalid" & ex.Message)

		Finally
			con.Close()
		End Try

		KalenderLesen = RueckgabeDatum


	End Function

	Public Structure DatumWert

		Public D1 As Date
		Public D2 As String
		Public D3 As String
		Public D4 As String

	End Structure

	Private Sub reset()

		F01 = F02 = F03 = F04 = F05 = F06 = F07 = F08 = F09 = F10 = 0
		F01_Doppelt = F02_Doppelt = F03_Doppelt = F04_Doppelt = F05_Doppelt = 0
		F06_Doppelt = F07_Doppelt = F08_Doppelt = F09_Doppelt = 0


	End Sub

	Public Function getMitarbeiter_OhneEinsatz(strStandort As String)


		Dim con As New SqlConnection
		Dim cmd As New SqlCommand
		Dim reader As DbDataReader

		Dim TotalCounter As Integer

		con.ConnectionString = SQLSERVER

		Try

			con.Open()
			cmd.Connection = con

			'cmd = New SqlCommand("SELECT * FROM tbMitarbeiter where Eintritt <= '" & Suchdatum & "' and Standort = '" & strStandort & "'", con)
			cmd = New SqlCommand("SELECT * FROM tbMitarbeiter where CheckDatum = '" & Suchdatum & "' and Standort = '" & strStandort & "'", con)

			reader = cmd.ExecuteReader

			Do While reader.Read



				If checkTermin(reader(2).ToString) = 0 Then

					If checkEinsatz(reader(2).ToString) = 0 Then
						TotalCounter = TotalCounter + 1
					End If

				End If

			Loop

			reader.Close()

		Catch ex As Exception

			Console.WriteLine("Invalid" & ex.Message)

		Finally
			con.Close()
		End Try

		Return TotalCounter


	End Function


	Function checkTermin(strAdressNummer As String)

		Dim con As New SqlConnection
		Dim cmd As New SqlCommand
		Dim count As Integer = 0

		con.ConnectionString = SQLSERVER

		Try

			con.Open()
			cmd.Connection = con

			cmd = New SqlCommand("SELECT Count(*) FROM tbTermin where Datum = '" & Suchdatum & "' and AdressNummer = '" & strAdressNummer & "'", con)
			count = cmd.ExecuteScalar

		Catch ex As Exception

			Console.WriteLine("Invalid" & ex.Message)

		Finally
			con.Close()
		End Try

		Return count

	End Function

	Function checkEinsatz(strAdressNummer As String)

		Dim con As New SqlConnection
		Dim cmd As New SqlCommand
		Dim count As Integer = 0

		con.ConnectionString = SQLSERVER

		Try

			con.Open()
			cmd.Connection = con

			cmd = New SqlCommand("SELECT Count(*) FROM tbEinsatz where Datum = '" & Suchdatum & "' and AdressNummer = '" & strAdressNummer & "'", con)
			count = cmd.ExecuteScalar

		Catch ex As Exception

			Console.WriteLine("Invalid" & ex.Message)

		Finally
			con.Close()
		End Try

		Return count

	End Function

End Module
