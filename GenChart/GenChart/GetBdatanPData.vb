﻿Imports System.Collections.Specialized
Imports System.Data.SqlClient
Imports System.Globalization

Public Class GetBDatanPData
    Public Shared Sub Main()
        Dim DateTimeB As New NameValueCollection
        Dim PlaceDataB As New NameValueCollection
        Dim BData As New NameValueCollection
        Dim PlaceData As New NameValueCollection
        Dim personalDetails As PersonalDetails = New PersonalDetails()
        Dim P_list(12) As String
        Dim H_List(12) As String
        Dim HstrCusp(12) As String
        Dim Hplanets(12) As String
        Dim BirthLagna(12, 2) As String
        Dim BirthBhav(12, 2) As String
        Dim BirthSouth(12) As String
        Dim Horo As New TASystem.TrueAstro

        Dim connection As SqlConnection = New SqlConnection("data source=49.50.103.132;initial catalog=ASTROLOGYSOFTWARE_DB;integrated security=False;User Id=sa;password=pSI)TA1t0K[)")
        Try
            Dim cmd As New SqlCommand($"SELECT TOP 1 HUSERID, HID, RECTIFIEDDATE, RECTIFIEDTIME, RECTIFIEDDST, RECTIFIEDPLACE, 
                                    RECTIFIEDLONGTITUDE, RECTIFIEDLONGTITUDEEW, RECTIFIEDLATITUDE,RECTIFIEDLATITUDENS
                                    FROM HMAIN
                                    INNER JOIN HREQUEST
                                    ON HMAIN.HUSERID = HREQUEST.RQUSERID
                                        AND HMAIN.HID = HREQUEST.RQHID
                                    WHERE HREQUEST.REQCAT = '9'
                                        AND HREQUEST.RQUNREAD = 'Y';", connection)
            Dim da As New SqlDataAdapter(cmd)
            Dim RowsData As New DataSet()
            da.Fill(RowsData)
            Dim UID = RowsData.Tables(0).Rows(0)(0).Trim.ToString()
            Dim HID = RowsData.Tables(0).Rows(0)(1).Trim.ToString()
            Dim RECTIFIEDDATE = CType(RowsData.Tables(0).Rows(0)(2), DateTime).ToString("yyyy-MM-dd HH:mm:ss.fff", CultureInfo.InvariantCulture)
            Dim RECTIFIEDTIME = CType(RowsData.Tables(0).Rows(0)(3).ToString(), DateTime).ToString("HH:mm:ss.fff", CultureInfo.InvariantCulture)
            Dim RECTIFIEDDST = RowsData.Tables(0).Rows(0)(4).ToString().Trim
            Dim RECTIFIEDPLACE = RowsData.Tables(0).Rows(0)(5).ToString().Trim
            Dim RECTIFIEDLONGTITUDE = RowsData.Tables(0).Rows(0)(6).Trim.ToString().Trim
            Dim RECTIFIEDLONGTITUDEEW = RowsData.Tables(0).Rows(0)(7).ToString().Trim
            Dim RECTIFIEDLATITUDE = RowsData.Tables(0).Rows(0)(8).ToString().Trim
            Dim RECTIFIEDLATITUDENS = RowsData.Tables(0).Rows(0)(9).ToString().Trim
            Dim Year = RECTIFIEDDATE.Split("-")(0)
            Dim Month = RECTIFIEDDATE.Split("-")(1)
            Dim Day = RECTIFIEDDATE.Split("-")(2).Substring(0, 2)
            Dim Hour = RECTIFIEDTIME.Split(":")(0)
            Dim Min = RECTIFIEDTIME.Split(":")(1)
            Dim Sec = RECTIFIEDTIME.Split(":")(2).Substring(0, 2)
            Dim MSec = RECTIFIEDTIME.Split(".")(1)
            BData.Add("Year", Year)
            BData.Add("Month", Month)
            BData.Add("Day", Day)
            BData.Add("Hour", Hour)
            BData.Add("Min", Min)
            BData.Add("Sec", Sec)
            BData.Add("mSec", MSec)
            DateTimeB = BData
            Dim Place = RECTIFIEDPLACE.Split("-")(0)
            Dim State = RECTIFIEDPLACE.Split("-")(1)
            Dim Country = RECTIFIEDPLACE.Split("-")(2)
#Region "Personal Details"
            personalDetails.DateofBirth = RECTIFIEDDATE.Split(" ")(0)
            personalDetails.DayofBirth = GetDayOfTheWeek(DateTime.Parse(RECTIFIEDDATE.Split(" ")(0).Replace("-", "/")).DayOfWeek)
            personalDetails.TimeofBirth = RECTIFIEDDATE.Split(" ")(1)
            personalDetails.PlaceofBirth = RECTIFIEDPLACE
            personalDetails.Latitude = RECTIFIEDLATITUDE
            personalDetails.Longitude = RECTIFIEDLONGTITUDE
            connection.Open()
            Dim command As New SqlCommand($"select USERNAME from UPROF WHERE USERID = '" + UID + "';", connection)
            Dim da0 As New SqlDataAdapter(command)
            Dim ds As New DataSet()
            da0.Fill(ds)
            connection.Close()
            personalDetails.NameoftheChartOwner = ds.Tables(0).Rows(0)(0).Trim.ToString()
            personalDetails.Rashi = "Not yet provided by DLL"
            personalDetails.Star = "Not yet provided by DLL"
            personalDetails.SunSign = "Not yet provided by DLL"
            personalDetails.BalanceDasa = "Not yet provided by DLL"
#End Region
            Dim lonB = RECTIFIEDLONGTITUDE
            Dim latB = RECTIFIEDLATITUDE
            Dim eW = RECTIFIEDLONGTITUDEEW
            Dim nS = RECTIFIEDLATITUDENS
            Dim TzB = "5.5"
            Dim DstB = RECTIFIEDDST
            PlaceData.Add("Place", Place)
            PlaceData.Add("State", State)
            PlaceData.Add("Country", Country)
            PlaceData.Add("lonB", lonB)
            PlaceData.Add("latB", latB)
            PlaceData.Add("eW", eW)
            PlaceData.Add("nS", nS)
            PlaceData.Add("TzB", TzB)
            PlaceData.Add("DstB", DstB)
            PlaceDataB = PlaceData
            Horo.getBirthPlanetCusp(DateTimeB, PlaceDataB, P_list, H_List)
            Horo.getBirthLagnaBhav(DateTimeB, PlaceDataB, BirthLagna, BirthBhav, BirthSouth)
            Dim DasaListP As New DataTable
            DasaListP.Columns.Add("Dasha")
            DasaListP.Columns.Add("Bhukti")
            DasaListP.Columns.Add("Antara")
            DasaListP.Columns.Add("Suk")
            DasaListP.Columns.Add("Pra")
            DasaListP.Columns.Add("Cl_Date")

            Horo.getBirthDasaFull(DateTimeB, PlaceDataB, DasaListP)
            Dim GetBDatanPData As GetBDatanPData = New GetBDatanPData()
            GetBDatanPData.UpdateHCusp(HID, UID, H_List)
            GetBDatanPData.UpdateHPlanet(HID, UID, P_list)
            Dim fillCuspTable As FillCuspTable = New FillCuspTable()
            fillCuspTable.FullCuspTableMain(HID, UID)
            Dim genChart As GenerateChart = New GenerateChart()
            genChart.GenerateChartMain(HID, UID, P_list, H_List, BirthLagna, BirthBhav, BirthSouth, DasaListP, personalDetails)
            GetBDatanPData.UpdateStatus(HID, UID)
        Catch ex As Exception
        Finally
            connection.Close()
        End Try
    End Sub
    Public Shared Function GetDayOfTheWeek(ByVal day As String) As String
        Dim DayOfWeek() = {"Sunday", "Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday"}
        Return DayOfWeek(Convert.ToInt32(day))
    End Function
    Sub UpdateHCusp(ByRef HID As String, ByRef UID As String, ByRef H_List As String())
        Dim con As New SqlConnection
        Dim cmd As New SqlCommand
        Try
            con.ConnectionString = "data source=49.50.103.132;initial catalog=HEADLETTERS_ENGINE;integrated security=False;User Id=sa;password=pSI)TA1t0K[)"
            con.Open()
            cmd.Connection = con
            Dim flag = False
            For Each rowData In H_List
                If flag = False Then
                    flag = True
                    Continue For
                End If
                cmd.CommandText = "INSERT INTO HCUSP VALUES ('" + UID + "','" + HID + "','" + GetCuspNo(rowData.Split("|")(0)) + "','" + rowData.Split("|")(1) + "','" + rowData.Split("|")(2) + "','" + rowData.Split("|")(3) + "','" + rowData.Split("|")(4) + "','" + rowData.Split("|")(5) + "');"
                cmd.ExecuteNonQuery()
            Next
        Catch ex As Exception
        Finally
            con.Close()
        End Try
    End Sub
    Sub UpdateHPlanet(ByRef HID As String, ByRef UID As String, ByRef P_List As String())
        Dim con As New SqlConnection
        Dim cmd As New SqlCommand
        Try
            con.ConnectionString = "data source=49.50.103.132;initial catalog=HEADLETTERS_ENGINE;integrated security=False;User Id=sa;password=pSI)TA1t0K[)"
            con.Open()
            cmd.Connection = con
            Dim flag = False
            For Each rowData In P_List
                If flag = False Then
                    flag = True
                    Continue For
                End If
                cmd.CommandText = "INSERT INTO HPLANET VALUES ('" + UID + "','" + HID + "','" + GetPlanetShortName(rowData.Split("|")(0)) + "','" + rowData.Split("|")(1) + "','" + rowData.Split("|")(2) + "','" + rowData.Split("|")(3) + "','" + rowData.Split("|")(4) + "','" + rowData.Split("|")(5) + "','" + rowData.Split("|")(6) + "','" + rowData.Split("|")(7) + "','" + rowData.Split("|")(8) + "','" + GetCuspPrefix(rowData.Split("|")(10)) + "');"
                cmd.ExecuteNonQuery()
            Next
        Catch ex As Exception
            Console.WriteLine("Error while inserting record on table..." & ex.Message, "Insert Records")
        Finally
            con.Close()
        End Try
    End Sub
    Sub UpdateStatus(ByRef HID As String, ByRef UID As String)
        Dim con As New SqlConnection
        Dim command As New SqlCommand
        Try
            con.ConnectionString = "data source=49.50.103.132;initial catalog=ASTROLOGYSOFTWARE_DB;integrated security=False;User Id=sa;password=pSI)TA1t0K[)"
            con.Open()
            command.Connection = con
            command.CommandText = $"UPDATE HREQUEST
                                    SET REQCAT = '9'
                                    WHERE RQUSERID = '" + UID + "'
                                        AND RQHID = '" + HID + "' 
                                            AND HREQUEST.REQCAT = '7' 
                                                AND HREQUEST.RQUNREAD = 'Y';"
            command.ExecuteNonQuery()
        Catch ex As Exception
        Finally
            con.Close()
        End Try
    End Sub
    Function GetCuspNo(ByVal s As String)
        Dim ov = "NA"
        Select Case s
            Case "Cusp I"
                ov = "01"
            Case "Cusp II"
                ov = "02"
            Case "Cusp III"
                ov = "03"
            Case "Cusp IV"
                ov = "04"
            Case "Cusp V"
                ov = "05"
            Case "Cusp VI"
                ov = "06"
            Case "Cusp VII"
                ov = "07"
            Case "Cusp VIII"
                ov = "08"
            Case "Cusp IX"
                ov = "09"
            Case "Cusp X"
                ov = "10"
            Case "Cusp XI"
                ov = "11"
            Case "Cusp XII"
                ov = "12"
            Case Else
        End Select
        Return ov
    End Function
    Function GetPlanetShortName(ByVal Pl As String)
        Dim Planet As String = "Not_A_Planet"
        Select Case Pl
            Case "Mars"
                Planet = "MA"
            Case "Venus"
                Planet = "VE"
            Case "Saturn"
                Planet = "SA"
            Case "Jupiter"
                Planet = "JU"
            Case "Sun"
                Planet = "SU"
            Case "Moon"
                Planet = "MO"
            Case "Mercury"
                Planet = "ME"
            Case "T.Rahu"
                Planet = "RA"
            Case "T.Ketu"
                Planet = "KE"
            Case "Neptune"
                Planet = "NE"
            Case "Pluto"
                Planet = "PL"
            Case "Uranus"
                Planet = "UR"
        End Select
        Return Planet
    End Function
    Function GetCuspPrefix(ByVal s As String)
        Dim ov = "NA"
        Select Case s
            Case "1"
                ov = "01"
            Case "2"
                ov = "02"
            Case "3"
                ov = "03"
            Case "4"
                ov = "04"
            Case "5"
                ov = "05"
            Case "6"
                ov = "06"
            Case "7"
                ov = "07"
            Case "8"
                ov = "08"
            Case "9"
                ov = "09"
            Case "10"
                ov = "10"
            Case "11"
                ov = "11"
            Case "12"
                ov = "12"
            Case Else
        End Select
        Return ov
    End Function
End Class
