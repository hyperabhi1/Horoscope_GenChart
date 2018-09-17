Imports System.Collections.Specialized
Imports System.Data.SqlClient
Public Class GetBDatanPData
	Sub GetBdatanPDataMain()
		Dim DateTimeB As New NameValueCollection
		Dim PlaceDataB As New NameValueCollection
		Dim BData As New NameValueCollection
		Dim PlaceData As New NameValueCollection
		Dim P_list(12) As String
		Dim H_List(12) As String
		Dim HstrCusp(12) As String
		Dim Hplanets(12) As String
		Dim BirthLagna(12, 2) As String
		Dim BirthBhav(12, 2) As String
		Dim BirthSouth(12) As String
		Dim Horo As New TASystem.TrueAstro

		Dim connection As SqlConnection = New SqlConnection("data source=WIN-KSTUPT6CJRC;initial catalog=ASTROLOGYSOFTWARE_DB;integrated security=True;multipleactiveresultsets=True;")
		connection.Open()
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
		connection.Close()
		Dim UID = RowsData.Tables(0).Rows(0)(0).Trim.ToString()
		Dim HID = RowsData.Tables(0).Rows(0)(1).Trim.ToString()
		Dim RECTIFIEDDATE = RowsData.Tables(0).Rows(0)(2).Trim.ToString()
		Dim RECTIFIEDTIME = RowsData.Tables(0).Rows(0)(3).Trim.ToString()
		Dim RECTIFIEDDST = RowsData.Tables(0).Rows(0)(4).Trim.ToString()
		Dim RECTIFIEDPLACE = RowsData.Tables(0).Rows(0)(5).Trim.ToString()
		Dim RECTIFIEDLONGTITUDE = RowsData.Tables(0).Rows(0)(6).Trim.ToString()
		Dim RECTIFIEDLONGTITUDEEW = RowsData.Tables(0).Rows(0)(7).Trim.ToString()
		Dim RECTIFIEDLATITUDE = RowsData.Tables(0).Rows(0)(8).Trim.ToString()
		Dim RECTIFIEDLATITUDENS = RowsData.Tables(0).Rows(0)(9).Trim.ToString()
        DateofBirth = RECTIFIEDDATE.Split(" ")(0)
        DayofBirth = GetDayOfTheWeek(DateTime.Parse(RECTIFIEDDATE.Split(" ")(0).Replace("-", "/")).DayOfWeek)
        TimeofBirth = RECTIFIEDDATE.Split(" ")(1)
        Dim Year = RECTIFIEDDATE.Split("-")(0)
		Dim Month = RECTIFIEDDATE.Split("-")(1)
		Dim Day = RECTIFIEDDATE.Split("-")(2).Substring(0, 2)
		Dim Hour = RECTIFIEDTIME.Split(":")(0).Substring(0, 2)
		Dim Min = RECTIFIEDTIME.Split(":")(1)
		Dim Sec = RECTIFIEDTIME.Split(":")(2)
		Dim MSec = RECTIFIEDTIME.Split(".")(1)
		BData.Add("Year", Year)
		BData.Add("Month", Month)
		BData.Add("Day", Day)
		BData.Add("Hour", Hour)
		BData.Add("Min", Min)
		BData.Add("Sec", Sec)
		BData.Add("mSec", MSec)
		DateTimeB = BData
        PlaceofBirth = RECTIFIEDPLACE
        Dim Place = RECTIFIEDPLACE.Split("-")(0)
		Dim State = RECTIFIEDPLACE.Split("-")(1)
        Dim Country = RECTIFIEDPLACE.Split("-")(2)
        Latitude = RECTIFIEDLATITUDE
        Longitude = RECTIFIEDLONGTITUDE

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
		UpdateHCusp(HID, UID, H_List)
		UpdateHPlanet(HID, UID, P_list)
		UpdateStatus(HID, UID)
	End Sub

	Sub UpdateHCusp(ByRef HID As String, ByRef UID As String, ByRef H_List As String())
		Dim con As New SqlConnection
		Dim cmd As New SqlCommand
		Try
			con.ConnectionString = "data source=WIN-KSTUPT6CJRC;initial catalog=ASTROLOGYSOFTWARE_DB;integrated security=True;multipleactiveresultsets=True;"
			con.Open()
			cmd.Connection = con
			Dim flag = False
			For Each rowData In H_List
				If flag = False Then
					flag = True
					Continue For
				End If
				cmd.CommandText = "INSERT INTO _HCUSP VALUES ('" + UID + "','" + HID + "','" + rowData.Split("|")(0) + "','" + rowData.Split("|")(1) + "','" + rowData.Split("|")(2) + "','" + rowData.Split("|")(3) + "','" + rowData.Split("|")(4) + "','" + rowData.Split("|")(5) + "','" + rowData.Split("|")(6) + "','" + rowData.Split("|")(7) + "');"
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
			con.ConnectionString = "data source=WIN-KSTUPT6CJRC;initial catalog=ASTROLOGYSOFTWARE_DB;integrated security=True;multipleactiveresultsets=True;"
			con.Open()
			cmd.Connection = con
			Dim flag = False
			For Each rowData In P_List
				If flag = False Then
					flag = True
					Continue For
				End If
				cmd.CommandText = "INSERT INTO _HPLANET VALUES ('" + UID + "','" + HID + "','" + rowData.Split("|")(0) + "','" + rowData.Split("|")(1) + "','" + rowData.Split("|")(2) + "','" + rowData.Split("|")(3) + "','" + rowData.Split("|")(4) + "','" + rowData.Split("|")(5) + "','" + rowData.Split("|")(6) + "','" + rowData.Split("|")(7) + "','" + rowData.Split("|")(8) + "','" + rowData.Split("|")(9) + "');"
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
            con.ConnectionString = "data source=WIN-KSTUPT6CJRC;initial catalog=ASTROLOGYSOFTWARE_DB;integrated security=True;multipleactiveresultsets=True;"
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
    Function GetDayOfTheWeek(ByVal day As String) As String
        Dim DayOfWeek() = {"Sunday", "Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday"}
        Return DayOfWeek(Convert.ToInt32(day))
    End Function
End Class
