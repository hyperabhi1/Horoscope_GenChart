Imports System.Data.SqlClient

Public Class Gen_MatchFile
    Dim c01(), c02(), c03(), c04(), c05(), c06(), c07(), c08(), c09(), c10(), c11(), c12() As String
    Dim counter_MA = 0
    Dim counter_ME = 0
    Dim counter_MO = 0
    Dim counter_SA = 0
    Dim counter_JU = 0
    Dim counter_VE = 0
    Dim counter_SU = 0
    Dim counter_AL = 0
    Dim ThreadStartCounter = 0
    Dim ThreadEndCounter = 0
    Dim duplicates = 0
    'Dim connstr = "data source=49.50.103.132;initial catalog=HEADLETTERS10;integrated security=False;User Id=sa;password=pSI)TA1t0K[)"
    Dim connstr = "data source=WIN-KSTUPT6CJRC;initial catalog=HEADLETTERS10;integrated security=True;multipleactiveresultsets=True;"
    'Dim connstr = "data source=DESKTOP-JBRFH9E;initial catalog=testdb;integrated security=True;"
    Sub Gen_MatchFileMain()
        FillCusp()
        TruncateMatch_File()
        Parallel.Invoke(Sub() Match_File_MA(), Sub() Match_File_ME(), Sub() Match_File_MO(), Sub() Match_File_JU(), Sub() Match_File_SA(), Sub() Match_File_SU(), Sub() Match_File_VE())
        NumberOfRecordsInMatchFile()
    End Sub
    Sub FillCusp()
        Dim SelectCUSP = "SELECT * FROM CUSP;"
        Dim con As New SqlConnection(connstr)
        Dim connection As SqlConnection = New SqlConnection(connstr)
        connection.Open()
        Dim cmd As New SqlCommand(SelectCUSP, con)
        Dim da As New SqlDataAdapter(cmd)
        Dim ds As New DataSet()
        da.Fill(ds)
        connection.Close()
        For Each row As DataRow In ds.Tables(0).Rows
            Dim l As Integer = row.Item(3).ToString().Trim.Length / 2
            Select Case row.Item(2).ToString()
                Case "01"
                    c01 = New String(l - 1) {}
                    Dim START = 0
                    For I As Integer = 0 To l - 1
                        c01(I) = row.Item(3).ToString().Trim.Substring(START, 2)
                        START = START + 2
                    Next

                Case "02"
                    c02 = New String(l - 1) {}
                    Dim START = 0
                    For I As Integer = 0 To l - 1
                        c02(I) = row.Item(3).ToString().Trim.Substring(START, 2)
                        START = START + 2
                    Next
                Case "03"
                    c03 = New String(l - 1) {}
                    Dim START = 0
                    For I As Integer = 0 To l - 1
                        c03(I) = row.Item(3).ToString().Trim.Substring(START, 2)
                        START = START + 2
                    Next
                Case "04"
                    c04 = New String(l - 1) {}
                    Dim START = 0
                    For I As Integer = 0 To l - 1
                        c04(I) = row.Item(3).ToString().Trim.Substring(START, 2)
                        START = START + 2
                    Next
                Case "05"
                    c05 = New String(l - 1) {}
                    Dim START = 0
                    For I As Integer = 0 To l - 1
                        c05(I) = row.Item(3).ToString().Trim.Substring(START, 2)
                        START = START + 2
                    Next
                Case "06"
                    c06 = New String(l - 1) {}
                    Dim START = 0
                    For I As Integer = 0 To l - 1
                        c06(I) = row.Item(3).ToString().Trim.Substring(START, 2)
                        START = START + 2
                    Next
                Case "07"
                    c07 = New String(l - 1) {}
                    Dim START = 0
                    For I As Integer = 0 To l - 1
                        c07(I) = row.Item(3).ToString().Trim.Substring(START, 2)
                        START = START + 2
                    Next
                Case "08"
                    c08 = New String(l - 1) {}
                    Dim START = 0
                    For I As Integer = 0 To l - 1
                        c08(I) = row.Item(3).ToString().Trim.Substring(START, 2)
                        START = START + 2
                    Next
                Case "09"
                    c09 = New String(l - 1) {}
                    Dim START = 0
                    For I As Integer = 0 To l - 1
                        c09(I) = row.Item(3).ToString().Trim.Substring(START, 2)
                        START = START + 2
                    Next
                Case "10"
                    c10 = New String(l - 1) {}
                    Dim START = 0
                    For I As Integer = 0 To l - 1
                        c10(I) = row.Item(3).ToString().Trim.Substring(START, 2)
                        START = START + 2
                    Next
                Case "11"
                    c11 = New String(l - 1) {}
                    Dim START = 0
                    For I As Integer = 0 To l - 1
                        c11(I) = row.Item(3).ToString().Trim.Substring(START, 2)
                        START = START + 2
                    Next
                Case "12"
                    c12 = New String(l - 1) {}
                    Dim START = 0
                    For I As Integer = 0 To l - 1
                        c12(I) = row.Item(3).ToString().Trim.Substring(START, 2)
                        START = START + 2
                    Next
                Case Else
            End Select
        Next
        connection.Close()
    End Sub
    Sub Match_File_ALL()

        Dim m_planet = "AL"
        Dim a8(8) As String
        Dim con As New SqlConnection(connstr)
        Dim connection As SqlConnection = New SqlConnection(connstr)
        connection.Open()
        Dim cmd As New SqlCommand("SELECT * FROM F2BASE;", con)
        Dim da As New SqlDataAdapter(cmd)
        Dim ds As New DataSet()
        da.Fill(ds)
        counter_AL = ds.Tables(0).Rows.Count - 1
        connection.Close()
        Parallel.Invoke(
                Sub() Process_match_Key_set_ALL(ds, m_planet),
                Sub() Process_match_Key_set_ALL(ds, m_planet),
                Sub() Process_match_Key_set_ALL(ds, m_planet),
                Sub() Process_match_Key_set_ALL(ds, m_planet),
                Sub() Process_match_Key_set_ALL(ds, m_planet),
                Sub() Process_match_Key_set_ALL(ds, m_planet),
                Sub() Process_match_Key_set_ALL(ds, m_planet),
                Sub() Process_match_Key_set_ALL(ds, m_planet),
                Sub() Process_match_Key_set_ALL(ds, m_planet),
                Sub() Process_match_Key_set_ALL(ds, m_planet),
                Sub() Process_match_Key_set_ALL(ds, m_planet),
                Sub() Process_match_Key_set_ALL(ds, m_planet),
                Sub() Process_match_Key_set_ALL(ds, m_planet),
                Sub() Process_match_Key_set_ALL(ds, m_planet),
                Sub() Process_match_Key_set_ALL(ds, m_planet),
                Sub() Process_match_Key_set_ALL(ds, m_planet),
                Sub() Process_match_Key_set_ALL(ds, m_planet),
                Sub() Process_match_Key_set_ALL(ds, m_planet),
                Sub() Process_match_Key_set_ALL(ds, m_planet),
                Sub() Process_match_Key_set_ALL(ds, m_planet),
                Sub() Process_match_Key_set_ALL(ds, m_planet),
                Sub() Process_match_Key_set_ALL(ds, m_planet),
                Sub() Process_match_Key_set_ALL(ds, m_planet),
                Sub() Process_match_Key_set_ALL(ds, m_planet),
                Sub() Process_match_Key_set_ALL(ds, m_planet),
                Sub() Process_match_Key_set_ALL(ds, m_planet),
                Sub() Process_match_Key_set_ALL(ds, m_planet),
                Sub() Process_match_Key_set_ALL(ds, m_planet),
                Sub() Process_match_Key_set_ALL(ds, m_planet),
                Sub() Process_match_Key_set_ALL(ds, m_planet),
                Sub() Process_match_Key_set_ALL(ds, m_planet),
                Sub() Process_match_Key_set_ALL(ds, m_planet),
                Sub() Process_match_Key_set_ALL(ds, m_planet),
                Sub() Process_match_Key_set_ALL(ds, m_planet),
                Sub() Process_match_Key_set_ALL(ds, m_planet),
                Sub() Process_match_Key_set_ALL(ds, m_planet),
                Sub() Process_match_Key_set_ALL(ds, m_planet),
                Sub() Process_match_Key_set_ALL(ds, m_planet),
                Sub() Process_match_Key_set_ALL(ds, m_planet),
                Sub() Process_match_Key_set_ALL(ds, m_planet))
    End Sub
    Sub Match_File_MA()
        Dim m_planet = "MA"
        Dim a8(8) As String
        Dim con As New SqlConnection(connstr)
        Dim connection As SqlConnection = New SqlConnection(connstr)
        connection.Open()
        Dim cmd As New SqlCommand("SELECT * FROM F2BASE WHERE p1 = 'MA';", con)
        Dim da As New SqlDataAdapter(cmd)
        Dim ds As New DataSet()
        da.Fill(ds)
        counter_MA = ds.Tables(0).Rows.Count - 1
        connection.Close()
        Parallel.Invoke(
               Sub() Process_match_Key_set_MA(ds, m_planet),
               Sub() Process_match_Key_set_MA(ds, m_planet),
               Sub() Process_match_Key_set_MA(ds, m_planet),
               Sub() Process_match_Key_set_MA(ds, m_planet),
               Sub() Process_match_Key_set_MA(ds, m_planet),
               Sub() Process_match_Key_set_MA(ds, m_planet),
               Sub() Process_match_Key_set_MA(ds, m_planet),
               Sub() Process_match_Key_set_MA(ds, m_planet),
               Sub() Process_match_Key_set_MA(ds, m_planet),
               Sub() Process_match_Key_set_MA(ds, m_planet),
               Sub() Process_match_Key_set_MA(ds, m_planet),
               Sub() Process_match_Key_set_MA(ds, m_planet),
               Sub() Process_match_Key_set_MA(ds, m_planet),
               Sub() Process_match_Key_set_MA(ds, m_planet),
               Sub() Process_match_Key_set_MA(ds, m_planet),
               Sub() Process_match_Key_set_MA(ds, m_planet),
               Sub() Process_match_Key_set_MA(ds, m_planet),
               Sub() Process_match_Key_set_MA(ds, m_planet),
               Sub() Process_match_Key_set_MA(ds, m_planet),
               Sub() Process_match_Key_set_MA(ds, m_planet),
               Sub() Process_match_Key_set_MA(ds, m_planet),
               Sub() Process_match_Key_set_MA(ds, m_planet),
               Sub() Process_match_Key_set_MA(ds, m_planet),
               Sub() Process_match_Key_set_MA(ds, m_planet),
               Sub() Process_match_Key_set_MA(ds, m_planet),
               Sub() Process_match_Key_set_MA(ds, m_planet),
               Sub() Process_match_Key_set_MA(ds, m_planet),
               Sub() Process_match_Key_set_MA(ds, m_planet),
               Sub() Process_match_Key_set_MA(ds, m_planet),
               Sub() Process_match_Key_set_MA(ds, m_planet),
               Sub() Process_match_Key_set_MA(ds, m_planet),
               Sub() Process_match_Key_set_MA(ds, m_planet),
               Sub() Process_match_Key_set_MA(ds, m_planet),
               Sub() Process_match_Key_set_MA(ds, m_planet),
               Sub() Process_match_Key_set_MA(ds, m_planet),
               Sub() Process_match_Key_set_MA(ds, m_planet),
               Sub() Process_match_Key_set_MA(ds, m_planet),
               Sub() Process_match_Key_set_MA(ds, m_planet),
               Sub() Process_match_Key_set_MA(ds, m_planet),
               Sub() Process_match_Key_set_MA(ds, m_planet))
    End Sub
    Sub Match_File_JU()
        Dim m_planet = "JU"
        Dim a8(8) As String
        Dim con As New SqlConnection(connstr)
        Dim connection As SqlConnection = New SqlConnection(connstr)
        connection.Open()
        Dim cmd As New SqlCommand("SELECT * FROM F2BASE WHERE p1 = 'JU';", con)
        Dim da As New SqlDataAdapter(cmd)
        Dim ds As New DataSet()
        da.Fill(ds)
        counter_JU = ds.Tables(0).Rows.Count - 1
        connection.Close()
        Parallel.Invoke(
                Sub() Process_match_Key_set_JU(ds, m_planet),
                Sub() Process_match_Key_set_JU(ds, m_planet),
                Sub() Process_match_Key_set_JU(ds, m_planet),
                Sub() Process_match_Key_set_JU(ds, m_planet),
                Sub() Process_match_Key_set_JU(ds, m_planet),
                Sub() Process_match_Key_set_JU(ds, m_planet),
                Sub() Process_match_Key_set_JU(ds, m_planet),
                Sub() Process_match_Key_set_JU(ds, m_planet),
                Sub() Process_match_Key_set_JU(ds, m_planet),
                Sub() Process_match_Key_set_JU(ds, m_planet),
                Sub() Process_match_Key_set_JU(ds, m_planet),
                Sub() Process_match_Key_set_JU(ds, m_planet),
                Sub() Process_match_Key_set_JU(ds, m_planet),
                Sub() Process_match_Key_set_JU(ds, m_planet),
                Sub() Process_match_Key_set_JU(ds, m_planet),
                Sub() Process_match_Key_set_JU(ds, m_planet),
                Sub() Process_match_Key_set_JU(ds, m_planet),
                Sub() Process_match_Key_set_JU(ds, m_planet),
                Sub() Process_match_Key_set_JU(ds, m_planet),
                Sub() Process_match_Key_set_JU(ds, m_planet),
                Sub() Process_match_Key_set_JU(ds, m_planet),
                Sub() Process_match_Key_set_JU(ds, m_planet),
                Sub() Process_match_Key_set_JU(ds, m_planet),
                Sub() Process_match_Key_set_JU(ds, m_planet),
                Sub() Process_match_Key_set_JU(ds, m_planet),
                Sub() Process_match_Key_set_JU(ds, m_planet),
                Sub() Process_match_Key_set_JU(ds, m_planet),
                Sub() Process_match_Key_set_JU(ds, m_planet),
                Sub() Process_match_Key_set_JU(ds, m_planet),
                Sub() Process_match_Key_set_JU(ds, m_planet),
                Sub() Process_match_Key_set_JU(ds, m_planet),
                Sub() Process_match_Key_set_JU(ds, m_planet),
                Sub() Process_match_Key_set_JU(ds, m_planet),
                Sub() Process_match_Key_set_JU(ds, m_planet),
                Sub() Process_match_Key_set_JU(ds, m_planet),
                Sub() Process_match_Key_set_JU(ds, m_planet),
                Sub() Process_match_Key_set_JU(ds, m_planet),
                Sub() Process_match_Key_set_JU(ds, m_planet),
                Sub() Process_match_Key_set_JU(ds, m_planet),
                Sub() Process_match_Key_set_JU(ds, m_planet))
    End Sub
    Sub Match_File_SA()
        Dim m_planet = "SA"
        Dim a8(8) As String
        Dim con As New SqlConnection(connstr)
        Dim connection As SqlConnection = New SqlConnection(connstr)
        connection.Open()
        Dim cmd As New SqlCommand("SELECT * FROM F2BASE WHERE p1 = 'SA';", con)
        Dim da As New SqlDataAdapter(cmd)
        Dim ds As New DataSet()
        da.Fill(ds)
        counter_SA = ds.Tables(0).Rows.Count - 1
        connection.Close()
        Parallel.Invoke(
                Sub() Process_match_Key_set_SA(ds, m_planet),
                Sub() Process_match_Key_set_SA(ds, m_planet),
                Sub() Process_match_Key_set_SA(ds, m_planet),
                Sub() Process_match_Key_set_SA(ds, m_planet),
                Sub() Process_match_Key_set_SA(ds, m_planet),
                Sub() Process_match_Key_set_SA(ds, m_planet),
                Sub() Process_match_Key_set_SA(ds, m_planet),
                Sub() Process_match_Key_set_SA(ds, m_planet),
                Sub() Process_match_Key_set_SA(ds, m_planet),
                Sub() Process_match_Key_set_SA(ds, m_planet),
                Sub() Process_match_Key_set_SA(ds, m_planet),
                Sub() Process_match_Key_set_SA(ds, m_planet),
                Sub() Process_match_Key_set_SA(ds, m_planet),
                Sub() Process_match_Key_set_SA(ds, m_planet),
                Sub() Process_match_Key_set_SA(ds, m_planet),
                Sub() Process_match_Key_set_SA(ds, m_planet),
                Sub() Process_match_Key_set_SA(ds, m_planet),
                Sub() Process_match_Key_set_SA(ds, m_planet),
                Sub() Process_match_Key_set_SA(ds, m_planet),
                Sub() Process_match_Key_set_SA(ds, m_planet),
                Sub() Process_match_Key_set_SA(ds, m_planet),
                Sub() Process_match_Key_set_SA(ds, m_planet),
                Sub() Process_match_Key_set_SA(ds, m_planet),
                Sub() Process_match_Key_set_SA(ds, m_planet),
                Sub() Process_match_Key_set_SA(ds, m_planet),
                Sub() Process_match_Key_set_SA(ds, m_planet),
                Sub() Process_match_Key_set_SA(ds, m_planet),
                Sub() Process_match_Key_set_SA(ds, m_planet),
                Sub() Process_match_Key_set_SA(ds, m_planet),
                Sub() Process_match_Key_set_SA(ds, m_planet),
                Sub() Process_match_Key_set_SA(ds, m_planet),
                Sub() Process_match_Key_set_SA(ds, m_planet),
                Sub() Process_match_Key_set_SA(ds, m_planet),
                Sub() Process_match_Key_set_SA(ds, m_planet),
                Sub() Process_match_Key_set_SA(ds, m_planet),
                Sub() Process_match_Key_set_SA(ds, m_planet),
                Sub() Process_match_Key_set_SA(ds, m_planet),
                Sub() Process_match_Key_set_SA(ds, m_planet),
                Sub() Process_match_Key_set_SA(ds, m_planet),
                Sub() Process_match_Key_set_SA(ds, m_planet))
    End Sub
    Sub Match_File_ME()
        Dim m_planet = "ME"
        Dim a8(8) As String
        Dim con As New SqlConnection(connstr)
        Dim connection As SqlConnection = New SqlConnection(connstr)
        connection.Open()
        Dim cmd As New SqlCommand("SELECT * FROM F2BASE WHERE p1 = 'ME';", con)
        Dim da As New SqlDataAdapter(cmd)
        Dim ds As New DataSet()
        da.Fill(ds)
        counter_ME = ds.Tables(0).Rows.Count - 1
        connection.Close()
        Parallel.Invoke(
                Sub() Process_match_Key_set_ME(ds, m_planet),
                Sub() Process_match_Key_set_ME(ds, m_planet),
                Sub() Process_match_Key_set_ME(ds, m_planet),
                Sub() Process_match_Key_set_ME(ds, m_planet),
                Sub() Process_match_Key_set_ME(ds, m_planet),
                Sub() Process_match_Key_set_ME(ds, m_planet),
                Sub() Process_match_Key_set_ME(ds, m_planet),
                Sub() Process_match_Key_set_ME(ds, m_planet),
                Sub() Process_match_Key_set_ME(ds, m_planet),
                Sub() Process_match_Key_set_ME(ds, m_planet),
                Sub() Process_match_Key_set_ME(ds, m_planet),
                Sub() Process_match_Key_set_ME(ds, m_planet),
                Sub() Process_match_Key_set_ME(ds, m_planet),
                Sub() Process_match_Key_set_ME(ds, m_planet),
                Sub() Process_match_Key_set_ME(ds, m_planet),
                Sub() Process_match_Key_set_ME(ds, m_planet),
                Sub() Process_match_Key_set_ME(ds, m_planet),
                Sub() Process_match_Key_set_ME(ds, m_planet),
                Sub() Process_match_Key_set_ME(ds, m_planet),
                Sub() Process_match_Key_set_ME(ds, m_planet),
                Sub() Process_match_Key_set_ME(ds, m_planet),
                Sub() Process_match_Key_set_ME(ds, m_planet),
                Sub() Process_match_Key_set_ME(ds, m_planet),
                Sub() Process_match_Key_set_ME(ds, m_planet),
                Sub() Process_match_Key_set_ME(ds, m_planet),
                Sub() Process_match_Key_set_ME(ds, m_planet),
                Sub() Process_match_Key_set_ME(ds, m_planet),
                Sub() Process_match_Key_set_ME(ds, m_planet),
                Sub() Process_match_Key_set_ME(ds, m_planet),
                Sub() Process_match_Key_set_ME(ds, m_planet),
                Sub() Process_match_Key_set_ME(ds, m_planet),
                Sub() Process_match_Key_set_ME(ds, m_planet),
                Sub() Process_match_Key_set_ME(ds, m_planet),
                Sub() Process_match_Key_set_ME(ds, m_planet),
                Sub() Process_match_Key_set_ME(ds, m_planet),
                Sub() Process_match_Key_set_ME(ds, m_planet),
                Sub() Process_match_Key_set_ME(ds, m_planet),
                Sub() Process_match_Key_set_ME(ds, m_planet),
                Sub() Process_match_Key_set_ME(ds, m_planet),
                Sub() Process_match_Key_set_ME(ds, m_planet))
    End Sub
    Sub Match_File_MO()
        Dim m_planet = "MO"
        Dim a8(8) As String
        Dim con As New SqlConnection(connstr)
        Dim connection As SqlConnection = New SqlConnection(connstr)
        connection.Open()
        Dim cmd As New SqlCommand("SELECT * FROM F2BASE WHERE p1 = 'MO';", con)
        Dim da As New SqlDataAdapter(cmd)
        Dim ds As New DataSet()
        da.Fill(ds)
        counter_MO = ds.Tables(0).Rows.Count - 1
        connection.Close()
        Parallel.Invoke(
                Sub() Process_match_Key_set_MO(ds, m_planet),
                Sub() Process_match_Key_set_MO(ds, m_planet),
                Sub() Process_match_Key_set_MO(ds, m_planet),
                Sub() Process_match_Key_set_MO(ds, m_planet),
                Sub() Process_match_Key_set_MO(ds, m_planet),
                Sub() Process_match_Key_set_MO(ds, m_planet),
                Sub() Process_match_Key_set_MO(ds, m_planet),
                Sub() Process_match_Key_set_MO(ds, m_planet),
                Sub() Process_match_Key_set_MO(ds, m_planet),
                Sub() Process_match_Key_set_MO(ds, m_planet),
                Sub() Process_match_Key_set_MO(ds, m_planet),
                Sub() Process_match_Key_set_MO(ds, m_planet),
                Sub() Process_match_Key_set_MO(ds, m_planet),
                Sub() Process_match_Key_set_MO(ds, m_planet),
                Sub() Process_match_Key_set_MO(ds, m_planet),
                Sub() Process_match_Key_set_MO(ds, m_planet),
                Sub() Process_match_Key_set_MO(ds, m_planet),
                Sub() Process_match_Key_set_MO(ds, m_planet),
                Sub() Process_match_Key_set_MO(ds, m_planet),
                Sub() Process_match_Key_set_MO(ds, m_planet),
                Sub() Process_match_Key_set_MO(ds, m_planet),
                Sub() Process_match_Key_set_MO(ds, m_planet),
                Sub() Process_match_Key_set_MO(ds, m_planet),
                Sub() Process_match_Key_set_MO(ds, m_planet),
                Sub() Process_match_Key_set_MO(ds, m_planet),
                Sub() Process_match_Key_set_MO(ds, m_planet),
                Sub() Process_match_Key_set_MO(ds, m_planet),
                Sub() Process_match_Key_set_MO(ds, m_planet),
                Sub() Process_match_Key_set_MO(ds, m_planet),
                Sub() Process_match_Key_set_MO(ds, m_planet),
                Sub() Process_match_Key_set_MO(ds, m_planet),
                Sub() Process_match_Key_set_MO(ds, m_planet),
                Sub() Process_match_Key_set_MO(ds, m_planet),
                Sub() Process_match_Key_set_MO(ds, m_planet),
                Sub() Process_match_Key_set_MO(ds, m_planet),
                Sub() Process_match_Key_set_MO(ds, m_planet),
                Sub() Process_match_Key_set_MO(ds, m_planet),
                Sub() Process_match_Key_set_MO(ds, m_planet),
                Sub() Process_match_Key_set_MO(ds, m_planet),
                Sub() Process_match_Key_set_MO(ds, m_planet))
    End Sub
    Sub Match_File_VE()
        Dim m_planet = "VE"
        Dim a8(8) As String
        Dim con As New SqlConnection(connstr)
        Dim connection As SqlConnection = New SqlConnection(connstr)
        connection.Open()
        Dim cmd As New SqlCommand("SELECT * FROM F2BASE WHERE p1 = 'VE';", con)
        Dim da As New SqlDataAdapter(cmd)
        Dim ds As New DataSet()
        da.Fill(ds)
        counter_VE = ds.Tables(0).Rows.Count - 1
        connection.Close()
        Parallel.Invoke(
                Sub() Process_match_Key_set_VE(ds, m_planet),
                Sub() Process_match_Key_set_VE(ds, m_planet),
                Sub() Process_match_Key_set_VE(ds, m_planet),
                Sub() Process_match_Key_set_VE(ds, m_planet),
                Sub() Process_match_Key_set_VE(ds, m_planet),
                Sub() Process_match_Key_set_VE(ds, m_planet),
                Sub() Process_match_Key_set_VE(ds, m_planet),
                Sub() Process_match_Key_set_VE(ds, m_planet),
                Sub() Process_match_Key_set_VE(ds, m_planet),
                Sub() Process_match_Key_set_VE(ds, m_planet),
                Sub() Process_match_Key_set_VE(ds, m_planet),
                Sub() Process_match_Key_set_VE(ds, m_planet),
                Sub() Process_match_Key_set_VE(ds, m_planet),
                Sub() Process_match_Key_set_VE(ds, m_planet),
                Sub() Process_match_Key_set_VE(ds, m_planet),
                Sub() Process_match_Key_set_VE(ds, m_planet),
                Sub() Process_match_Key_set_VE(ds, m_planet),
                Sub() Process_match_Key_set_VE(ds, m_planet),
                Sub() Process_match_Key_set_VE(ds, m_planet),
                Sub() Process_match_Key_set_VE(ds, m_planet),
                Sub() Process_match_Key_set_VE(ds, m_planet),
                Sub() Process_match_Key_set_VE(ds, m_planet),
                Sub() Process_match_Key_set_VE(ds, m_planet),
                Sub() Process_match_Key_set_VE(ds, m_planet),
                Sub() Process_match_Key_set_VE(ds, m_planet),
                Sub() Process_match_Key_set_VE(ds, m_planet),
                Sub() Process_match_Key_set_VE(ds, m_planet),
                Sub() Process_match_Key_set_VE(ds, m_planet),
                Sub() Process_match_Key_set_VE(ds, m_planet),
                Sub() Process_match_Key_set_VE(ds, m_planet),
                Sub() Process_match_Key_set_VE(ds, m_planet),
                Sub() Process_match_Key_set_VE(ds, m_planet),
                Sub() Process_match_Key_set_VE(ds, m_planet),
                Sub() Process_match_Key_set_VE(ds, m_planet),
                Sub() Process_match_Key_set_VE(ds, m_planet),
                Sub() Process_match_Key_set_VE(ds, m_planet),
                Sub() Process_match_Key_set_VE(ds, m_planet),
                Sub() Process_match_Key_set_VE(ds, m_planet),
                Sub() Process_match_Key_set_VE(ds, m_planet),
                Sub() Process_match_Key_set_VE(ds, m_planet))
    End Sub
    Sub Match_File_SU()
        Dim m_planet = "SU"
        Dim a8(8) As String
        Dim con As New SqlConnection(connstr)
        Dim connection As SqlConnection = New SqlConnection(connstr)
        connection.Open()
        Dim cmd As New SqlCommand("SELECT * FROM F2BASE WHERE p1 = 'SU';", con)
        Dim da As New SqlDataAdapter(cmd)
        Dim ds As New DataSet()
        da.Fill(ds)
        counter_SU = ds.Tables(0).Rows.Count - 1
        connection.Close()
        Parallel.Invoke(
                Sub() Process_match_Key_set_SU(ds, m_planet),
                Sub() Process_match_Key_set_SU(ds, m_planet),
                Sub() Process_match_Key_set_SU(ds, m_planet),
                Sub() Process_match_Key_set_SU(ds, m_planet),
                Sub() Process_match_Key_set_SU(ds, m_planet),
                Sub() Process_match_Key_set_SU(ds, m_planet),
                Sub() Process_match_Key_set_SU(ds, m_planet),
                Sub() Process_match_Key_set_SU(ds, m_planet),
                Sub() Process_match_Key_set_SU(ds, m_planet),
                Sub() Process_match_Key_set_SU(ds, m_planet),
                Sub() Process_match_Key_set_SU(ds, m_planet),
                Sub() Process_match_Key_set_SU(ds, m_planet),
                Sub() Process_match_Key_set_SU(ds, m_planet),
                Sub() Process_match_Key_set_SU(ds, m_planet),
                Sub() Process_match_Key_set_SU(ds, m_planet),
                Sub() Process_match_Key_set_SU(ds, m_planet),
                Sub() Process_match_Key_set_SU(ds, m_planet),
                Sub() Process_match_Key_set_SU(ds, m_planet),
                Sub() Process_match_Key_set_SU(ds, m_planet),
                Sub() Process_match_Key_set_SU(ds, m_planet),
                Sub() Process_match_Key_set_SU(ds, m_planet),
                Sub() Process_match_Key_set_SU(ds, m_planet),
                Sub() Process_match_Key_set_SU(ds, m_planet),
                Sub() Process_match_Key_set_SU(ds, m_planet),
                Sub() Process_match_Key_set_SU(ds, m_planet),
                Sub() Process_match_Key_set_SU(ds, m_planet),
                Sub() Process_match_Key_set_SU(ds, m_planet),
                Sub() Process_match_Key_set_SU(ds, m_planet),
                Sub() Process_match_Key_set_SU(ds, m_planet),
                Sub() Process_match_Key_set_SU(ds, m_planet),
                Sub() Process_match_Key_set_SU(ds, m_planet),
                Sub() Process_match_Key_set_SU(ds, m_planet),
                Sub() Process_match_Key_set_SU(ds, m_planet),
                Sub() Process_match_Key_set_SU(ds, m_planet),
                Sub() Process_match_Key_set_SU(ds, m_planet),
                Sub() Process_match_Key_set_SU(ds, m_planet),
                Sub() Process_match_Key_set_SU(ds, m_planet),
                Sub() Process_match_Key_set_SU(ds, m_planet),
                Sub() Process_match_Key_set_SU(ds, m_planet),
                Sub() Process_match_Key_set_SU(ds, m_planet))
    End Sub
    Sub Process_match_Key_set_ALL(ByVal RowsData As DataSet, ByRef m_planet As String)
        ThreadStartCounter += 1
        Dim a8(55) As String
        While counter_AL > -1
            Dim i = counter_AL
            counter_AL -= 1
            Try
                Dim pa = m_planet
                Dim p1 = RowsData.Tables(0).Rows(i)(0).Trim.ToString()
                Dim p2 = RowsData.Tables(0).Rows(i)(1).Trim.ToString()
                Dim p3 = RowsData.Tables(0).Rows(i)(2).Trim.ToString()
                Dim p4 = RowsData.Tables(0).Rows(i)(3).Trim.ToString()
                Dim p5 = RowsData.Tables(0).Rows(i)(4).Trim.ToString()
                Dim p6 = RowsData.Tables(0).Rows(i)(5).Trim.ToString()
                Dim m_key = pa + p1 + p2 + p3 + p4 + p5 + p6

                a8(0) = pa
                a8(1) = p1
                a8(2) = p2
                a8(3) = p3
                a8(4) = p4
                a8(5) = p5
                a8(6) = p6

                a8(7) = pa
                a8(8) = p6
                a8(9) = p5
                a8(10) = p4
                a8(11) = p3
                a8(12) = p2
                a8(13) = p1

                a8(14) = p1
                a8(15) = p2
                a8(16) = p3
                a8(17) = p4
                a8(18) = p5
                a8(19) = p6
                a8(20) = pa

                a8(21) = p1
                a8(22) = pa
                a8(23) = p6
                a8(24) = p5
                a8(25) = p4
                a8(26) = p3
                a8(27) = p2

                a8(28) = p2
                a8(29) = p3
                a8(30) = p4
                a8(31) = p5
                a8(32) = p6
                a8(33) = pa
                a8(34) = p1

                a8(35) = p2
                a8(36) = p1
                a8(37) = pa
                a8(38) = p6
                a8(39) = p5
                a8(40) = p4
                a8(41) = p3

                a8(42) = p3
                a8(43) = p4
                a8(44) = p5
                a8(45) = p6
                a8(46) = pa
                a8(47) = p1
                a8(48) = p2

                a8(49) = p3
                a8(50) = p2
                a8(51) = p1
                a8(52) = pa
                a8(53) = p6
                a8(54) = p5
                a8(55) = p4
                Parallel.Invoke(
                Sub() Match_Key(pa, m_key, a8, c01, "01"),
                Sub() Match_Key(pa, m_key, a8, c02, "02"),
                Sub() Match_Key(pa, m_key, a8, c03, "03"),
                Sub() Match_Key(pa, m_key, a8, c04, "04"),
                Sub() Match_Key(pa, m_key, a8, c05, "05"),
                Sub() Match_Key(pa, m_key, a8, c06, "06"),
                Sub() Match_Key(pa, m_key, a8, c07, "07"),
                Sub() Match_Key(pa, m_key, a8, c08, "08"),
                Sub() Match_Key(pa, m_key, a8, c09, "09"),
                Sub() Match_Key(pa, m_key, a8, c10, "10"),
                Sub() Match_Key(pa, m_key, a8, c11, "11"),
                Sub() Match_Key(pa, m_key, a8, c12, "12"))
            Catch ex As Exception
                Exit While
            End Try
        End While
        ThreadEndCounter += 1
    End Sub
    Sub Process_match_Key_set_MA(ByVal RowsData As DataSet, ByRef m_planet As String)
        ThreadStartCounter += 1
        Dim a8(55) As String
        While counter_MA > -1
            Dim i = counter_MA
            counter_MA -= 1
            Try
                Dim pa = m_planet
                Dim p1 = RowsData.Tables(0).Rows(i)(0).Trim.ToString()
                Dim p2 = RowsData.Tables(0).Rows(i)(1).Trim.ToString()
                Dim p3 = RowsData.Tables(0).Rows(i)(2).Trim.ToString()
                Dim p4 = RowsData.Tables(0).Rows(i)(3).Trim.ToString()
                Dim p5 = RowsData.Tables(0).Rows(i)(4).Trim.ToString()
                Dim p6 = RowsData.Tables(0).Rows(i)(5).Trim.ToString()
                Dim m_key = pa + p1 + p2 + p3 + p4 + p5 + p6

                a8(0) = pa
                a8(1) = p1
                a8(2) = p2
                a8(3) = p3
                a8(4) = p4
                a8(5) = p5
                a8(6) = p6

                a8(7) = pa
                a8(8) = p6
                a8(9) = p5
                a8(10) = p4
                a8(11) = p3
                a8(12) = p2
                a8(13) = p1

                a8(14) = p1
                a8(15) = p2
                a8(16) = p3
                a8(17) = p4
                a8(18) = p5
                a8(19) = p6
                a8(20) = pa

                a8(21) = p1
                a8(22) = pa
                a8(23) = p6
                a8(24) = p5
                a8(25) = p4
                a8(26) = p3
                a8(27) = p2

                a8(28) = p2
                a8(29) = p3
                a8(30) = p4
                a8(31) = p5
                a8(32) = p6
                a8(33) = pa
                a8(34) = p1

                a8(35) = p2
                a8(36) = p1
                a8(37) = pa
                a8(38) = p6
                a8(39) = p5
                a8(40) = p4
                a8(41) = p3

                a8(42) = p3
                a8(43) = p4
                a8(44) = p5
                a8(45) = p6
                a8(46) = pa
                a8(47) = p1
                a8(48) = p2

                a8(49) = p3
                a8(50) = p2
                a8(51) = p1
                a8(52) = pa
                a8(53) = p6
                a8(54) = p5
                a8(55) = p4
                Parallel.Invoke(
                Sub() Match_Key(pa, m_key, a8, c01, "01"),
                Sub() Match_Key(pa, m_key, a8, c02, "02"),
                Sub() Match_Key(pa, m_key, a8, c03, "03"),
                Sub() Match_Key(pa, m_key, a8, c04, "04"),
                Sub() Match_Key(pa, m_key, a8, c05, "05"),
                Sub() Match_Key(pa, m_key, a8, c06, "06"),
                Sub() Match_Key(pa, m_key, a8, c07, "07"),
                Sub() Match_Key(pa, m_key, a8, c08, "08"),
                Sub() Match_Key(pa, m_key, a8, c09, "09"),
                Sub() Match_Key(pa, m_key, a8, c10, "10"),
                Sub() Match_Key(pa, m_key, a8, c11, "11"),
                Sub() Match_Key(pa, m_key, a8, c12, "12"))
            Catch ex As Exception
                Exit While
            End Try
        End While
        ThreadEndCounter += 1
    End Sub
    Sub Process_match_Key_set_ME(ByVal RowsData As DataSet, ByRef m_planet As String)
        ThreadStartCounter += 1
        Dim a8(55) As String
        While counter_ME > -1
            Dim i = counter_ME
            counter_ME -= 1
            Try
                Dim pa = m_planet
                Dim p1 = RowsData.Tables(0).Rows(i)(0).Trim.ToString()
                Dim p2 = RowsData.Tables(0).Rows(i)(1).Trim.ToString()
                Dim p3 = RowsData.Tables(0).Rows(i)(2).Trim.ToString()
                Dim p4 = RowsData.Tables(0).Rows(i)(3).Trim.ToString()
                Dim p5 = RowsData.Tables(0).Rows(i)(4).Trim.ToString()
                Dim p6 = RowsData.Tables(0).Rows(i)(5).Trim.ToString()
                Dim m_key = pa + p1 + p2 + p3 + p4 + p5 + p6

                a8(0) = pa
                a8(1) = p1
                a8(2) = p2
                a8(3) = p3
                a8(4) = p4
                a8(5) = p5
                a8(6) = p6

                a8(7) = pa
                a8(8) = p6
                a8(9) = p5
                a8(10) = p4
                a8(11) = p3
                a8(12) = p2
                a8(13) = p1

                a8(14) = p1
                a8(15) = p2
                a8(16) = p3
                a8(17) = p4
                a8(18) = p5
                a8(19) = p6
                a8(20) = pa

                a8(21) = p1
                a8(22) = pa
                a8(23) = p6
                a8(24) = p5
                a8(25) = p4
                a8(26) = p3
                a8(27) = p2

                a8(28) = p2
                a8(29) = p3
                a8(30) = p4
                a8(31) = p5
                a8(32) = p6
                a8(33) = pa
                a8(34) = p1

                a8(35) = p2
                a8(36) = p1
                a8(37) = pa
                a8(38) = p6
                a8(39) = p5
                a8(40) = p4
                a8(41) = p3

                a8(42) = p3
                a8(43) = p4
                a8(44) = p5
                a8(45) = p6
                a8(46) = pa
                a8(47) = p1
                a8(48) = p2

                a8(49) = p3
                a8(50) = p2
                a8(51) = p1
                a8(52) = pa
                a8(53) = p6
                a8(54) = p5
                a8(55) = p4
                Parallel.Invoke(
                Sub() Match_Key(pa, m_key, a8, c01, "01"),
                Sub() Match_Key(pa, m_key, a8, c02, "02"),
                Sub() Match_Key(pa, m_key, a8, c03, "03"),
                Sub() Match_Key(pa, m_key, a8, c04, "04"),
                Sub() Match_Key(pa, m_key, a8, c05, "05"),
                Sub() Match_Key(pa, m_key, a8, c06, "06"),
                Sub() Match_Key(pa, m_key, a8, c07, "07"),
                Sub() Match_Key(pa, m_key, a8, c08, "08"),
                Sub() Match_Key(pa, m_key, a8, c09, "09"),
                Sub() Match_Key(pa, m_key, a8, c10, "10"),
                Sub() Match_Key(pa, m_key, a8, c11, "11"),
                Sub() Match_Key(pa, m_key, a8, c12, "12"))
            Catch ex As Exception
                Exit While
            End Try
        End While
        ThreadEndCounter += 1
    End Sub
    Sub Process_match_Key_set_MO(ByVal RowsData As DataSet, ByRef m_planet As String)
        ThreadStartCounter += 1
        Dim a8(55) As String
        While counter_MO > -1
            Dim i = counter_MO
            counter_MO -= 1
            Try
                Dim pa = m_planet
                Dim p1 = RowsData.Tables(0).Rows(i)(0).Trim.ToString()
                Dim p2 = RowsData.Tables(0).Rows(i)(1).Trim.ToString()
                Dim p3 = RowsData.Tables(0).Rows(i)(2).Trim.ToString()
                Dim p4 = RowsData.Tables(0).Rows(i)(3).Trim.ToString()
                Dim p5 = RowsData.Tables(0).Rows(i)(4).Trim.ToString()
                Dim p6 = RowsData.Tables(0).Rows(i)(5).Trim.ToString()
                Dim m_key = pa + p1 + p2 + p3 + p4 + p5 + p6

                a8(0) = pa
                a8(1) = p1
                a8(2) = p2
                a8(3) = p3
                a8(4) = p4
                a8(5) = p5
                a8(6) = p6

                a8(7) = pa
                a8(8) = p6
                a8(9) = p5
                a8(10) = p4
                a8(11) = p3
                a8(12) = p2
                a8(13) = p1

                a8(14) = p1
                a8(15) = p2
                a8(16) = p3
                a8(17) = p4
                a8(18) = p5
                a8(19) = p6
                a8(20) = pa

                a8(21) = p1
                a8(22) = pa
                a8(23) = p6
                a8(24) = p5
                a8(25) = p4
                a8(26) = p3
                a8(27) = p2

                a8(28) = p2
                a8(29) = p3
                a8(30) = p4
                a8(31) = p5
                a8(32) = p6
                a8(33) = pa
                a8(34) = p1

                a8(35) = p2
                a8(36) = p1
                a8(37) = pa
                a8(38) = p6
                a8(39) = p5
                a8(40) = p4
                a8(41) = p3

                a8(42) = p3
                a8(43) = p4
                a8(44) = p5
                a8(45) = p6
                a8(46) = pa
                a8(47) = p1
                a8(48) = p2

                a8(49) = p3
                a8(50) = p2
                a8(51) = p1
                a8(52) = pa
                a8(53) = p6
                a8(54) = p5
                a8(55) = p4
                Parallel.Invoke(
                Sub() Match_Key(pa, m_key, a8, c01, "01"),
                Sub() Match_Key(pa, m_key, a8, c02, "02"),
                Sub() Match_Key(pa, m_key, a8, c03, "03"),
                Sub() Match_Key(pa, m_key, a8, c04, "04"),
                Sub() Match_Key(pa, m_key, a8, c05, "05"),
                Sub() Match_Key(pa, m_key, a8, c06, "06"),
                Sub() Match_Key(pa, m_key, a8, c07, "07"),
                Sub() Match_Key(pa, m_key, a8, c08, "08"),
                Sub() Match_Key(pa, m_key, a8, c09, "09"),
                Sub() Match_Key(pa, m_key, a8, c10, "10"),
                Sub() Match_Key(pa, m_key, a8, c11, "11"),
                Sub() Match_Key(pa, m_key, a8, c12, "12"))
            Catch ex As Exception
                Exit While
            End Try
        End While
        ThreadEndCounter += 1
    End Sub
    Sub Process_match_Key_set_JU(ByVal RowsData As DataSet, ByRef m_planet As String)
        ThreadStartCounter += 1
        Dim a8(55) As String
        While counter_JU > -1
            Dim i = counter_JU
            counter_JU -= 1
            Try
                Dim pa = m_planet
                Dim p1 = RowsData.Tables(0).Rows(i)(0).Trim.ToString()
                Dim p2 = RowsData.Tables(0).Rows(i)(1).Trim.ToString()
                Dim p3 = RowsData.Tables(0).Rows(i)(2).Trim.ToString()
                Dim p4 = RowsData.Tables(0).Rows(i)(3).Trim.ToString()
                Dim p5 = RowsData.Tables(0).Rows(i)(4).Trim.ToString()
                Dim p6 = RowsData.Tables(0).Rows(i)(5).Trim.ToString()
                Dim m_key = pa + p1 + p2 + p3 + p4 + p5 + p6

                a8(0) = pa
                a8(1) = p1
                a8(2) = p2
                a8(3) = p3
                a8(4) = p4
                a8(5) = p5
                a8(6) = p6

                a8(7) = pa
                a8(8) = p6
                a8(9) = p5
                a8(10) = p4
                a8(11) = p3
                a8(12) = p2
                a8(13) = p1

                a8(14) = p1
                a8(15) = p2
                a8(16) = p3
                a8(17) = p4
                a8(18) = p5
                a8(19) = p6
                a8(20) = pa

                a8(21) = p1
                a8(22) = pa
                a8(23) = p6
                a8(24) = p5
                a8(25) = p4
                a8(26) = p3
                a8(27) = p2

                a8(28) = p2
                a8(29) = p3
                a8(30) = p4
                a8(31) = p5
                a8(32) = p6
                a8(33) = pa
                a8(34) = p1

                a8(35) = p2
                a8(36) = p1
                a8(37) = pa
                a8(38) = p6
                a8(39) = p5
                a8(40) = p4
                a8(41) = p3

                a8(42) = p3
                a8(43) = p4
                a8(44) = p5
                a8(45) = p6
                a8(46) = pa
                a8(47) = p1
                a8(48) = p2

                a8(49) = p3
                a8(50) = p2
                a8(51) = p1
                a8(52) = pa
                a8(53) = p6
                a8(54) = p5
                a8(55) = p4
                Parallel.Invoke(
                Sub() Match_Key(pa, m_key, a8, c01, "01"),
                Sub() Match_Key(pa, m_key, a8, c02, "02"),
                Sub() Match_Key(pa, m_key, a8, c03, "03"),
                Sub() Match_Key(pa, m_key, a8, c04, "04"),
                Sub() Match_Key(pa, m_key, a8, c05, "05"),
                Sub() Match_Key(pa, m_key, a8, c06, "06"),
                Sub() Match_Key(pa, m_key, a8, c07, "07"),
                Sub() Match_Key(pa, m_key, a8, c08, "08"),
                Sub() Match_Key(pa, m_key, a8, c09, "09"),
                Sub() Match_Key(pa, m_key, a8, c10, "10"),
                Sub() Match_Key(pa, m_key, a8, c11, "11"),
                Sub() Match_Key(pa, m_key, a8, c12, "12"))
            Catch ex As Exception
                Exit While
            End Try
        End While
        ThreadEndCounter += 1
    End Sub
    Sub Process_match_Key_set_SA(ByVal RowsData As DataSet, ByRef m_planet As String)
        ThreadStartCounter += 1
        Dim a8(55) As String
        While counter_SA > -1
            Dim i = counter_SA
            counter_SA -= 1
            Try
                Dim pa = m_planet
                Dim p1 = RowsData.Tables(0).Rows(i)(0).Trim.ToString()
                Dim p2 = RowsData.Tables(0).Rows(i)(1).Trim.ToString()
                Dim p3 = RowsData.Tables(0).Rows(i)(2).Trim.ToString()
                Dim p4 = RowsData.Tables(0).Rows(i)(3).Trim.ToString()
                Dim p5 = RowsData.Tables(0).Rows(i)(4).Trim.ToString()
                Dim p6 = RowsData.Tables(0).Rows(i)(5).Trim.ToString()
                Dim m_key = pa + p1 + p2 + p3 + p4 + p5 + p6

                a8(0) = pa
                a8(1) = p1
                a8(2) = p2
                a8(3) = p3
                a8(4) = p4
                a8(5) = p5
                a8(6) = p6

                a8(7) = pa
                a8(8) = p6
                a8(9) = p5
                a8(10) = p4
                a8(11) = p3
                a8(12) = p2
                a8(13) = p1

                a8(14) = p1
                a8(15) = p2
                a8(16) = p3
                a8(17) = p4
                a8(18) = p5
                a8(19) = p6
                a8(20) = pa

                a8(21) = p1
                a8(22) = pa
                a8(23) = p6
                a8(24) = p5
                a8(25) = p4
                a8(26) = p3
                a8(27) = p2

                a8(28) = p2
                a8(29) = p3
                a8(30) = p4
                a8(31) = p5
                a8(32) = p6
                a8(33) = pa
                a8(34) = p1

                a8(35) = p2
                a8(36) = p1
                a8(37) = pa
                a8(38) = p6
                a8(39) = p5
                a8(40) = p4
                a8(41) = p3

                a8(42) = p3
                a8(43) = p4
                a8(44) = p5
                a8(45) = p6
                a8(46) = pa
                a8(47) = p1
                a8(48) = p2

                a8(49) = p3
                a8(50) = p2
                a8(51) = p1
                a8(52) = pa
                a8(53) = p6
                a8(54) = p5
                a8(55) = p4
                Parallel.Invoke(
                Sub() Match_Key(pa, m_key, a8, c01, "01"),
                Sub() Match_Key(pa, m_key, a8, c02, "02"),
                Sub() Match_Key(pa, m_key, a8, c03, "03"),
                Sub() Match_Key(pa, m_key, a8, c04, "04"),
                Sub() Match_Key(pa, m_key, a8, c05, "05"),
                Sub() Match_Key(pa, m_key, a8, c06, "06"),
                Sub() Match_Key(pa, m_key, a8, c07, "07"),
                Sub() Match_Key(pa, m_key, a8, c08, "08"),
                Sub() Match_Key(pa, m_key, a8, c09, "09"),
                Sub() Match_Key(pa, m_key, a8, c10, "10"),
                Sub() Match_Key(pa, m_key, a8, c11, "11"),
                Sub() Match_Key(pa, m_key, a8, c12, "12"))
            Catch ex As Exception
                Exit While
            End Try
        End While
        ThreadEndCounter += 1
    End Sub
    Sub Process_match_Key_set_SU(ByVal RowsData As DataSet, ByRef m_planet As String)
        ThreadStartCounter += 1
        Dim a8(55) As String
        While counter_SU > -1
            Dim i = counter_SU
            counter_SU -= 1
            Try
                Dim pa = m_planet
                Dim p1 = RowsData.Tables(0).Rows(i)(0).Trim.ToString()
                Dim p2 = RowsData.Tables(0).Rows(i)(1).Trim.ToString()
                Dim p3 = RowsData.Tables(0).Rows(i)(2).Trim.ToString()
                Dim p4 = RowsData.Tables(0).Rows(i)(3).Trim.ToString()
                Dim p5 = RowsData.Tables(0).Rows(i)(4).Trim.ToString()
                Dim p6 = RowsData.Tables(0).Rows(i)(5).Trim.ToString()
                Dim m_key = pa + p1 + p2 + p3 + p4 + p5 + p6

                a8(0) = pa
                a8(1) = p1
                a8(2) = p2
                a8(3) = p3
                a8(4) = p4
                a8(5) = p5
                a8(6) = p6

                a8(7) = pa
                a8(8) = p6
                a8(9) = p5
                a8(10) = p4
                a8(11) = p3
                a8(12) = p2
                a8(13) = p1

                a8(14) = p1
                a8(15) = p2
                a8(16) = p3
                a8(17) = p4
                a8(18) = p5
                a8(19) = p6
                a8(20) = pa

                a8(21) = p1
                a8(22) = pa
                a8(23) = p6
                a8(24) = p5
                a8(25) = p4
                a8(26) = p3
                a8(27) = p2

                a8(28) = p2
                a8(29) = p3
                a8(30) = p4
                a8(31) = p5
                a8(32) = p6
                a8(33) = pa
                a8(34) = p1

                a8(35) = p2
                a8(36) = p1
                a8(37) = pa
                a8(38) = p6
                a8(39) = p5
                a8(40) = p4
                a8(41) = p3

                a8(42) = p3
                a8(43) = p4
                a8(44) = p5
                a8(45) = p6
                a8(46) = pa
                a8(47) = p1
                a8(48) = p2

                a8(49) = p3
                a8(50) = p2
                a8(51) = p1
                a8(52) = pa
                a8(53) = p6
                a8(54) = p5
                a8(55) = p4
                Parallel.Invoke(
                Sub() Match_Key(pa, m_key, a8, c01, "01"),
                Sub() Match_Key(pa, m_key, a8, c02, "02"),
                Sub() Match_Key(pa, m_key, a8, c03, "03"),
                Sub() Match_Key(pa, m_key, a8, c04, "04"),
                Sub() Match_Key(pa, m_key, a8, c05, "05"),
                Sub() Match_Key(pa, m_key, a8, c06, "06"),
                Sub() Match_Key(pa, m_key, a8, c07, "07"),
                Sub() Match_Key(pa, m_key, a8, c08, "08"),
                Sub() Match_Key(pa, m_key, a8, c09, "09"),
                Sub() Match_Key(pa, m_key, a8, c10, "10"),
                Sub() Match_Key(pa, m_key, a8, c11, "11"),
                Sub() Match_Key(pa, m_key, a8, c12, "12"))
            Catch ex As Exception
                Exit While
            End Try
        End While
        ThreadEndCounter += 1
    End Sub
    Sub Process_match_Key_set_VE(ByVal RowsData As DataSet, ByRef m_planet As String)
        ThreadStartCounter += 1
        Dim a8(55) As String
        While counter_VE > -1
            Dim i = counter_VE
            counter_VE -= 1
            Try
                Dim pa = m_planet
                Dim p1 = RowsData.Tables(0).Rows(i)(0).Trim.ToString()
                Dim p2 = RowsData.Tables(0).Rows(i)(1).Trim.ToString()
                Dim p3 = RowsData.Tables(0).Rows(i)(2).Trim.ToString()
                Dim p4 = RowsData.Tables(0).Rows(i)(3).Trim.ToString()
                Dim p5 = RowsData.Tables(0).Rows(i)(4).Trim.ToString()
                Dim p6 = RowsData.Tables(0).Rows(i)(5).Trim.ToString()
                Dim m_key = pa + p1 + p2 + p3 + p4 + p5 + p6

                a8(0) = pa
                a8(1) = p1
                a8(2) = p2
                a8(3) = p3
                a8(4) = p4
                a8(5) = p5
                a8(6) = p6

                a8(7) = pa
                a8(8) = p6
                a8(9) = p5
                a8(10) = p4
                a8(11) = p3
                a8(12) = p2
                a8(13) = p1

                a8(14) = p1
                a8(15) = p2
                a8(16) = p3
                a8(17) = p4
                a8(18) = p5
                a8(19) = p6
                a8(20) = pa

                a8(21) = p1
                a8(22) = pa
                a8(23) = p6
                a8(24) = p5
                a8(25) = p4
                a8(26) = p3
                a8(27) = p2

                a8(28) = p2
                a8(29) = p3
                a8(30) = p4
                a8(31) = p5
                a8(32) = p6
                a8(33) = pa
                a8(34) = p1

                a8(35) = p2
                a8(36) = p1
                a8(37) = pa
                a8(38) = p6
                a8(39) = p5
                a8(40) = p4
                a8(41) = p3

                a8(42) = p3
                a8(43) = p4
                a8(44) = p5
                a8(45) = p6
                a8(46) = pa
                a8(47) = p1
                a8(48) = p2

                a8(49) = p3
                a8(50) = p2
                a8(51) = p1
                a8(52) = pa
                a8(53) = p6
                a8(54) = p5
                a8(55) = p4
                Parallel.Invoke(
                Sub() Match_Key(pa, m_key, a8, c01, "01"),
                Sub() Match_Key(pa, m_key, a8, c02, "02"),
                Sub() Match_Key(pa, m_key, a8, c03, "03"),
                Sub() Match_Key(pa, m_key, a8, c04, "04"),
                Sub() Match_Key(pa, m_key, a8, c05, "05"),
                Sub() Match_Key(pa, m_key, a8, c06, "06"),
                Sub() Match_Key(pa, m_key, a8, c07, "07"),
                Sub() Match_Key(pa, m_key, a8, c08, "08"),
                Sub() Match_Key(pa, m_key, a8, c09, "09"),
                Sub() Match_Key(pa, m_key, a8, c10, "10"),
                Sub() Match_Key(pa, m_key, a8, c11, "11"),
                Sub() Match_Key(pa, m_key, a8, c12, "12"))
            Catch ex As Exception
                Exit While
            End Try
        End While
        ThreadEndCounter += 1
    End Sub
    Sub Match_Key(ByRef m_planet As String, ByVal m_key As String, ByVal combo() As String, ByVal Cuspp() As String, ByVal cloc As String)
        Dim y2() As String = {"YY", "YY", "YY", "YY", "YY", "YY", "YY", "YY", "YY", "YY", "YY", "YY", "YY", "YY", "YY", "YY", "YY", "YY", "YY", "YY", "YY", "YY", "YY", "YY", "YY", "YY", "YY", "YY", "YY", "YY", "YY", "YY", "YY", "YY", "YY", "YY", "YY", "YY", "YY", "YY", "YY", "YY", "YY", "YY", "YY", "YY", "YY", "YY", "YY", "YY", "YY", "YY", "YY", "YY", "YY", "YY"}
        Dim c = Cuspp.Length - 1
        Dim z1(c) As String
        Dim mstr = ""
        For i As Integer = 0 To 55
            For j As Integer = 0 To c
                If combo(i) = Cuspp(j) Then
                    y2(i) = "XX"
                End If
            Next
        Next
        For mega As Integer = 0 To 7
            Dim pstr2 = ""
            mstr = ""
            Array.Copy(Cuspp, z1, c + 1)
            Dim start = (mega) * 7
            Dim finish = start + 6
            For i = start To finish
                If y2(i) IsNot "XX" Then
                    Exit For
                End If
                For j = 0 To c
                    If z1(j) = combo(i) Then
                        z1(j) = "XX"
                        mstr = mstr + combo(i)
                        Exit For
                    Else
                        If Cuspp(j) = combo(i) Then
                            mstr = mstr + combo(i)
                        End If
                    End If
                Next
            Next
            Dim pattern As Boolean = True
            For m As Integer = 0 To c
                If z1(m) IsNot "XX" Then
                    pattern = False
                End If
            Next
            If pattern = True Then
                For i As Integer = 0 To c
                    pstr2 = pstr2 + Cuspp(i)
                Next
                Dim uid As String = "XXXXXXXXXX"
                Dim hid As String = "100001"
                Dim con As New SqlConnection
                Dim cmd As New SqlCommand
                Try
                    con.ConnectionString = connstr
                    con.Open()
                    cmd.Connection = con
                    cmd.CommandText = "INSERT INTO MATCH_FILE VALUES ('" + uid + "','" + hid + "','" + m_planet + "','" + m_key + "','" + cloc + "','" + pstr2 + "');"
                    cmd.ExecuteNonQuery()
                Catch ex As Exception
                    duplicates += 1
                Finally
                    con.Close()
                End Try
            End If
        Next
    End Sub
    Sub NumberOfRecordsInMatchFile()
        Dim con As New SqlConnection(connstr)
        Dim connection As SqlConnection = New SqlConnection(connstr)
        connection.Open()
        Dim cmd As New SqlCommand("SELECT * FROM MATCH_FILE;", con)
        Dim da As New SqlDataAdapter(cmd)
        Dim ds As New DataSet()
        da.Fill(ds)
        connection.Close()
    End Sub
    Sub TruncateMatch_File()
        Dim con As New SqlConnection
        Dim cmd As New SqlCommand
        Try
            con.ConnectionString = connstr
            con.Open()
            cmd.Connection = con
            cmd.CommandText = "TRUNCATE TABLE MATCH_FILE;"
            cmd.ExecuteNonQuery()
        Catch ex As Exception
        Finally
            con.Close()
        End Try
    End Sub

End Class
