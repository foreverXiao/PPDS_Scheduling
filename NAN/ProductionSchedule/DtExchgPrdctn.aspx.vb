Imports System
Imports System.Configuration
Imports System.Data
Imports System.Data.OleDb
Imports Microsoft.VisualBasic
Imports System.Diagnostics
Imports System.Runtime.InteropServices
Imports System.IO
Imports System.Globalization
Imports System.Threading
Imports System.Net



Partial Class ProductionSchedule_DtExchgPrdctn
    Inherits basepage1

    Public Const excelDatabaseFile = "schedule"
    Private Const Separator1 = "#"

    Protected Sub Page_Init1(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Init
        accesslevel = 2 ' everyone could access this page
    End Sub

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Response.Clear()
        Dim query1 As String, connstr As String, actionRequested As String
        Dim startD As Date = CDate(Today).AddDays(-1)

        Dim daysFromStart As Single = 3
        Dim pixelsPerDay As Long = 180
        Dim sndbck As System.Text.StringBuilder = New System.Text.StringBuilder()



        'Considering special case in Nansha, its sorting is by dat_finish_date instead of dat_start_date like Shanghai and Chongqing
        Dim DAT_START_DATE As String = "dat_start_date", DAT_FINISH_DATE As String = "dat_finish_date"
        If valueOf("strGanttStyle").ToUpper.StartsWith("SORTBYFINISHTIME") Then
            DAT_START_DATE = "dat_finish_date"
            DAT_FINISH_DATE = "dat_start_date"
        End If

        actionRequested = Request.Params("action")



        Dim temp As New NameValueCollection

        If Cache("production") Is Nothing Then 'need refresh data for planning period parameter

            Dim connParam As OleDbConnection = New OleDbConnection(ConfigurationManager.ConnectionStrings(dbConnForParam).ProviderName & ConfigurationManager.ConnectionStrings(dbConnForParam).ConnectionString)
            connParam.Open()


            'get the lastest data from table Esch_Na_tbl_plnprmtr to preset some planning parameter

            Dim command1 As New OleDbCommand("Select paramname, paramvalue From Esch_Na_tbl_plnprmtr where category = 'production'", connParam)
            Dim reader1 As OleDbDataReader
            reader1 = command1.ExecuteReader()
            Do While reader1.Read()
                temp.Add(reader1("paramname").ToString, reader1("paramvalue").ToString)
            Loop

            reader1.Close()

            command1.Dispose()

            connParam.Close()
            connParam.Dispose()


            Cache.Insert("production", temp, Nothing, Cache.NoAbsoluteExpiration, TimeSpan.FromMinutes(5))

        End If


        temp = CType(Cache("production"), NameValueCollection)

        If temp("autoStart").ToLower = "no" Then
            startD = CDate(temp("starttime"))
        End If
        daysFromStart = CSng(temp("daystobeshown"))



        'Get production lines' information ===
        If actionRequested = "productionlines" Then

            Dim arrayLine1 As Integer() = arrayOfLines()

            Dim i As Integer = arrayLine1.Count
            Dim j1 As Integer = 0


            sndbck.Append("<table border='1' ><tr>")
            For j1 = 0 To (i - 1)
                If Not (arrayLine1(j1) = valueOf("intDummyLine")) Then  'ignore the dummy production line # 333
                    sndbck.Append("<td onclick='dsplyOrderByLine(this);' class='linemenu' id='" & arrayLine1(j1) & "'>" & "L" & arrayLine1(j1) & "</td>")
                End If
            Next
            sndbck.Append("</tr></table>")

        End If

        'provide content for scalar bar
        If actionRequested = "scalar" Then
            For j1 = 1 To 30
                sndbck.Append("<div class='sc' style='text-align:center;width:" & (pixelsPerDay - 1) & "px;left:" & (pixelsPerDay * (j1 - 1)) & "px;" & IIf(j1 Mod 2 = 1, "background-color:#00FFFF;", "") & "'>" & DateAdd("d", j1 - 1, startD).ToString("ddd  MMM d") & "</div>")
            Next
        End If

        'provide background for background mark lines
        If actionRequested = "bckgrndImg" Then
            For j1 = 1 To 30
                sndbck.Append("<div class='bg' style='width:" & (pixelsPerDay - 1) & "px;left:" & (pixelsPerDay * (j1 - 1)) & "px;'></div>")
            Next
        End If



        'generate excel file as the database to show gantt chart
        If actionRequested = "generateExcelFile" Then
            If True OrElse CacheFrom("sso") Is Nothing Then
                'sndbck.Append("../usermanage/loginBySSO.aspx?oldurl=../dragDrop/GanttChart.aspx")
                'Else
                Dim counter1 As Integer
                listOutDifferencesBetweenScheduleUpdates(counter1)

                connstr = ConfigurationManager.ConnectionStrings(dbConnectionName).ProviderName & ConfigurationManager.ConnectionStrings(dbConnectionName).ConnectionString
                Dim conn As OleDbConnection = New OleDbConnection(connstr)

                'query1 = "Select table1.txt_order_key,int_line_no,dat_start_date,dat_finish_date,txt_lot_no,int_span,planned_production_qty,flt_unallocate_qty,flt_order_qty,txt_color,txt_item_no,dat_etd,txt_VIP,txt_process_technics,txt_FDA,txt_remark,txt_payment_status,table2.txt_pull_screw From  " & _
                ' " ((Select txt_order_key,int_line_no,dat_start_date,dat_finish_date,txt_lot_no,int_span,planned_production_qty,flt_unallocate_qty,flt_order_qty,txt_color,txt_item_no,dat_etd,txt_VIP,txt_process_technics,txt_FDA,txt_remark,txt_payment_status From Esch_Na_tbl_orders  Where CAST(int_line_no AS VARCHAR(5)) <> '" & valueOf("intDummyLine") & "' And  ((dat_start_date  between " & dateSeparator & startD & dateSeparator & " And " & dateSeparator & DateAdd("h", daysFromStart * 24, startD) & dateSeparator & ") Or (dat_finish_date  between " & dateSeparator & startD & dateSeparator & " And " & dateSeparator & DateAdd("h", daysFromStart * 24, startD) & dateSeparator & "))) as table1 Left Join (select txt_order_key,txt_pull_screw From Esch_Na_tbl_prductn_prmtr) as table2 On (table1.txt_order_key = table2.txt_order_key)) Order by int_line_no,table1.dat_finish_date ASC, table1.dat_start_date ASC "

                'select orders between start time and finish time, in case some consecutive lots,  its finish time is far late than the set time 
                Dim noDummyLineAndSpecificStartFinishTime As String = " CAST(int_line_no AS VARCHAR(5)) <> '" & valueOf("intDummyLine") & "' And  ((dat_start_date  between " & dateSeparator & startD & dateSeparator & " And " & dateSeparator & DateAdd("h", daysFromStart * 24, startD) & dateSeparator & ") Or (dat_finish_date  between " & dateSeparator & startD & dateSeparator & " And " & dateSeparator & DateAdd("h", daysFromStart * 24, startD) & dateSeparator & "))"
                Dim lotList As String = " SELECT DISTINCT txt_lot_no FROM Esch_Na_tbl_orders  Where NOT (txt_lot_no IS NULL) AND txt_lot_no <> '' AND " & noDummyLineAndSpecificStartFinishTime
                query1 = " Select txt_order_key,int_line_no,dat_start_date,dat_finish_date,txt_lot_no,int_span,planned_production_qty,flt_unallocate_qty,flt_order_qty,txt_color,txt_item_no,dat_etd,txt_VIP,txt_process_technics,txt_FDA,txt_remark,txt_payment_status FROM  Esch_Na_tbl_orders WHERE ( CAST(int_line_no AS VARCHAR(5)) <> '" & valueOf("intDummyLine") & "' AND txt_lot_no IN (" & lotList & ")) OR " & noDummyLineAndSpecificStartFinishTime
                query1 = " SELECT table1.txt_order_key,int_line_no,dat_start_date,dat_finish_date,txt_lot_no,int_span,planned_production_qty,flt_unallocate_qty,flt_order_qty,txt_color,txt_item_no,dat_etd,txt_VIP,txt_process_technics,txt_FDA,txt_remark,txt_payment_status,txt_pull_screw FROM ( (" & query1 & ") AS table1 LEFT JOIN Esch_Na_tbl_prductn_prmtr ON (table1.txt_order_key = Esch_Na_tbl_prductn_prmtr.txt_order_key)) Order by int_line_no, " & DAT_START_DATE & "  ASC, " & DAT_FINISH_DATE & " ASC"


                Dim dtUpdateFrom As OleDbDataAdapter = New OleDbDataAdapter(query1, conn)
                Dim dtFromTable As New DataTable()
                dtUpdateFrom.Fill(dtFromTable)

                Dim eScheduleFile As String
                Dim useTextMethodToPostSchedule As Boolean = True
                If Not String.IsNullOrEmpty(valueOf("bnlPostScheduleViaURL")) AndAlso CBool(valueOf("bnlPostScheduleViaURL")) Then
                    useTextMethodToPostSchedule = False
                End If



                Dim filePath As String = ConfigurationManager.AppSettings("excelFolder")
                Dim excelFileName As String = generateExcel(Nothing, dtFromTable, filePath, startD, startD.AddDays(CInt(daysFromStart) + 1), pixelsPerDay, eScheduleFile, useTextMethodToPostSchedule)

                dtFromTable.Dispose()
                dtUpdateFrom.Dispose()

                sndbck.Append(counter1)


                If False AndAlso useTextMethodToPostSchedule Then 'not for Nansha

                    Dim fileLocation As String = valueOf("strWebScheduleLocation")
                    If String.IsNullOrEmpty(fileLocation) Then
                        fileLocation = filePath
                        If fileLocation.EndsWith("\") Then
                            fileLocation = fileLocation.Remove(fileLocation.Length - 1, 1)
                        End If
                    End If

                    Response.AddHeader("Content-Disposition", "attachment; filename=" & eScheduleFile)
                    Response.TransmitFile(fileLocation & "\" & eScheduleFile)
                    Response.End()

                End If

            End If

        End If



        If actionRequested = "fileTime" Then
            Dim scheduleFileTime As DateTime = Now()
            Dim filePath As String = ConfigurationManager.AppSettings("excelFolder")
            If Not filePath.EndsWith("\") Then
                filePath &= "\"
            End If

            If File.Exists(filePath & excelDatabaseFile & ".xlsx") Then
                scheduleFileTime = File.GetLastWriteTime(filePath & excelDatabaseFile & ".xlsx")
            Else

                If File.Exists(filePath & excelDatabaseFile & ".xls") Then
                    scheduleFileTime = File.GetLastWriteTime(filePath & excelDatabaseFile & ".xlsx")
                End If
            End If

            sndbck.Append("===" & scheduleFileTime.ToString("ddd MMM d HH:mm") & "===")
        End If



        If actionRequested = "orderlines" Then

            Dim excelconnectionstr As String
            Dim continueOrNot = True


            Dim filePath As String = ConfigurationManager.AppSettings("excelFolder")
            If Not filePath.EndsWith("\") Then
                filePath &= "\"
            End If

            If File.Exists(filePath & excelDatabaseFile & ".xlsx") Then
                excelconnectionstr = String.Format("provider=Microsoft.ACE.OLEDB.12.0; Data Source='{0}';" & "Extended Properties='Excel 12.0 Xml;HDR=YES;IMEX=1'", filePath & excelDatabaseFile & ".xlsx")
            Else
                If File.Exists(filePath & excelDatabaseFile & ".xls") Then
                    excelconnectionstr = String.Format("provider=Microsoft.Jet.OLEDB.4.0; Data Source='{0}';" & "Extended Properties='Excel 8.0;HDR=YES;IMEX=1'", filePath & excelDatabaseFile & ".xls")
                Else
                    actionRequested = "orderlines1"
                    continueOrNot = False
                End If
            End If

            If continueOrNot Then

                Dim connexcl As New System.Data.OleDb.OleDbConnection(excelconnectionstr)
                connexcl.Open()

                Dim dtADPexcel As OleDbDataAdapter = New OleDbDataAdapter("SELECT * FROM [Sheet1$] ", connexcl)



                query1 = "SELECT * FROM [Sheet1$] WHERE CSTR(int_line_no) = '" & CInt(Request.Params("lineno")) & "' "

                Dim command1 As New OleDbCommand(query1, connexcl)
                Dim reader1 As OleDbDataReader
                Try
                    reader1 = command1.ExecuteReader()

                    Dim pullScrew As String = String.Empty
                    Dim marginTop As String = String.Empty
                    Dim txt_process_technics As String = String.Empty
                    Dim SPANandETD As String = String.Empty, leftPosForRSDspan As String = String.Empty


                    Do While reader1.Read()

                        If DBNull.Value.Equals(reader1("txt_pull_screw")) OrElse String.IsNullOrEmpty(reader1("txt_pull_screw")) Then
                            pullScrew = String.Empty
                            marginTop = String.Empty
                        Else
                            'additionGap += 3 * 6
                            If String.Equals(reader1("txt_pull_screw"), "yellow", StringComparison.OrdinalIgnoreCase) Then
                                pullScrew = "<span style='color:#FFD700;font-size:150%;'>►</span>"
                            Else
                                pullScrew = "<span style='color:red;font-size:150%;'>►</span>"
                            End If
                            marginTop = "margin-top:-0.22em;"

                        End If

                        If DBNull.Value.Equals(reader1("txt_process_technics")) OrElse String.IsNullOrEmpty(reader1("txt_process_technics")) Then
                            txt_process_technics = String.Empty
                        Else
                            txt_process_technics = reader1("txt_process_technics")
                            txt_process_technics = txt_process_technics.Remove(txt_process_technics.IndexOf(" "))
                        End If

                        SPANandETD = reader1("int_span") & " (" & reader1("dat_etd") & ")"
                        leftPosForRSDspan = "left:-" & (SPANandETD.Length + 0) * 8 & "px;"


                        sndbck.Append("<div class='g-" & reader1("txt_color").ToString() & "' style='font-size:16px;white-space: nowrap;position:relative;border:1px solid black;margin:5px 0px 5px 0px;left:" & CLng(DateDiff("n", startD, reader1("dat_start_date")) * pixelsPerDay / (24 * 60)) & "px;height:20px;width:" & (CLng(DateDiff("n", reader1("dat_start_date"), reader1("dat_finish_date")) * pixelsPerDay / (24 * 60)) - 2) & "px;' id='" & reader1("txt_order_key") & "'><span style='" & leftPosForRSDspan & "position:relative;'>" & SPANandETD & "&nbsp;" & "</span><span style='color:red;position:absolute;white-space:pre;left:" & (CLng(DateDiff("n", reader1("dat_start_date"), reader1("dat_finish_date")) * pixelsPerDay / (24 * 60)) - 3) & "px;" & marginTop & "'>" & pullScrew & "</span>" &
                                      "<span style='position:absolute;white-space:pre;left:" & (CLng(DateDiff("n", reader1("dat_start_date"), reader1("dat_finish_date")) * pixelsPerDay / (24 * 60)) + 20) & "px'>" & reader1("txt_lot_no") & "&nbsp;&nbsp;&nbsp;" & reader1("txt_item_no") & " " & reader1("planned_production_qty") & "KG (screw:" & txt_process_technics & ")&nbsp;&nbsp;" & reader1("txt_FDA") & "&nbsp;" & reader1("txt_VIP") & "(" & reader1("txt_order_key") & ") " & reader1("txt_payment_status") & " " & reader1("txt_remark") & "</span></div>")
                    Loop

                Catch ex As Exception
                    sndbck.Append("<div style='color:red'>" & ex.Message & "</div>")

                Finally
                    reader1.Close()
                End Try

                command1.Dispose()

                dtADPexcel.Dispose()

                If connexcl IsNot Nothing Then
                    connexcl.Close()
                    connexcl.Dispose()
                End If


            End If


        End If



        Response.Write(sndbck.ToString())
        Response.End()


    End Sub



    ''' <summary>
    ''' do comparison between posted schedule and the current schedule in Access database
    ''' </summary>
    ''' <remarks></remarks>
    Public Sub listOutDifferencesBetweenScheduleUpdates(ByRef counter As Integer)

        'get data from schedule Excel file
        'Dim excelDatabaseFile As String = "schedule"  'by reference to page productionSchedule/DtExchgPrdctn.aspx
        Dim excelconnectionstr As String
        Dim continueOrNot = True
        Dim dtTableExcel As DataTable = New DataTable

        Dim filePath As String = ConfigurationManager.AppSettings("excelFolder")
        If Not filePath.EndsWith("\") Then
            filePath &= "\"
        End If

        If File.Exists(filePath & excelDatabaseFile & ".xlsx") Then
            excelconnectionstr = String.Format("provider=Microsoft.ACE.OLEDB.12.0; Data Source='{0}';" & "Extended Properties='Excel 12.0 Xml;HDR=YES;IMEX=1'", filePath & excelDatabaseFile & ".xlsx")
        Else
            If File.Exists(filePath & excelDatabaseFile & ".xls") Then
                excelconnectionstr = String.Format("provider=Microsoft.Jet.OLEDB.4.0; Data Source='{0}';" & "Extended Properties='Excel 8.0;HDR=YES;IMEX=1'", filePath & excelDatabaseFile & ".xls")
            Else
                continueOrNot = False
            End If
        End If

        If continueOrNot Then

            Dim connexcl As New System.Data.OleDb.OleDbConnection(excelconnectionstr)
            connexcl.Open()

            Dim dtADPexcel As OleDbDataAdapter = New OleDbDataAdapter("SELECT * FROM [Sheet1$] ", connexcl)
            dtADPexcel.Fill(dtTableExcel)

            dtADPexcel.Dispose()

            If connexcl IsNot Nothing Then
                connexcl.Close()
                connexcl.Dispose()
            End If

        End If


        'get data from table Esch_Na_tbl_batch_no_group_and_batch_rules
        Dim conn As OleDbConnection = New OleDbConnection(ConfigurationManager.ConnectionStrings(dbConnectionName).ProviderName & ConfigurationManager.ConnectionStrings(dbConnectionName).ConnectionString)
        Dim dtAdapterFrom As OleDbDataAdapter = New OleDbDataAdapter("SELECT txt_order_key,int_line_no,txt_item_no,txt_lot_no,planned_production_qty,dat_start_date FROM Esch_Na_tbl_orders WHERE (dat_start_date > " & dateSeparator & Today.AddMonths(-1) & dateSeparator & ") And (dat_start_date < " & dateSeparator & Today.AddMonths(1) & dateSeparator & ")", conn)
        Dim dtTableFrom As DataTable = New DataTable
        dtAdapterFrom.Fill(dtTableFrom)

        Dim dtAdapterTo1 As OleDbDataAdapter = New OleDbDataAdapter("SELECT * FROM Esch_Na_tbl_differencesInBetweens ", conn)
        Dim cmdbAccessCmdBuilder1 As New OleDbCommandBuilder(dtAdapterTo1)
        dtAdapterTo1.DeleteCommand = cmdbAccessCmdBuilder1.GetDeleteCommand()
        dtAdapterTo1.InsertCommand = cmdbAccessCmdBuilder1.GetInsertCommand()
        Dim dtTableTo1 As DataTable = New DataTable
        dtAdapterTo1.Fill(dtTableTo1)
        Dim keys(0) As DataColumn
        keys(0) = dtTableTo1.Columns("txt_order_key")
        dtTableTo1.PrimaryKey = keys


        For Each a As DataRow In dtTableTo1.Rows
            a.Delete()
        Next


        'iterate 
        counter = 0
        For Each b As DataRow In dtTableExcel.Rows
            Dim newOrder() As DataRow = dtTableFrom.Select("txt_order_key = '" & b.Item("txt_order_key") & "'")
            If newOrder.Count > 0 Then
                If newOrder(0).Item("planned_production_qty") <> b.Item("planned_production_qty") Then
                    Dim newRow As DataRow = dtTableTo1.NewRow

                    newRow.Item("txt_order_key") = newOrder(0).Item("txt_order_key")
                    newRow.Item("txt_line_no") = newOrder(0).Item("int_line_no")
                    newRow.Item("txt_item_no") = newOrder(0).Item("txt_item_no")
                    newRow.Item("txt_lot_no") = newOrder(0).Item("txt_lot_no")
                    newRow.Item("planned_production_qty") = newOrder(0).Item("planned_production_qty")
                    newRow.Item("last_time_qty") = b.Item("planned_production_qty") 'planned_production_qty in last time
                    newRow.Item("dat_start_date") = newOrder(0).Item("dat_start_date")

                    dtTableTo1.Rows.Add(newRow)

                    counter += 1

                End If

            End If
        Next



        dtAdapterTo1.Update(dtTableTo1)


        dtTableExcel.Dispose()

        cmdbAccessCmdBuilder1.Dispose()

        dtTableTo1.Dispose()

        dtTableFrom.Dispose()

        dtAdapterTo1.Dispose()

        dtAdapterFrom.Dispose()

        'conn.Close()
        conn.Dispose()
    End Sub



    ''' <summary>
    ''' get data from dataTable and a excel is created to store the data, and generated excel is expected for downloading
    ''' </summary>
    ''' <param name="StatusLabel">to show the warning message or others</param>
    ''' <returns>return a excel file name</returns>
    Public Overloads Function generateExcel(ByRef StatusLabel As Label, ByRef dtTable As DataTable, ByVal filePath As String, ByVal startTime1 As DateTime, ByVal endTime1 As DateTime, ByVal pixelsPerDay As Integer, ByRef eScheduleFile As String, ByVal useTextMethodToPostSchedule As Boolean) As String

        If filePath.EndsWith("\") Then
            filePath = filePath.Remove(filePath.Length - 1, 1)
        End If

        Dim excelfilename As String = excelDatabaseFile
        Dim errorMessage As StringBuilder = New StringBuilder()


        Dim xlApp As New Object
        Dim xlBook As New Object
        Dim startTimeBeforexlApp As DateTime
        Dim startTimeAfterxlApp As DateTime

        Try
            startTimeBeforexlApp = DateTime.Now()
            xlApp = Server.CreateObject("Excel.Application")
            startTimeAfterxlApp = DateTime.Now()


            xlBook = xlApp.Workbooks.Add()

        Catch ex1 As Exception
            errorMessage.Append("R1" & ex1.Message)
            msgPopUP("Failed to open excel application." & errorMessage.ToString & "<br />", StatusLabel)
            Return String.Empty
        End Try

        Dim xclFileExtsn As String
        xclFileExtsn = IIf(CInt(xlApp.version) >= 12, ".xlsx", ".xls") 'see which version of excel you installed in the PC along with IIS (excel 2007 above)

        xlApp.visible = False 'make excel application open not visible
        xlApp.DisplayAlerts = False

        Try

            Dim rCount As Integer = dtTable.Rows.Count
            Dim cCount As Integer = dtTable.Columns.Count

            Dim dataArray(rCount + 1, cCount) As Object

            With xlBook.Worksheets(1)
                Dim cl As String
                For j As Integer = 0 To cCount - 1
                    dataArray(0, j) = dtTable.Columns(j).ColumnName
                    cl = dtTable.Columns(j).DataType.ToString
                    If cl.IndexOf("DateTime") > -1 Then
                        .columns(j + 1).NumberFormat = "m/d/yyyy"
                    Else
                        If Not (cl.IndexOf("Int") > -1) Then
                            .columns(j + 1).NumberFormat = "@"
                        End If
                    End If

                Next

            End With






            Dim rw As DataRow
            For i As Integer = 1 To rCount
                rw = dtTable.Rows(i - 1)
                For j As Integer = 0 To cCount - 1
                    dataArray(i, j) = rw.Item(j)
                Next
            Next

            xlBook.Worksheets(1).Range("A1").Resize(rCount + 1, cCount).Value = dataArray

        Catch ex2 As Exception
            'Page.ClientScript.RegisterClientScriptBlock(Me.GetType(), "script1", "<script language='javascript'>alert('Failed to create records.');</script>", False)
            errorMessage.Append("R2" & ex2.Message)
            msgPopUP("Failed to create records." & errorMessage.ToString() & "<br />", StatusLabel)
        Finally


        End Try




        Dim FileName As System.IO.FileInfo, FileNameExist As Boolean = True
        Try
            FileName = New System.IO.FileInfo(filePath & "\" & excelfilename & xclFileExtsn)
        Catch ex3 As Exception
            errorMessage.Append("R3" & ex3.Message)
            FileNameExist = False
        End Try

        If FileNameExist Then
            Try
                File.Delete(filePath & "\" & excelfilename & xclFileExtsn)
            Catch ex4 As Exception
                errorMessage.Append("R4" & ex4.Message)
                'Page.ClientScript.RegisterClientScriptBlock(Me.GetType(), "script1", "<script language='javascript'>alert('The file is in use by other program.');</script>", False)
                msgPopUP("The file is in use by other program." & errorMessage.ToString(), StatusLabel)
                Return "The file is in use by other program"
            End Try
        End If

        Try
            xlBook.SaveAs(filePath & "\" & excelfilename & xclFileExtsn)
        Catch ex5 As Exception
            errorMessage.Append("R5" & ex5.Message)
            xlBook.SaveCopyAs(filePath & "\" & excelfilename & xclFileExtsn) 'on server 2008, method of SaveAs is not supported
        End Try

        'Dim baseAdd As Integer = xlApp.Hinstance ' used as a reference to kill this process

        Try

            xlBook.close()

            xlApp.Quit()
        Catch ex6 As Exception
            errorMessage.Append("R6" & ex6.Message)
            msgPopUP("There is problem with Excel closure." & errorMessage.ToString, StatusLabel)

            Try
                Marshal.FinalReleaseComObject(xlBook)
                xlBook = Nothing
                Marshal.FinalReleaseComObject(xlApp)
                xlApp = Nothing

                GC.Collect()
                GC.WaitForPendingFinalizers()

            Catch ex As Exception  'below process is a little rude,immediately kill the process
                For Each p As Process In Process.GetProcessesByName("Excel")
                    If p.StartTime >= startTimeBeforexlApp AndAlso p.StartTime <= startTimeAfterxlApp Then
                        p.Kill()
                        Exit For
                    End If

                Next
            End Try

        End Try

        If useTextMethodToPostSchedule Then
            errorMessage.Append(generateTxtFile(StatusLabel, dtTable, filePath, startTime1, endTime1, pixelsPerDay, eScheduleFile))
            
            '=====copy file to another folder
            Try
                File.Copy(filePath & "\" & eScheduleFile, ConfigurationManager.AppSettings("productionScheduleFolder") & "\" & String.Format(ConfigurationManager.AppSettings("productionFileName"), Today), True)

            Catch ex As Exception
                errorMessage.Append("Error happened when copy mps text file to another folder: " & ex.Message)
            End Try

        Else
            errorMessage.Append(postScheduleToUrl(Nothing, dtTable, startTime1, endTime1, pixelsPerDay))
        End If



        Return excelfilename & xclFileExtsn

    End Function

    Public Function postScheduleToUrl(ByRef StatusLabel As Label, ByRef dtTable As DataTable, ByVal startTime1 As DateTime, ByVal endTime1 As DateTime, ByVal pixelsPerDay As Integer) As String
        Dim client As New WebClient


        Dim originalCulture As CultureInfo = Thread.CurrentThread.CurrentCulture
        Thread.CurrentThread.CurrentCulture = New CultureInfo("en-US")
        Dim outLinesText As New StringBuilder()
        Dim pullScrew As String = String.Empty
        Dim marginTop As String = String.Empty
        Dim SPANandETD As String = String.Empty, leftPosForRSDspan As String = String.Empty

        Dim linesList = From c In dtTable.AsEnumerable() Order By c.Item("int_line_no") Select c.Item("int_line_no") Distinct
        Dim linesString As String = String.Join(",", linesList)

        For Each a As DataRow In dtTable.Rows

            If Not DBNull.Value.Equals(a.Item("txt_pull_screw")) Then
                'additionGap += 3 * 6
                If String.Equals(a.Item("txt_pull_screw"), "yellow", StringComparison.OrdinalIgnoreCase) Then
                    'pullScrew = "<span style='color:#FFD700;font-size:150%;'>►</span>"
                    pullScrew = "#FFD700"
                Else
                    pullScrew = "red"
                End If
                marginTop = "margin-top:-0.22em;"
            Else
                pullScrew = String.Empty
                marginTop = String.Empty
            End If

            SPANandETD = a.Item("int_span") & " (" & a.Item("dat_etd") & ")"
            leftPosForRSDspan = "left:-" & (SPANandETD.Length + 0) * 8 & "px;"

            outLinesText.Append(a.Item("int_line_no") & "@" & pixelsPerDay & "@" & startTime1 & "@" & endTime1 & "@" & _
                                DateDiff("n", startTime1, a.Item("dat_start_date")) & "@" & _
                                DateDiff("n", a.Item("dat_start_date"), a.Item("dat_finish_date")) & "@" & _
                                a.Item("txt_lot_no") & "@" & _
                                a.Item("txt_item_no") & "@" & _
                                a.Item("planned_production_qty") & "KG@(" & _
                                a.Item("txt_order_key") & ")@" & _
                                a.Item("txt_VIP") & "@" & _
                                a.Item("txt_remark") & "@" & _
                                leftPosForRSDspan & "@" & _
                                SPANandETD & "@" & _
                                marginTop & "@" & _
                                pullScrew & "@" & _
                                linesString)


            outLinesText.Append("^")

        Next
        outLinesText.Remove(outLinesText.Length - 1, 1) 'delete the last vbCrLf
        'date time format based on original culture
        Thread.CurrentThread.CurrentCulture = originalCulture
        outLinesText.Clear()
        outLinesText.Append("test test!!!!!!!!!!!")
        Dim myNameValueCollection As New NameValueCollection()
        myNameValueCollection.Add("postSchedule", outLinesText.ToString)


        client.UploadValues("http://3.242.96.41/sh_intranet/E-ProdMeeting/Eminutes/esch_mps1.asp", "POST", myNameValueCollection)

        Return String.Empty

    End Function


    ''' <summary>
    ''' Generate one text file to display schedule
    ''' </summary>
    ''' <param name="StatusLabel"></param>
    ''' <param name="dtTable"></param>
    ''' <param name="filePath"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function generateTxtFile(ByRef StatusLabel As Label, ByRef dtTable As DataTable, ByVal filePath As String, ByVal startTime1 As DateTime, ByVal endTime1 As DateTime, ByVal pixelsPerDay As Integer, ByRef eScheduleFile As String) As String

        Dim msgReturn As New StringBuilder()

        eScheduleFile = String.Format("mps{0:yyyyMMdd}", Today) & ".txt"
        eScheduleFile = "mps.txt"
        Dim fileLocation As String = valueOf("strWebScheduleLocation")

        If String.IsNullOrEmpty(fileLocation) Then
            fileLocation = filePath
            If fileLocation.EndsWith("\") Then
                fileLocation = fileLocation.Remove(fileLocation.Length - 1, 1)
            End If
        End If

        Dim originalCulture As CultureInfo = Thread.CurrentThread.CurrentCulture
        Thread.CurrentThread.CurrentCulture = New CultureInfo("en-US")
        Dim outLinesText As New StringBuilder()
        Dim pullScrew As String = String.Empty
        Dim marginTop As String = String.Empty
        Dim SPANandETD As String = String.Empty, leftPosForRSDspan As String = String.Empty

        Dim linesList = From c In dtTable.AsEnumerable() Order By c.Item("int_line_no") Select c.Item("int_line_no") Distinct
        Dim linesString As String = String.Join(",", linesList)

        For Each a As DataRow In dtTable.Rows

            If Not DBNull.Value.Equals(a.Item("txt_pull_screw")) Then
                'additionGap += 3 * 6
                If String.Equals(a.Item("txt_pull_screw"), "yellow", StringComparison.OrdinalIgnoreCase) Then
                    'pullScrew = "<span style='color:#FFD700;font-size:150%;'>►</span>"
                    pullScrew = "#FFD700"
                Else
                    pullScrew = "red"
                End If
                marginTop = "margin-top:-0.22em;"
            Else
                pullScrew = String.Empty
                marginTop = String.Empty
            End If

            SPANandETD = a.Item("int_span") & " (" & a.Item("dat_etd") & ")"
            leftPosForRSDspan = "left:-" & (SPANandETD.Length + 0) * 8 & "px;"

            outLinesText.Append(a.Item("int_line_no") & "@" & pixelsPerDay & "@" & startTime1 & "@" & endTime1 & "@" & _
                                DateDiff("n", startTime1, a.Item("dat_start_date")) & "@" & _
                                DateDiff("n", a.Item("dat_start_date"), a.Item("dat_finish_date")) & "@" & _
                                a.Item("txt_lot_no") & "@" & _
                                a.Item("txt_item_no") & "@" & _
                                a.Item("planned_production_qty") & "KG@(" & _
                                a.Item("txt_order_key") & ")@" & _
                                a.Item("txt_VIP") & "@" & _
                                a.Item("txt_remark") & "@" & _
                                leftPosForRSDspan & "@" & _
                                SPANandETD & "@" & _
                                marginTop & "@" & _
                                pullScrew & "@(screw:" & _
                                a.Item("txt_process_technics") & ") @" & _
                                linesString)


            outLinesText.Append("^")

        Next

        If outLinesText.Length > 0 Then
            outLinesText.Remove(outLinesText.Length - 1, 1) 'delete the last vbCrLf
        Else
            outLinesText.Append("No data!!!!!!!! You might choose the wrong period.")
        End If

        'date time format based on original culture
        Thread.CurrentThread.CurrentCulture = originalCulture

        Try
            Using outFile As New StreamWriter(fileLocation & "/" & eScheduleFile)
                outFile.Write(outLinesText.ToString)
            End Using
        Catch ex As Exception
            msgReturn.AppendLine("<div style='color:red;'> " & ex.Message & "</div>")
            msgReturn.AppendLine("<div style='color:red;'> " & "You do not have the access right to write in folder " & fileLocation & "/" & eScheduleFile & "?</div>")
        End Try

        Return msgReturn.ToString

    End Function


End Class
