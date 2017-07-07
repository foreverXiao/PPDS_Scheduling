Imports System.IO
Imports System.Text
Imports System.Data
Imports System.Data.SqlClient
Imports System.Globalization
Imports Microsoft.VisualBasic
Imports System.Linq
Imports System.Threading
Imports System.Runtime.InteropServices
Imports System.Diagnostics



Partial Class SCMrelated_ordersLoading
    Inherits FrequentPlanActions

    'Protected Sub Page_Init1(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Init
    '    accesslevel = 4 ' set a low security level for this page in order to allow people to change their password
    'End Sub


    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load



        'StatusLabel.Text = _dtTable.Rows.Count



        'GV1.DataMember = "Table"




        If Not Page.IsPostBack Then
            Dim listdt As DataTable = New DataTable()

            ' Define the columns of the table.
            listdt.Columns.Add(New DataColumn("hour", GetType(Integer)))
            listdt.Columns.Add(New DataColumn("hourValue", GetType(Integer)))

            Dim dr As DataRow
            For i As Integer = 0 To 23
                dr = listdt.NewRow()
                dr("hour") = i
                dr("hourValue") = i
                listdt.Rows.Add(dr)
            Next

            Dim dv As DataView = New DataView(listdt)

            ddlHour1.DataSource = dv
            ddlHour1.DataTextField = "hour"
            ddlHour1.DataValueField = "hourValue"
            ddlHour1.DataBind()

            ddlHour2.DataSource = dv
            ddlHour2.DataTextField = "hour"
            ddlHour2.DataValueField = "hourValue"
            ddlHour2.DataBind()

            Dim listdt1 As DataTable = New DataTable()

            ' Define the columns of the table.
            listdt1.Columns.Add(New DataColumn("minute", GetType(Integer)))
            listdt1.Columns.Add(New DataColumn("minuteValue", GetType(Integer)))

            Dim dr1 As DataRow
            For j As Integer = 0 To 59
                dr1 = listdt1.NewRow()
                dr1("minute") = j
                dr1("minuteValue") = j
                listdt1.Rows.Add(dr1)
            Next

            Dim dv1 As DataView = New DataView(listdt1)

            ddlMinute1.DataSource = dv1
            ddlMinute1.DataTextField = "minute"
            ddlMinute1.DataValueField = "minuteValue"
            ddlMinute1.DataBind()

            ddlMinute2.DataSource = dv1
            ddlMinute2.DataTextField = "minute"
            ddlMinute2.DataValueField = "minuteValue"
            ddlMinute2.DataBind()

            'date time format based on  culture  of en-US
            Dim originalCulture As CultureInfo = Thread.CurrentThread.CurrentCulture
            Thread.CurrentThread.CurrentCulture = New CultureInfo("en-US")

            Dim startDate As DateTime = Today.AddMonths(-3)
            If startDate.DayOfWeek = DayOfWeek.Sunday Then
                startDate = startDate.AddDays(-6)
            Else
                startDate = startDate.AddDays(1 - startDate.DayOfWeek)
            End If
            earlierTime.Text = startDate.ToShortDateString
            laterTime.Text = startDate.AddYears(1).ToShortDateString

            'date time format based on original culture
            Thread.CurrentThread.CurrentCulture = originalCulture

            ddlHour1.SelectedIndex = 7
            ddlHour2.SelectedIndex = ddlHour1.SelectedIndex
            ddlMinute1.SelectedIndex = 30
            ddlMinute2.SelectedIndex = ddlMinute1.SelectedIndex



        End If

    End Sub

    Protected Sub dataTableToExcel(ByRef dt As DataTable)
        Dim response1 As HttpResponse = HttpContext.Current.Response()

        Dim filename As String = "test.xls"
        ' first let's clean up the response1.object
        response1.Clear()
        response1.Charset = ""

        ' set the response1 mime type for excel
        response1.ContentType = "application/vnd.ms-excel"
        response1.AddHeader("Content-Disposition", "attachment;filename=" & filename)

        ' create a string writer
        Dim sw As StringWriter = New StringWriter()

        Using sw
            Dim htw As HtmlTextWriter = New HtmlTextWriter(sw)

            Using htw
                'instantiate a datagrid
                Dim dg As DataGrid = New DataGrid()
                dg.DataSource = dt
                dg.DataBind()
                dg.RenderControl(htw)
                response1.Write(sw.ToString())
                response1.End()
            End Using
        End Using

    End Sub




    Protected Sub dataTableToExcel1(ByRef dt As DataTable)
        Dim attachment As String = "attachment;filename=city.xls"
        Response.ClearContent()
        Response.AddHeader("content-disposition", attachment)
        Response.ContentType = "application/vnd.ms-excel"

        Dim tab As String = ""
        For Each dc As DataColumn In dt.Columns
            Response.Write(tab + dc.ColumnName)
            tab = "\t"
        Next


        Response.Write("\n")


        For Each dr As DataRow In dt.Rows
            tab = ""

            For i As Integer = 0 To dt.Columns.Count - 1
                Response.Write(tab & dr(i).ToString())
                tab = "\t"
            Next
            Response.Write("\n")

        Next
        Response.End()
        'Response.Flush()
        'Response.End()


    End Sub


    Protected Sub Download1_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles Download1.Click

        Dim msgReturn As New StringBuilder()

        'File Path and File Name

        Dim FileNameExist As Boolean = True

        Dim filePath As String = ConfigurationManager.AppSettings("excelFolder")
        If filePath.EndsWith("\") Then filePath = filePath.Remove(filePath.Length - 1, 1)

        If Not Directory.Exists(filePath) Then
            Try
                Directory.CreateDirectory(filePath)
            Catch ex As Exception
                msgReturn.AppendLine("<div style='color:red;'>" & "You might not have right to create folder:M2:" & ex.Message & "</div>")
                'msgPopUP("You might not have right to create folder: ." & filePath & "<br />" & errorMessage.ToString, StatusLabel)
                FileNameExist = False
            End Try
        End If


        Dim FileName As System.IO.FileInfo
        Dim myFile As FileStream
        Dim _BinaryReader As BinaryReader


        Dim downloadFileName As String = CType(CacheFrom("loading"), String)

        Dim fileExtsn As String = ".xlsx"

        If FileNameExist AndAlso (Not String.IsNullOrEmpty(downloadFileName)) Then

            If downloadFileName.IndexOf(fileExtsn) < 0 Then
                fileExtsn = ".xls"
            End If


            Try
                FileName = New System.IO.FileInfo(filePath & "\" & downloadFileName)
                myFile = New FileStream(filePath & "\" & downloadFileName, FileMode.Open, FileAccess.Read, FileShare.Read)
                'Reads file as binary values
                _BinaryReader = New BinaryReader(myFile)
            Catch ex As Exception
                'errorMessage.Append("" & ex3.Message)
                msgReturn.AppendLine("<div style='color:red;'>" & "M3: The file does not exist. You need generate the file first." & "</div>")
                FileNameExist = False
            End Try


        Else
            FileNameExist = False
        End If



        If FileNameExist Then
            Try

                'StatusLabel.Text = "Download status: File is available now!<br />" & errorMessage.ToString

                Dim startBytes As Long = 0
                Dim lastUpdateTimeStamp As String = File.GetLastWriteTimeUtc(filePath).ToString("r")
                Dim _EncodedData = HttpUtility.UrlEncode(downloadFileName, Encoding.UTF8) & lastUpdateTimeStamp

                Response.Clear()
                Response.Buffer = False

                Response.AddHeader("Accept-Ranges", "bytes")
                Response.AppendHeader("ETag", "'" & _EncodedData & "'")
                Response.AppendHeader("Last-Modified", lastUpdateTimeStamp)
                Response.ContentType = "application/octet-stream"
                Response.AddHeader("Content-Disposition", "attachment;filename=" & String.Format(pageTitleReference() & "-{0:yyyy-MM-dd-HH-mm-ss}" & fileExtsn, System.DateTime.Now))
                Response.AddHeader("Content-Length", (FileName.Length - startBytes).ToString())
                Response.AddHeader("Connection", "Keep-Alive")
                Response.ContentEncoding = Encoding.UTF8

                'Send data
                _BinaryReader.BaseStream.Seek(startBytes, SeekOrigin.Begin)

                'Dividing the data in 10240 bytes package
                Dim maxCount As Integer = CInt(Math.Ceiling((FileName.Length - startBytes + 0.0) / 10240))

                'Download in block of 10k bytes
                Dim i As Integer = 0
                Do While ((i < maxCount) AndAlso Response.IsClientConnected)
                    Response.BinaryWrite(_BinaryReader.ReadBytes(10240))
                    Response.Flush()
                    i += 1
                Loop

                'If i < maxCount Then
                '    'Return
                'Else
                '    'Return
                'End If

            Catch ex As Exception
                msgReturn.AppendLine("<div style='color:red;'>" & "Download status: can not download the file. The following error occured::M4:" & ex.Message)
                'msgPopUP("Download status: can not download the file. The following error occured: " & errorMessage.ToString, StatusLabel)
            Finally
                _BinaryReader.Close()
                myFile.Close()
            End Try


        End If

        StatusLabel.Text = msgReturn.ToString


    End Sub





    ''' <summary>
    ''' Generate the FG list to be converted to RM usage
    ''' </summary>
    Protected Sub prdctnOrdrs_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles prdctnOrdrs.Click


        If Not IsDate(earlierTime.Text) Then
            earlierTime.ForeColor = Drawing.Color.Red
            msgPopUP("Illegal date time!", Message)
            Return
        End If


        If Not IsDate(laterTime.Text) Then
            laterTime.ForeColor = Drawing.Color.Red
            msgPopUP("Illegal date time!", Message)
            Return
        End If


        Dim msgRtrn As New StringBuilder

        Dim originalCulture As CultureInfo = Thread.CurrentThread.CurrentCulture
        Thread.CurrentThread.CurrentCulture = New CultureInfo("en-US")


        Dim start As DateTime = CDate(earlierTime.Text).AddHours(ddlHour1.SelectedIndex).AddMinutes(ddlMinute1.SelectedIndex)
        Dim finish As DateTime = CDate(laterTime.Text).AddHours(ddlHour2.SelectedIndex).AddMinutes(ddlMinute2.SelectedIndex)

        If start.CompareTo(finish) > 0 Then
            finish = start
            start = CDate(laterTime.Text).AddHours(ddlHour2.SelectedIndex).AddMinutes(ddlMinute2.SelectedIndex)
        End If


        Dim connstr As String = ConfigurationManager.ConnectionStrings(dbConnectionName).ConnectionString
        Dim conn As SqlConnection = New SqlConnection(connstr)
        Dim dtAdapter1 As SqlDataAdapter = New SqlDataAdapter("SELECT planned_production_qty,dat_start_date,dat_finish_date,int_line_no,flt_working_hours,int_change_over_time,txt_VIP,txt_gl_class From Esch_CQ_tbl_orders " &
                                                                   " WHERE ((dat_start_date between " & dateSeparator & start & dateSeparator & " And " & dateSeparator & finish & dateSeparator & ") Or " &
                                                                   "(dat_finish_date between " & dateSeparator & start & dateSeparator & " And " & dateSeparator & finish & dateSeparator & ") Or (dat_start_date < " & dateSeparator & start & dateSeparator & " And dat_finish_date > " & dateSeparator & finish & dateSeparator & " )) And ( planned_production_qty > 0 ) And (flt_working_hours > 0) And (CAST(int_line_no as VARCHAR(5)) <> '" & valueOf("intDummyLine") & "')" &
                                                                    " ORDER BY int_line_no,dat_start_date", conn)



        Dim _dtTable As DataTable = New DataTable()

        dtAdapter1.Fill(_dtTable)


        'Dim connParam As SqlConnection = New SqlConnection(ConfigurationManager.ConnectionStrings(dbConnForParam).ConnectionString)
        'Dim hasException As Boolean = False
        ''finishTime_exPlantDate_Span1(connParam, _dtTable.Select(Nothing), hasException)

        'connParam.Dispose()

        'remove some columns which are not necessary
        '_dtTable.Columns.Remove("int_span")
        '_dtTable.Columns.Remove("txt_end_user")
        '_dtTable.Columns.Remove("flt_working_hours")
        '_dtTable.Columns.Remove("int_change_over_time")
        '_dtTable.Columns.Remove("txt_grade")
        '_dtTable.Columns.Remove("dat_new_explant")
        '_dtTable.Columns.Remove("dat_etd")

        For Each r As DataRow In _dtTable.Select("(txt_VIP is Null) Or (txt_gl_class Is Null) ")
            r.Item("txt_VIP") = "non-VIP"
            If DBNull.Value.Equals(r.Item("txt_gl_class")) Then
                r.Item("txt_gl_class") = "Missed_Category"
            End If
        Next




        dtAdapter1.Dispose()
        conn.Dispose()




        Dim outputTable As DataTable = New DataTable()
        outputTable.Columns.Add(New DataColumn("year", GetType(Integer)))
        outputTable.Columns.Add(New DataColumn("month", GetType(Integer)))
        outputTable.Columns.Add(New DataColumn("week", GetType(Integer)))
        outputTable.Columns.Add(New DataColumn("date", GetType(DateTime)))
        outputTable.Columns.Add(New DataColumn("line_no", GetType(String)))
        outputTable.Columns.Add(New DataColumn("VIP", GetType(String)))
        outputTable.Columns.Add(New DataColumn("gl_class", GetType(String)))
        outputTable.Columns.Add(New DataColumn("dbl_loading", GetType(Integer)))

        Dim newRow As DataRow

        Dim currentStart As DateTime = start
        'If start.DayOfWeek = DayOfWeek.Sunday Then
        '    currentStart = start.AddDays(-6)
        'Else
        '    currentStart = start.AddDays(1 - start.DayOfWeek)
        'End If

        Dim currentFinish As DateTime = currentStart.AddDays(7)
        If currentFinish.DayOfWeek = DayOfWeek.Sunday Then
            currentFinish = currentFinish.AddDays(-6)
        Else
            currentFinish = currentFinish.AddDays(1 - currentStart.DayOfWeek)
        End If

        If currentFinish.CompareTo(finish) > 0 Then
            currentFinish = finish
        End If



        While currentStart.CompareTo(finish) < 0

            'Dim thisWeeksOrders = _dtTable.Select("(dat_start_date >=" & currentStart & " And dat_start_date <=" & currentFinish & ") Or (dat_finish_date  >=" & currentStart & " And  dat_finish_date <= " & currentFinish & ")").AsEnumerable

            'Dim newTable = From e In thisWeeksOrders Group e By key = New With {e.int_line_no, e.txt_VIP, e.txt_gl_class} Into group1 _
            '                         Select New abc with {.key1 =   e.int_line_no,.key2 = e.txt_VIP,.key3 = e.txt_gl_class,.qty = DateDiff('n',e.dat_start_date,e.dat_finish_date)/}


            Dim newRows() As DataRow = _dtTable.Select("(dat_start_date >= " & dateSeparator & currentStart & dateSeparator & " And dat_start_date <= " & dateSeparator & currentFinish & dateSeparator & ") Or (dat_finish_date  >= " & dateSeparator & currentStart & dateSeparator & " And  dat_finish_date <= " & dateSeparator & currentFinish & dateSeparator & ") Or (dat_finish_date > " & dateSeparator & currentFinish & dateSeparator & " And dat_start_date < " & dateSeparator & currentStart & dateSeparator & ")", "int_line_no,txt_VIP,txt_gl_class")

            Dim curLine As Integer, curVip As String, curGl_class As String, accumulatedQty As Double
            Dim first As DateTime, second As DateTime
            If newRows.Count > 0 Then
                curLine = newRows(0).Item("int_line_no")
                curVip = newRows(0).Item("txt_VIP")
                curGl_class = newRows(0).Item("txt_gl_class")
                first = IIf(CDate(newRows(0).Item("dat_start_date")).CompareTo(currentStart) > 0, CDate(newRows(0).Item("dat_start_date")), currentStart)
                second = IIf(CDate(newRows(0).Item("dat_finish_date")).CompareTo(currentFinish) < 0, CDate(newRows(0).Item("dat_finish_date")), currentFinish)
                accumulatedQty = DateDiff(DateInterval.Minute, first, second) / DateDiff(DateInterval.Minute, CDate(newRows(0).Item("dat_start_date")), CDate(newRows(0).Item("dat_finish_date"))) * newRows(0).Item("planned_production_qty")


                For i As Integer = 1 To newRows.Count - 1

                    first = IIf(CDate(newRows(i).Item("dat_start_date")).CompareTo(currentStart) > 0, CDate(newRows(i).Item("dat_start_date")), currentStart)
                    second = IIf(CDate(newRows(i).Item("dat_finish_date")).CompareTo(currentFinish) < 0, CDate(newRows(i).Item("dat_finish_date")), currentFinish)

                    If curLine = newRows(i).Item("int_line_no") AndAlso curVip = newRows(i).Item("txt_VIP") AndAlso curGl_class = newRows(i).Item("txt_gl_class") Then
                        accumulatedQty += DateDiff(DateInterval.Minute, first, second) / DateDiff(DateInterval.Minute, CDate(newRows(i).Item("dat_start_date")), CDate(newRows(i).Item("dat_finish_date"))) * newRows(i).Item("planned_production_qty")
                    Else
                        newRow = outputTable.NewRow()
                        newRow.Item("year") = currentStart.Year
                        newRow.Item("month") = currentStart.Month
                        newRow.Item("week") = Thread.CurrentThread.CurrentCulture.Calendar.GetWeekOfYear(currentStart, System.Globalization.CalendarWeekRule.FirstDay, DayOfWeek.Monday)
                        newRow.Item("date") = currentStart
                        newRow.Item("line_no") = curLine
                        newRow.Item("VIP") = curVip
                        newRow.Item("gl_class") = curGl_class
                        newRow.Item("dbl_loading") = CInt(accumulatedQty)
                        outputTable.Rows.Add(newRow)

                        curLine = newRows(i).Item("int_line_no")
                        curVip = newRows(i).Item("txt_VIP")
                        curGl_class = newRows(i).Item("txt_gl_class")
                        accumulatedQty = DateDiff(DateInterval.Minute, first, second) / DateDiff(DateInterval.Minute, CDate(newRows(i).Item("dat_start_date")), CDate(newRows(i).Item("dat_finish_date"))) * newRows(i).Item("planned_production_qty")

                    End If
                Next


                newRow = outputTable.NewRow()
                newRow.Item("year") = currentStart.Year
                newRow.Item("month") = currentStart.Month
                newRow.Item("week") = Thread.CurrentThread.CurrentCulture.Calendar.GetWeekOfYear(currentStart, System.Globalization.CalendarWeekRule.FirstDay, DayOfWeek.Monday)
                newRow.Item("line_no") = curLine
                newRow.Item("date") = currentStart
                newRow.Item("line_no") = curLine
                newRow.Item("VIP") = curVip
                newRow.Item("gl_class") = curGl_class
                newRow.Item("dbl_loading") = CInt(accumulatedQty)
                outputTable.Rows.Add(newRow)


            End If


            'If currentStart.DayOfWeek = DayOfWeek.Sunday Then
            '    currentStart = currentStart.AddDays(-6)
            'Else
            '    currentStart = currentStart.AddDays(1 - currentStart.DayOfWeek)
            'End If

            currentStart = currentFinish

            currentFinish = currentStart.AddDays(7)
            If currentFinish.CompareTo(finish) > 0 Then
                currentFinish = finish
            End If

        End While

        Thread.CurrentThread.CurrentCulture = originalCulture

        _dtTable.Dispose()


        'Dim first As DateTime, second As DateTime
        'For Each r As DataRow In _dtTable.Select(Nothing, Nothing, System.Data.DataViewRowState.ModifiedCurrent)
        '    first = IIf(CDate(r.Item("dat_start_date")).CompareTo(start) > 0, CDate(r.Item("dat_start_date")), start)
        '    second = IIf(CDate(r.Item("dat_finish_date")).CompareTo(finish) < 0, CDate(r.Item("dat_finish_date")), finish)

        '    r.Item("dat_start_date") = first

        '    If DateDiff(DateInterval.Minute, CDate(r.Item("dat_start_date")), CDate(r.Item("dat_finish_date"))) > 0 Then
        '        r.Item("planned_production_qty") = DateDiff(DateInterval.Minute, first, second) / DateDiff(DateInterval.Minute, CDate(r.Item("dat_start_date")), CDate(r.Item("dat_finish_date"))) * r.Item("planned_production_qty")
        '    Else
        '        r.Item("planned_production_qty") = 0
        '    End If


        '    If CDate(r.Item("dat_start_date")).CompareTo(start) < 0 AndAlso CDate(r.Item("dat_finish_date")).CompareTo(finish) > 0 Then
        '        _dtTable.Rows.Remove(r)
        '    End If



        'Next

        'total quantity
        Dim totalQty As Long = CLng(outputTable.Compute("sum(dbl_loading)", Nothing))


        Dim dtView As DataView = New DataView(outputTable)
        GV1.DataSource = dtView
        GV1.DataBind()

        msgRtrn.AppendLine("<div style='color:blue;'>The number of rows is " & dtView.Count & " and total quantity is " & String.Format("{0:##,###,###}", totalQty) & " kg" & "<br />(" & Now.ToString & ")</div>")

        'File Path and File Name
        Dim filePath As String = ConfigurationManager.AppSettings("excelFolder")



        If Not Directory.Exists(filePath) Then
            Try
                Directory.CreateDirectory(filePath)
            Catch ex As Exception
                'errorMessage.Append("M2" & ex.Message)
                'msgPopUP("You might not have right to create folder: ." & filePath & "<br />" & errorMessage.ToString, StatusLabel)
                msgRtrn.AppendLine("<div style='color:red;'>" & "You might not have right to create folder:M2:" & ex.Message & "</div>")
                Return
            End Try
        End If


        If filePath.EndsWith("\") Then filePath = filePath.Remove(filePath.Length - 1, 1)
        Dim excelReturnMsg As String = String.Empty
        Dim downloadFileName As String
        downloadFileName = generateExcel(excelReturnMsg, outputTable, filePath)
        outputTable.Dispose()
        msgRtrn.AppendLine(excelReturnMsg) 'get the error message from function generateExcel


        CacheInsert("loading", downloadFileName, 30)


        Message.Text = msgRtrn.ToString()

    End Sub





End Class

