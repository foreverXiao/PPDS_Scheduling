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



Partial Class SCMrelated_ordersForOF
    Inherits FrequentPlanActions

    Protected Sub Page_Init1(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Init
        accesslevel = 4 ' set a low security level for this page in order to allow people to change their password
    End Sub

    'Shared downloadFileName As String


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

            earlierTime.Text = DateTime.Today.AddDays(-30).ToShortDateString
            laterTime.Text = DateTime.Today.AddDays(180).ToShortDateString

            'date time format based on original culture
            Thread.CurrentThread.CurrentCulture = originalCulture

            ddlHour1.SelectedIndex = 7
            ddlHour2.SelectedIndex = ddlHour1.SelectedIndex
            ddlMinute1.SelectedIndex = 30
            ddlMinute2.SelectedIndex = ddlMinute1.SelectedIndex



        End If

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


        Dim downloadFileName As String = CType(CacheFrom("OrderF"), String)


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

        Dim start As DateTime = CDate(earlierTime.Text).AddHours(ddlHour1.SelectedIndex).AddMinutes(ddlMinute1.SelectedIndex)
        Dim finish As DateTime = CDate(laterTime.Text).AddHours(ddlHour2.SelectedIndex).AddMinutes(ddlMinute2.SelectedIndex)
        Dim connstr As String = ConfigurationManager.ConnectionStrings(dbConnectionName).ConnectionString
        Dim conn As SqlConnection = New SqlConnection(connstr)
        'Dim dtAdapter1 As SqlDataAdapter = New SqlDataAdapter("SELECT txt_item_no,planned_production_qty,dat_start_date,dat_finish_date,int_line_no,txt_lot_no,txt_order_key,txt_local_so,int_span,txt_end_user,flt_working_hours,int_change_over_time,txt_grade,dat_new_explant,dat_etd,dat_rdd,dat_order_added From Esch_CQ_tbl_orders " &
        Dim dtAdapter1 As SqlDataAdapter = New SqlDataAdapter("SELECT * From Esch_CQ_tbl_orders " &
                                                                   " WHERE ((dat_start_date between " & dateSeparator & start & dateSeparator & " And " & dateSeparator & finish & dateSeparator & ") Or " &
                                                                   "(dat_finish_date between " & dateSeparator & start & dateSeparator & " And " & dateSeparator & finish & dateSeparator & "))  Or (CAST(int_line_no as VARCHAR(5)) = '" & valueOf("intDummyLine") & "')" &
                                                                    " ORDER BY int_line_no,dat_start_date", conn)



        Dim _dtTable As DataTable = New DataTable()

        dtAdapter1.Fill(_dtTable)


        'Dim connParam As SqlConnection = New SqlConnection(ConfigurationManager.ConnectionStrings(dbConnForParam).ConnectionString)
        'Dim hasException As Boolean = False
        'finishTime_exPlantDate_Span1(connParam, _dtTable.Select(Nothing), hasException)

        'connParam.Dispose()

        'remove some columns which are not necessary
        Try
            _dtTable.Columns.Remove("txt_allocation_status")
            _dtTable.Columns.Remove("lng_VIP_lead_time")
            _dtTable.Columns.Remove("lng_AdvanceOfRevision")
            _dtTable.Columns.Remove("txt_auxiliary_code")
            _dtTable.Columns.Remove("txt_auxiliary_code_for_line_no")
            _dtTable.Columns.Remove("txt_line_assign")
            _dtTable.Columns.Remove("txt_FromUser")
            _dtTable.Columns.Remove("txt_ToUser")
            '_dtTable.Columns.Remove("int_key")
        Catch ex As Exception
            msgRtrn.AppendLine("<div style='color:red;'>" & ex.Message & "</div>")
        End Try


        dtAdapter1.Dispose()
        conn.Dispose()




        'total quantity
        Dim totalQty As Long = CLng(_dtTable.Compute("sum(planned_production_qty)", Nothing))


        Dim dtView As DataView = New DataView(_dtTable)
        GV1.DataSource = dtView
        GV1.DataBind()

        msgRtrn.AppendLine("<div style='color:blue;'>The number of rows is " & dtView.Count & " and total quantity is " & String.Format("{0:##,###}", totalQty) & " kg<br />(" & Now.ToString & ")</div>")

        'StatusLabel.Text = _dtTable.Rows.Count


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
        downloadFileName = generateExcel(excelReturnMsg, _dtTable, filePath)
        _dtTable.Dispose()
        msgRtrn.AppendLine(excelReturnMsg) 'get the error message from function generateExcel


        CacheInsert("OrderF", downloadFileName, 30)

        Message.Text = msgRtrn.ToString()





    End Sub





End Class

