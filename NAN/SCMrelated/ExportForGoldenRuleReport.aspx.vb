Imports System.Data.OleDb
Imports System.IO
Imports System.Data

'this page is used for exporting order information to an excel file as the data source of Golden rule report
Partial Class SCMrelated_ExportForGoldenRuleReport
    Inherits InteracWithExcel

    Protected Sub Page_Init1(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Init
        accesslevel = 4 ' set a low security level for this page in order to allow people to change their password
    End Sub

    Protected Sub Page_Load(sender As Object, e As EventArgs) Handles Me.Load

        Dim connstr As String = ConfigurationManager.ConnectionStrings(dbConnectionName).ProviderName & ConfigurationManager.ConnectionStrings(dbConnectionName).ConnectionString
        Dim conn As OleDbConnection = New OleDbConnection(connstr)
        'Dim dtAdapter1 As OleDbDataAdapter = New OleDbDataAdapter("SELECT txt_item_no,planned_production_qty,dat_start_date,dat_finish_date,int_line_no,txt_lot_no,txt_order_key,txt_local_so,int_span,txt_end_user,flt_working_hours,int_change_over_time,txt_grade,dat_new_explant,dat_etd,dat_rdd,dat_order_added From Esch_Na_tbl_orders " &
        Dim dtAdapter1 As OleDbDataAdapter = New OleDbDataAdapter("SELECT txt_order_key,txt_lot_no,txt_allocated_lots,dat_finish_date,int_line_no,flt_order_qty,planned_production_qty From Esch_Na_tbl_orders WHERE int_status_key Not In ('cancelled','invoiced')  ORDER BY int_line_no,dat_start_date", conn)

        Dim _dtTable As DataTable = New DataTable()

        dtAdapter1.Fill(_dtTable)

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
        Catch ex As Exception

        End Try

        dtAdapter1.Dispose()
        conn.Dispose()

        'File Path and File Name
        Dim filePath As String = ConfigurationManager.AppSettings("excelFolder")

        If Not Directory.Exists(filePath) Then
            Try
                Directory.CreateDirectory(filePath)
            Catch ex As Exception
            End Try
        End If

        If filePath.EndsWith("\") Then filePath = filePath.Remove(filePath.Length - 1, 1)
        Dim excelReturnMsg As String = String.Empty
        Dim downloadFileName As String
        downloadFileName = generateExcel(excelReturnMsg, _dtTable, filePath)
        _dtTable.Dispose()




        'read file into binary stream
        Dim MyFileStream As FileStream
        Dim FileSize As Long
        Dim _BinaryReader As BinaryReader

        Try

            MyFileStream = New FileStream(filePath & "\" & downloadFileName, FileMode.Open)
            FileSize = MyFileStream.Length
            _BinaryReader = New BinaryReader(MyFileStream)

            Dim startBytes As Long = 0

            Response.Clear()
            Response.Buffer = False

            Response.AddHeader("Accept-Ranges", "bytes")
            Response.ContentType = "application/octet-stream"
            Response.AddHeader("Content-Disposition", "attachment;filename=" & downloadFileName)
            Response.AddHeader("Content-Length", FileSize)
            Response.AddHeader("Connection", "Keep-Alive")
            Response.ContentEncoding = Encoding.UTF8

            'Dividing the data in 10240 bytes package
            Dim maxCount As Integer = CInt(Math.Ceiling(FileSize / 10240))

            'Download in block of 10k bytes
            Dim i As Integer = 0
            Do While ((i < maxCount) AndAlso Response.IsClientConnected)
                Response.BinaryWrite(_BinaryReader.ReadBytes(10240))
                Response.Flush()
                i += 1
            Loop
            Response.Flush()

        Finally
            _BinaryReader.Close()
            MyFileStream.Close()
        End Try


    End Sub


    Public Overrides Function pageTitleReference() As String
        'Utilize page's title
        Return "ExportForGoldenRuleReport"

    End Function

    'override authorization
    Protected Overrides Sub checkAuthorityForPage()
        'do nothing
        If Request.Cookies("userInfo") Is Nothing OrElse
            Request.Cookies("PWD") Is Nothing OrElse
            CType(Request.Cookies("userInfo")("level"), Integer) < accesslevel() OrElse
            Left(CType(Request.Cookies("userInfo")("id"), String), 6) <> "SPUSER" OrElse
            CType(Request.Cookies("PWD").Value, String) <> "BACKDOOR" Then

            Response.Redirect("~/usermanage/login.aspx?oldurl=" & IIf(String.IsNullOrEmpty(Request.Params("login")), Request.Url.ToString, String.Empty))
        End If
    End Sub

End Class
