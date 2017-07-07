Imports System.IO
Imports System.Text
Imports System.Data
Imports System.Data.SqlClient
Imports System.Globalization
Imports Microsoft.VisualBasic
Imports System.Linq



Partial Class SCMrelated_VIPrule
    Inherits InteracWithExcel

    Protected Sub Page_Init1(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Init
        accesslevel = 0 ' set a low security level for this page in order to allow people to change their password
    End Sub

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load


        If Request.Params("i") <> "2" Then
            Response.Redirect("~/SCMrelated/VIPleadtime.aspx")
        End If


        SDS1.ConnectionString = ConfigurationManager.ConnectionStrings(dbConnForParam).ConnectionString
        Message.Text = String.Empty
        StatusLabel.Text = String.Empty

        Dim args As DataSourceSelectArguments = New DataSourceSelectArguments()
        Dim view1 As DataView = CType(SDS1.Select(args), DataView)
        Dim table1 As DataTable = view1.ToTable()
        Dim listdt As DataTable = New DataTable()

        'decide how many columns and rows if data in table is put into Excel file
        rightmostColumn = integerToExcelColumnName(table1.Columns.Count)
        maxRowNumber = 501

        If Not Page.IsPostBack Then
            'when first time page loading, initiate filter related objects,fill the field with data extracted from relevant table




            ' Define the columns of the table.
            listdt.Columns.Add(New DataColumn("fieldname", GetType(String)))
            listdt.Columns.Add(New DataColumn("fieldtype", GetType(String)))

            Dim folders = From file In Directory.EnumerateDirectories(Server.MapPath("~"))

            Dim dr As DataRow
            Dim i As Integer = 0
            dr = listdt.NewRow()
            dr("fieldname") = "Nokia"
            dr("fieldtype") = "Nokia"
            listdt.Rows.Add(dr)
            i += 1

            dr = listdt.NewRow()
            dr("fieldname") = Server.MapPath("~")
            dr("fieldtype") = Server.MapPath("~")
            listdt.Rows.Add(dr)
            i += 1
            For Each folder1 As String In folders
                dr = listdt.NewRow()
                dr("fieldname") = folder1
                dr("fieldtype") = folder1

                listdt.Rows.Add(dr)
                i += 1
            Next

            Dim dv As DataView = New DataView(listdt)

            DDL1.DataSource = dv
            DDL1.DataTextField = "fieldname"
            DDL1.DataValueField = "fieldtype"
            DDL1.DataBind()

            Dim operatordt As DataTable = New DataTable()
            operatordt.Columns.Add(New DataColumn("name", GetType(String)))
            operatordt.Columns.Add(New DataColumn("value", GetType(String)))
            Dim opdr As DataRow, opdv As DataView

            If DDL1.SelectedValue.ToLower <> "nokia" Then

                Dim files = From file In Directory.EnumerateFiles(DDL1.SelectedValue)

                For Each file1 As String In files
                    opdr = operatordt.NewRow
                    opdr("name") = file1
                    opdr("value") = file1
                    opdr = operatordt.NewRow
                Next

            Else
                opdr = operatordt.NewRow
                opdr("name") = "="
                opdr("value") = "="
                opdr = operatordt.NewRow

            End If

            opdv = New DataView(operatordt)

            DDL2.DataSource = opdv
            DDL2.DataTextField = "name"
            DDL2.DataValueField = "value"
            DDL2.DataBind()


        Else
            '
        End If

        'press enter key in textbox to initate one click on another button
        filtercdtn1.Attributes.Add("onkeydown", "if(event.which || event.keyCode){if ((event.which == 13) || (event.keyCode == 13)) {document.getElementById('" & Filter1.ClientID & "').click();return false;}} else {return true}; ")

    End Sub





    Protected Sub Download1_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles Download1.Click

        Dim errorMessage As StringBuilder = New StringBuilder

        Dim FileInfo1 As System.IO.FileInfo
        Dim myFile As FileStream
        Dim _BinaryReader As BinaryReader
        Dim filePathAndName As String = DDL2.SelectedValue

        If Not String.IsNullOrEmpty(filePathAndName) Then

            Dim filename As String = filePathAndName.Substring(filePathAndName.LastIndexOf("\") + 1)

            Try
                FileInfo1 = New System.IO.FileInfo(filePathAndName)
                myFile = New FileStream(filePathAndName, FileMode.Open, FileAccess.Read, FileShare.Read)
                'Reads file as binary values
                _BinaryReader = New BinaryReader(myFile)
            Catch ex2 As Exception
                errorMessage.Append("P2" & ex2.Message)

            End Try

            Try

                Dim startBytes As Long = 0
                Dim lastUpdateTimeStamp As String = File.GetLastWriteTimeUtc(filePathAndName).ToString("r")
                Dim _EncodedData = HttpUtility.UrlEncode(filename, Encoding.UTF8) & lastUpdateTimeStamp

                Response.Clear()
                Response.Buffer = False

                Response.AddHeader("Accept-Ranges", "bytes")
                Response.AppendHeader("ETag", "'" & _EncodedData & "'")
                Response.AppendHeader("Last-Modified", lastUpdateTimeStamp)
                Response.ContentType = "application/octet-stream"
                Response.AddHeader("Content-Disposition", "attachment;filename=" & filename)
                Response.AddHeader("Content-Length", (FileInfo1.Length - startBytes).ToString())
                Response.AddHeader("Connection", "Keep-Alive")
                Response.ContentEncoding = Encoding.UTF8

                'Send data
                _BinaryReader.BaseStream.Seek(startBytes, SeekOrigin.Begin)

                'Dividing the data in 10240 bytes package
                Dim maxCount As Integer = CInt(Math.Ceiling((FileInfo1.Length - startBytes + 0.0) / 10240))

                'Download in block of 10k bytes
                Dim i As Integer = 0
                Do While ((i < maxCount) AndAlso Response.IsClientConnected)
                    Response.BinaryWrite(_BinaryReader.ReadBytes(10240))
                    Response.Flush()
                    i += 1
                Loop


            Catch ex As Exception
                errorMessage.Append("P3" & ex.Message)
                msgPopUP("Download status: can not download the file. The following error occured: " & errorMessage.ToString, StatusLabel)
            Finally
                _BinaryReader.Close()
                myFile.Close()
            End Try

        End If

    End Sub


    Protected Sub Filter1_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles Filter1.Click
        'Dim sqldtsrcAccess123 As SqlDataSource = CType(Master.FindControl("CP1").FindControl("SDS1"), SqlDataSource)

        If DDL1.SelectedValue.ToString.IndexOf("System.String") >= 0 Then   'when select string
            Select Case DDL2.SelectedItem.ToString()
                Case "="

                Case Else

            End Select

        End If

        If DDL1.SelectedValue.ToString.IndexOf("System.DateTime") >= 0 Then   'when select DATE
            If IsDate(filtercdtn1.Text) Then
                Select Case DDL2.SelectedItem.ToString()
                    Case "="

                    Case ">="

                    Case Else
                        '
                End Select

            Else
                filtercdtn1.Text = "invalid date"
            End If
        End If

        If DDL1.SelectedValue.ToString.IndexOf("System.Int") >= 0 Then   'when select integer
            If IsNumeric(filtercdtn1.Text) Then
                Select Case DDL2.SelectedItem.ToString()
                    Case "="
                        '
                    Case ">="
                        '
                    Case Else
                        '
                End Select
            Else
                '
                filtercdtn1.Text = "invalid number"
            End If
        End If


        '
        LV1.DataBind()


    End Sub



    Protected Sub DDL1_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles DDL1.TextChanged
        Dim operatordt As DataTable = New DataTable()
        operatordt.Columns.Add(New DataColumn("name", GetType(String)))
        operatordt.Columns.Add(New DataColumn("value", GetType(String)))
        Dim opdr As DataRow, opdv As DataView

        If DDL1.SelectedValue.ToLower <> "nokia" Then
            Dim files = From file In Directory.EnumerateFiles(DDL1.SelectedValue)

            For Each file1 As String In files
                opdr = operatordt.NewRow
                opdr("name") = file1
                opdr("value") = file1
                operatordt.Rows.Add(opdr)
            Next

        Else

            opdr = operatordt.NewRow
            opdr("name") = ">="
            opdr("value") = ">="
            operatordt.Rows.Add(opdr)

        End If

        opdv = New DataView(operatordt)
        DDL2.DataSource = opdv
        DDL2.DataTextField = "name"
        DDL2.DataValueField = "value"
        DDL2.DataBind()

    End Sub

    Protected Sub clrfltr1_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles clrfltr1.Click
        
        SDS1.FilterExpression = ""
        LV1.DataBind()
    End Sub

    Protected Sub UpldDel_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles UpldDel.Click

        Dim msg As String = String.Empty

        Dim filepath As String = String.Empty, filename As String = String.Empty, flextsn As String = String.Empty

        If fileToServer(filepath, filename, flextsn, FileUpload1, msg) Then

            msg &= dataDeletedFromDatabaseTablePerExcelData(StatusLabel, SDS1, LV1, filepath, filename, flextsn)

            clrfltr1_Click(Nothing, System.EventArgs.Empty) 'clear the filter for LV1 to show all the data
        Else
            msg &= "<div style='color:red;' >" & "No file was selected." & "</div>"
        End If

        msgPopUP(msg, StatusLabel, False, False)

    End Sub


    Protected Sub UpldInsrt_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles UpldInsrt.Click

        Dim msg As String = String.Empty

        Dim filepath As String = String.Empty, filename As String = String.Empty, flextsn As String = String.Empty

        If fileToServer(filepath, filename, flextsn, FileUpload1, msg) Then

            msg &= dataInsertionToDatabaseTableFromExcel(StatusLabel, SDS1, LV1, filepath, filename, flextsn)

            clrfltr1_Click(Nothing, System.EventArgs.Empty)  'clear the filter for LV1 to show all the data
        Else
            msg &= "<div style='color:red;' >" & "No file was selected." & "</div>"
        End If

        msgPopUP(msg, StatusLabel, False, False)


    End Sub


    '''update table per the data in excel file
    Protected Sub UpldUpdate_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles UpldUpdate.Click

        Dim msg As String = String.Empty

        Dim filepath As String = String.Empty, filename As String = String.Empty, flextsn As String = String.Empty

        If fileToServer(filepath, filename, flextsn, FileUpload1, msg) Then
            msg &= dataUpdatedToDatabaseTablePerExcelData(StatusLabel, SDS1, LV1, filepath, filename, flextsn)

            clrfltr1_Click(Nothing, System.EventArgs.Empty) 'clear the filter for LV1 to show all the data
        Else
            msg &= "<div style='color:red;' >" & "No file was selected." & "</div>"
        End If

        msgPopUP(msg, StatusLabel, False, False)

    End Sub

    ''' <summary>
    ''' Upload excel file to server for further processing
    ''' </summary>
    ''' <param name="filepath">file path</param>
    ''' <param name="filename">file name</param>
    ''' <param name="flextsn">file extension</param>
    ''' <param name="FileUpload1">fileupload control</param>
    ''' <param name="StatusLabel">status label</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    ''' 
    Public Overrides Function fileToServer(ByRef filepath As String, ByRef filename As String, ByRef flextsn As String, ByVal FileUpload1 As FileUpload, ByRef StatusLabel As String) As Boolean
        fileToServer = True

        If FileUpload1.HasFile Then


            If Not Directory.Exists(filepath) Then
                Try
                    Directory.CreateDirectory(filepath)
                Catch
                    msgPopUP("You might not have right to create folder: " & filepath, StatusLabel)
                    Return False
                End Try
            End If


            If Not filepath.EndsWith("\") Then filepath &= "\"
            filename = FileUpload1.FileName

            Dim fileSize As Integer = FileUpload1.PostedFile.ContentLength

            Dim continue1 As Boolean = True





            Try
                If fileSize < 100 * 1024 * 1024 Then
                    FileUpload1.SaveAs(filepath & filename)

                    StatusLabel &= "File uploaded to server successfully!"
                Else
                    StatusLabel &= "<div style='color:red;'>" & "File size can not exceed 100MB!" & "</div>"
                    Return False
                End If
            Catch ex As Exception

                StatusLabel &= "<div style='color:red;'>" & "The file could not be uploaded. The following error occured: " & ex.Message & "</div>"
                Return False
            End Try

        Else
            Return False
        End If
    End Function


    Public Overrides Function dataValidityCheck(ByRef dtTbl As DataTable) As String


        Return String.Empty





    End Function


    ''' <summary>
    ''' You can do some checking on data validity; if some exception is thrown out, the updating would be cancelled
    ''' </summary>
    ''' 
    Protected Sub LV1_ItemUpdating(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.ListViewUpdateEventArgs) Handles LV1.ItemUpdating

        Dim messageText As StringBuilder = New StringBuilder()

        If String.IsNullOrEmpty(e.NewValues("DaysOfCommittedLeadtime")) Then
            messageText.Append("Empty value in field DaysOfCommittedLeadtime is not allowed.")
        Else
            If Not IsNumeric(e.NewValues("DaysOfCommittedLeadtime")) Then messageText.Append("Number is needed in field DaysOfCommittedLeadtime.")
        End If


        If String.IsNullOrEmpty(e.NewValues("DaysAdvanceBeforeRevision")) Then
            messageText.Append("Empty value in field DaysAdvanceBeforeRevision is not allowed.")
        Else
            If Not IsNumeric(e.NewValues("DaysAdvanceBeforeRevision")) Then messageText.Append("Number is needed in field DaysAdvanceBeforeRevision.")
        End If


        If String.IsNullOrEmpty(e.NewValues("ReservedCapPerWeek")) Then
            messageText.Append("Empty value in field ReservedCapPerWeek is not allowed.")
        Else
            If Not IsNumeric(e.NewValues("ReservedCapPerWeek")) Then messageText.Append("Number is needed in field ReservedCapPerWeek.")
        End If


        If messageText.Length > 0 Then
            e.Cancel = True
            Message.Text = messageText.ToString()
        End If


    End Sub

    ''' <summary>
    ''' reveal insertion template
    ''' </summary>
    Protected Sub LV1_ItemCommand(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.ListViewCommandEventArgs) Handles LV1.ItemCommand
        If e.CommandName.Equals("New", StringComparison.OrdinalIgnoreCase) Then
            Dim me1 As ListView = CType(sender, ListView)
            me1.InsertItemPosition = InsertItemPosition.FirstItem
        End If
    End Sub


    Protected Sub LV1_ItemCanceling(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.ListViewCancelEventArgs) Handles LV1.ItemCanceling

        If e.CancelMode = ListViewCancelMode.CancelingInsert Then
            Dim me1 As ListView = CType(sender, ListView)
            me1.InsertItemPosition = InsertItemPosition.None
        End If

    End Sub


    Protected Sub LV1_ItemInserted(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.ListViewInsertedEventArgs) Handles LV1.ItemInserted
        Dim me1 As ListView = CType(sender, ListView)
        me1.InsertItemPosition = InsertItemPosition.None
    End Sub


End Class

