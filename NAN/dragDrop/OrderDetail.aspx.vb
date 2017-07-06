Imports System.IO
Imports System.Text
Imports System.Data
Imports System.Data.OleDb
Imports System.Globalization
Imports Microsoft.VisualBasic
Imports System.Linq



Partial Class dragDrop_OrderDetail
    Inherits FrequentPlanActions




    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load


        Dim connstr As String = ConfigurationManager.ConnectionStrings(dbConnectionName).ProviderName & ConfigurationManager.ConnectionStrings(dbConnectionName).ConnectionString

        SDS1.ConnectionString = connstr

        maxRowNumber = 15001

        'If Not IsNothing(Request.Cookies("userInfo")) AndAlso Not IsNothing(CacheFrom( "Oexception")) Then
        If Not IsNothing(CacheFrom("Oexception")) Then
            exceptionR.Visible = True
        End If


        If Not String.IsNullOrEmpty(Cache("lastUploadTime")) Then
            lblUpdateTime.Text = FormatDateTime(CDate(Cache("lastUploadTime")), DateFormat.GeneralDate)
        End If


        pageLoadInitiate(SDS1, DDL1, DDL2, filtercdtn1, Filter1, Download1, hiddenBT)


    End Sub


    Protected Sub Download1_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles Download1.Click


        downloadExcelFileFromSqlDataSource(SDS1, StatusLabel, hiddenBT)


    End Sub




    Protected Sub Filter1_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles Filter1.Click

        filterClickHnadler(filtercdtn1, DDL1, DDL2, SDS1, LV1)



    End Sub



    Protected Sub DDL1_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles DDL1.TextChanged
        DDLchangeSelection(DDL1, DDL2)
    End Sub

    Protected Sub clrfltr1_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles clrfltr1.Click
        clrfltrClickHnadler(SDS1, LV1)

    End Sub

    ''' <summary>
    ''' first to click the button update,second to click the button routineCheck
    ''' </summary>
    Protected Sub UpldUpdate_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles UpldUpdate.Click

        Dim msg As String = String.Empty
        Dim userName As String = lockKeyTable(priority.excelUpload)

        If String.IsNullOrEmpty(userName) Then
            Dim filepath As String = String.Empty, filename As String = String.Empty, flextsn As String = String.Empty
            If fileToServer(filepath, filename, flextsn, FileUpload1, msg) Then

                msg &= dataUpdatedToDatabaseTablePerExcelData(StatusLabel, SDS1, LV1, filepath, filename, flextsn)
                clrfltr1_Click(Nothing, System.EventArgs.Empty)  'clear the filter for LV1 to show all the data
            Else
                msg &= "<div style='color:red;'>No file was selected.</div>"
            End If

            unlockKeyTable(priority.excelUpload)
        Else
            msg &= "<div style='color:red;'>" & userName & " is using the order detail table" & "</div>"
        End If

        If Not CacheFrom("Oexception") Is Nothing Then
            exceptionR.Visible = True
        End If
        msgPopUP(msg, StatusLabel, False, False)

    End Sub



    Protected Sub UpldDel_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles UpldDel.Click

        Dim msg As String = String.Empty
        Dim userName As String = lockKeyTable(priority.excelUpload)
        If String.IsNullOrEmpty(userName) Then

            Dim filepath As String = String.Empty, filename As String = String.Empty, flextsn As String = String.Empty
            If fileToServer(filepath, filename, flextsn, FileUpload1, msg) Then
                msg &= dataDeletedFromDatabaseTablePerExcelData(StatusLabel, SDS1, LV1, filepath, filename, flextsn)
                clrfltr1_Click(Nothing, System.EventArgs.Empty) 'clear the filter for LV1 to show all the data
            Else
                msg &= "<div style='color:red;'>No file was selected.</div>"
            End If

            unlockKeyTable(priority.excelUpload)
        Else
            msg &= "<div style='color:red;font-size:150%;'>" & userName & " is using the order detail table" & "</div>"
        End If

        msgPopUP(msg, StatusLabel, False, False)


    End Sub


    ''' <summary>
    ''' first to click the button insert,second to click the button routineCheck
    ''' </summary>
    Protected Sub UpldInsrt_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles UpldInsrt.Click
        Dim msg As String = String.Empty

        Dim userName As String = lockKeyTable(priority.excelUpload)
        If String.IsNullOrEmpty(userName) Then

            Dim filepath As String = String.Empty, filename As String = String.Empty, flextsn As String = String.Empty
            If fileToServer(filepath, filename, flextsn, FileUpload1, msg) Then

                msg &= dataInsertionToDatabaseTableFromExcel(StatusLabel, SDS1, LV1, filepath, filename, flextsn)
                clrfltr1_Click(Nothing, System.EventArgs.Empty)  'clear the filter for LV1 to show all the data

            Else
                msg &= "<div style='color:red;'>No file was selected.</div>"
            End If

            unlockKeyTable(priority.excelUpload)
        Else
            msg &= "<div style='color:red;'>" & userName & " is using the order detail table" & "</div>"
        End If

        If Not CacheFrom("Oexception") Is Nothing Then
            exceptionR.Visible = True
        End If

        msgPopUP(msg, StatusLabel, False, False)

    End Sub

    Protected Sub overwrite_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles overwrite.Click
        Dim msg As String = String.Empty
        Dim userName As String = lockKeyTable(priority.excelUpload)

        If String.IsNullOrEmpty(userName) Then
            Dim filepath As String = String.Empty, filename As String = String.Empty, flextsn As String = String.Empty
            Dim continue1 As Boolean = False
            If fileToServer(filepath, filename, flextsn, FileUpload1, msg) Then
                msg &= dataDeletedFromDatabaseTablePerExcelDataForOverwrite(StatusLabel, SDS1, LV1, filepath, filename, flextsn, continue1) & "<br />"
                If continue1 Then
                    msg &= dataInsertionToDatabaseTableFromExcelForOverwrite(StatusLabel, SDS1, LV1, filepath, filename, flextsn)
                    Cache.Insert("lastUploadTime", Now(), Nothing, Cache.NoAbsoluteExpiration, TimeSpan.FromDays(2))
                End If
                clrfltr1_Click(Nothing, System.EventArgs.Empty)  'clear the filter for LV1 to show all the data
            Else
                msg &= "<div style='color:red;'>No file was selected.</div>"
            End If

            If Not CacheFrom("Oexception") Is Nothing Then
                exceptionR.Visible = True
            End If

            unlockKeyTable(priority.excelUpload)
        Else
            msg &= "<div style='color:red;'>" & userName & " is using the order detail table" & "</div>"
        End If

        If Not String.IsNullOrEmpty(Cache("lastUploadTime")) Then
            lblUpdateTime.Text = FormatDateTime(CDate(Cache("lastUploadTime")), DateFormat.GeneralDate)
        End If
        msgPopUP(msg, StatusLabel, False, False)

    End Sub

    ''' <summary>
    ''' update data in database table based on the input from excel file
    ''' override it to gain faster speed
    ''' </summary>
    ''' <param name="StatusLabel"></param>
    ''' <param name="SDS1">sqlDataSource Control represnt database table</param>
    ''' <param name="LV1">a listview control to display data, to get DataKeys by reference to it</param>
    ''' <param name="filepath">excel file path</param>
    ''' <param name="filename">excel file name</param>
    ''' <param name="flextsn">excel file suffix</param>
    Public Overrides Function dataUpdatedToDatabaseTablePerExcelData(ByRef StatusLabel As Label, ByRef SDS1 As SqlDataSource, ByRef LV1 As ListView, ByVal filepath As String, ByVal filename As String, ByVal flextsn As String) As String

        Dim msgRtrn As New StringBuilder

        Dim continue1 As Boolean = True
        Dim howManyRecordsUpdated As Integer = 0
        Dim excelconnectionstr As String
        Dim keyname00 As String = LV1.DataKeyNames(0)

        If flextsn = ".xls" Then
            excelconnectionstr = String.Format("provider=Microsoft.Jet.OLEDB.4.0; Data Source='{0}';" & "Extended Properties='Excel 8.0;HDR=YES;IMEX=1'", filepath & filename)
        Else
            excelconnectionstr = String.Format("provider=Microsoft.ACE.OLEDB.12.0; Data Source='{0}';" & "Extended Properties='Excel 12.0 Xml;HDR=YES;IMEX=1'", filepath & filename)
        End If

        Dim connexcl As New System.Data.OleDb.OleDbConnection(excelconnectionstr)
        connexcl.Open()

        Dim frstSheetName As String = "Sheet1$"


        'Try
        '    ' Get the name of the first worksheet:
        '    Dim dbSchema As DataTable = connexcl.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, Nothing)

        '    If (dbSchema Is Nothing OrElse dbSchema.Rows.Count < 1) Then
        '    End If

        '    frstSheetName = dbSchema.Rows(0)("TABLE_NAME").ToString()

        'Catch ex As Exception
        '    msgRtrn.AppendLine("<div style='color:red;'>" & ex.Message & "</div>")
        '    continue1 = False

        'End Try



        Dim excelslctsql As String = SDS1.SelectCommand
        excelslctsql = excelslctsql.Substring(0, excelslctsql.IndexOf(" FROM "))
        Dim dtADPexcel As OleDbDataAdapter = New OleDbDataAdapter(excelslctsql & " FROM [" & frstSheetName & "A1:" & rightmostColumn() & maxRowNumber.ToString() & "] WHERE (" & keyname00 & " Is Not Null)", connexcl)
        'Dim dtADPexcel As OleDbDataAdapter = New OleDbDataAdapter("Select * From [" & frstSheetName & "A1:BT15001] ", connexcl)
        Dim dataDS1 As DataSet = New DataSet

        Try
            dtADPexcel.Fill(dataDS1, "update")
        Catch ex As Exception
            'Page.ClientScript.RegisterClientScriptBlock(Me.GetType(), "script1", "<script language='javascript'>alert('Maybe you did not get the right excel file! row number limit is '" & maxRowNumber.ToString() & ");</script>", False)
            msgRtrn.AppendLine("<div style='color:red;'>" & ex.Message & " Maybe you did not get the right excel file! row number limit is " & maxRowNumber.ToString() & "</div>")
            continue1 = False
            dtADPexcel.Dispose()
        End Try




        Dim connstr As String = ConfigurationManager.ConnectionStrings(dbConnectionName).ProviderName & ConfigurationManager.ConnectionStrings(dbConnectionName).ConnectionString

        Dim conn As OleDbConnection = New OleDbConnection(connstr)
        conn.Open()

        Dim selectSQL As String = SDS1.SelectCommand


        Dim dtAdapter1 As OleDbDataAdapter = New OleDbDataAdapter(selectSQL, conn)


        'use aunto command generation mechnism to generate standard insert SQL clause
        Dim cmdbAccessCmdBuilder As New OleDbCommandBuilder(dtAdapter1)
        dtAdapter1.UpdateCommand = cmdbAccessCmdBuilder.GetUpdateCommand()


        If continue1 Then

            dtAdapter1.Fill(dataDS1, "tobeupdated")
            Dim keys(1) As DataColumn
            keys(0) = dataDS1.Tables("tobeupdated").Columns(keyname00)
            dataDS1.Tables("tobeupdated").PrimaryKey = keys

        End If



        If continue1 Then

            Dim msstring As String = String.Empty
            'check data validity
            msstring = dataValidityCheck(dataDS1.Tables("update"))

            If Not String.IsNullOrEmpty(msstring) Then
                continue1 = False
                'msgPopUP(msstring, StatusLabel, True)
                msgRtrn.AppendLine("<div style='color:red;'>" & msstring & "</div>")
                'StatusLabel.Text = msstring
                'Page.ClientScript.RegisterClientScriptBlock(Me.GetType(), "script1", "<script language='javascript'>alert('" & msstring & " Update is cancelled.');</script>", False)
            End If
        End If

        If continue1 Then
            Try

                dtADPexcel.FillLoadOption = LoadOption.Upsert
                dtADPexcel.Fill(dataDS1, "tobeupdated")

                For Each a As DataRow In dataDS1.Tables("tobeupdated").Select(Nothing, Nothing, System.Data.DataViewRowState.Added)
                    a.Delete()
                Next


                'do some routine calculation
                Dim hasException As Boolean = False

                Dim changedRows() As DataRow = dataDS1.Tables("tobeupdated").Select("int_line_no <> " & CInt(valueOf("intDummyLine")), Nothing, System.Data.DataViewRowState.ModifiedCurrent)
                msgRtrn.AppendLine(UpdateOrderCompletionPercentage1(conn, changedRows))
                'msgRtrn.AppendLine(calculateRSD1(conn, changedRows))
                msgRtrn.AppendLine(finishTime_exPlantDate_Span1(conn, changedRows, hasException))

                howManyRecordsUpdated = changedRows.Count
                'howManyRecordsUpdated = dataDS1.Tables("tobeupdated").Select(Nothing, Nothing, System.Data.DataViewRowState.ModifiedCurrent).Count

                If hasException Then
                    'exceptionR.Visible = True
                    CacheInsert("Oexception", 1)
                End If


                'considering additional days to add on original ex-plant date and quality issue related part
                msgRtrn.AppendLine(additionDaysOnExplantDate(conn, dataDS1.Tables("tobeupdated"), DataViewRowState.ModifiedCurrent))


                'dtAdapter1.UpdateBatchSize = 512  'access database does not support this
                dtAdapter1.Update(dataDS1, "tobeupdated")


                msgRtrn.AppendLine("Action of update upon the excel file is completed. The number of updated records is " & howManyRecordsUpdated)
                'msgPopUP("Action of update upon the excel file is completed. The number of updated records is " & howManyRecordsUpdated, StatusLabel, False, False)
            Catch ex As Exception
                'msgPopUP("Something wrong with the update operation." & ex.Message, StatusLabel, True, True)
                msgRtrn.AppendLine("<div style='color:red;'>" & "Something wrong with the update operation." & ex.Message & "</div>")
            End Try

        End If

        cmdbAccessCmdBuilder.Dispose()

        dataDS1.Dispose()

        dtAdapter1.Dispose()
        dtADPexcel.Dispose()

        If Not (conn.State = ConnectionState.Closed) Then
            conn.Close()
        End If
        conn.Dispose()

        If connexcl IsNot Nothing Then
            connexcl.Close()
            connexcl.Dispose()
        End If
        'connexcl = Nothing



        Return msgRtrn.ToString

    End Function





    ''' <summary>
    ''' delete records in database table based on the input from excel file
    ''' override the sub to get the faster processing
    ''' </summary>
    ''' <param name="StatusLabel"></param>
    ''' <param name="SDS1">sqlDataSource Control represnt database table</param>
    ''' <param name="LV1">a listview control to display data, to get DataKeys by reference to it</param>
    ''' <param name="filepath">excel file path</param>
    ''' <param name="filename">excel file name</param>
    ''' <param name="flextsn">excel file suffix</param>
    Public Overrides Function dataDeletedFromDatabaseTablePerExcelData(ByRef StatusLabel As Label, ByRef SDS1 As SqlDataSource, ByRef LV1 As ListView, ByVal filepath As String, ByVal filename As String, ByVal flextsn As String) As String

        Dim msgRtrn As New StringBuilder

        Dim continue1 As Boolean = True
        Dim howManyRecordsDeleted As Integer = 0
        Dim excelconnectionstr As String

        Dim keyname00 As String = LV1.DataKeyNames(0)

        If flextsn = ".xls" Then
            excelconnectionstr = String.Format("provider=Microsoft.Jet.OLEDB.4.0; Data Source='{0}';" & "Extended Properties='Excel 8.0;HDR=YES;IMEX=1'", filepath & filename)
        Else
            excelconnectionstr = String.Format("provider=Microsoft.ACE.OLEDB.12.0; Data Source='{0}';" & "Extended Properties='Excel 12.0 Xml;HDR=YES;IMEX=1'", filepath & filename)
        End If

        Dim connexcl As New System.Data.OleDb.OleDbConnection(excelconnectionstr)
        connexcl.Open()

        Dim frstSheetName As String = "Sheet1$"

        'Try
        '    ' Get the name of the first worksheet:
        '    Dim dbSchema As DataTable = connexcl.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, Nothing)

        '    If (dbSchema Is Nothing OrElse dbSchema.Rows.Count < 1) Then
        '    End If

        '    frstSheetName = dbSchema.Rows(0)("TABLE_NAME").ToString()
        'Catch ex As Exception
        '    msgRtrn.AppendLine("<div style='color:red;'>" & ex.Message & "</div>")
        '    continue1 = False

        'End Try

        Dim excelslctsql As String = SDS1.SelectCommand.ToString
        excelslctsql = excelslctsql.Substring(0, excelslctsql.IndexOf(" FROM "))
        Dim dtADPexcel As OleDbDataAdapter = New OleDbDataAdapter(excelslctsql & " FROM [" & frstSheetName & "A1:" & rightmostColumn() & maxRowNumber.ToString() & "] WHERE (" & keyname00 & " Is Not Null)", connexcl)
        Dim dataDS1 As DataSet = New DataSet
        Try
            dtADPexcel.Fill(dataDS1, "update")
        Catch ex As Exception
            'Page.ClientScript.RegisterClientScriptBlock(Me.GetType(), "script1", "<script language='javascript'>alert('Maybe you did not get the right excel file! Or columns missed, or row number exceeds 15001 .');</script>", False)
            continue1 = False
            'msgPopUP("Maybe you did not get the right excel file! Or columns missed, or row number exceeds " & maxRowNumber.ToString(), StatusLabel)
            msgRtrn.AppendLine("<div style='color:red;'>" & ex.Message & " Maybe you did not get the right excel file! row number limit is " & maxRowNumber.ToString() & "</div>")
        End Try

        dtADPexcel.Dispose()

        If connexcl IsNot Nothing Then
            connexcl.Close()
            connexcl.Dispose()
        End If
        'connexcl = Nothing

        Dim connstr As String = ConfigurationManager.ConnectionStrings(dbConnectionName).ProviderName & ConfigurationManager.ConnectionStrings(dbConnectionName).ConnectionString
        Dim conn As OleDbConnection = New OleDbConnection(connstr)


        Dim selectSQL As String = SDS1.SelectCommand
        If selectSQL.Contains(" ORDER BY ") Then
            selectSQL = selectSQL.Replace(" ORDER BY ", " WHERE CAST(int_line_no as VARCHAR(5)) In (" & lineListOwnedByUser() & ") ORDER BY ")
        Else
            selectSQL &= " WHERE CAST(int_line_no as VARCHAR(5)) In (" & lineListOwnedByUser() & ")"
        End If

        Dim dtAdapter1 As OleDbDataAdapter = New OleDbDataAdapter(selectSQL, conn)


        dtAdapter1.DeleteCommand = New OleDbCommand(SDS1.DeleteCommand, conn)
        Dim paramenter1 As OleDbParameter = dtAdapter1.DeleteCommand.Parameters.Add("@txt_order_key", OleDbType.VarChar, 22)
        paramenter1.SourceColumn = keyname00
        paramenter1.SourceVersion = DataRowVersion.Original



        If continue1 Then
            Try
                dtAdapter1.Fill(dataDS1, "tobeupdated")


                'delete all records requested in excel 


                Dim orders1 As DataTable = dataDS1.Tables("tobeupdated")
                Dim dltdetail As DataTable = dataDS1.Tables("update")
                Dim query1 = (From order1 In dltdetail.AsEnumerable() _
                            Select order1.Field(Of String)(keyname00)).Distinct()
                'check if there are duplicate records in the update file
                If continue1 Then

                    If query1.Count < dltdetail.Rows.Count Then
                        continue1 = False
                        'msgPopUP("Duplicate records exist in your update file.", StatusLabel)
                        msgRtrn.AppendLine("<div style='color:red;'> Duplicate records exist in your update file. </div>")
                    End If
                End If


                If continue1 Then

                    Dim orderList As New List(Of String)(query1)



                    For Each tobeDelete As DataRow In dataDS1.Tables("tobeupdated").Rows

                        'Dim a As Integer = 1

                        If orderList.Contains(tobeDelete.Item(keyname00)) Then
                            tobeDelete.Delete()
                            howManyRecordsDeleted += 1
                        End If


                    Next



                End If

            Catch ex As Exception
                continue1 = False

                msgRtrn.AppendLine("<div style='color:red;'>" & "Database maybe do not have the records you are asking to delete. " & ex.Message & "</div>")
            End Try

        End If

        If continue1 Then
            Try
                'delete records
                'dtAdapter1.UpdateBatchSize = 512
                dtAdapter1.Update(dataDS1, "tobeupdated")

                msgRtrn.AppendLine("Action on deletion upon excel file is completed. The number of deleted records is " & howManyRecordsDeleted)
            Catch ex As Exception
                msgRtrn.AppendLine("<div style='color:red;'> Something wrong when deleting. " & ex.Message & "</div>")
            End Try

        End If



        dataDS1.Dispose()

        dtAdapter1.Dispose()


        conn.Dispose()



        Return msgRtrn.ToString

    End Function


    ''' <summary>
    ''' check the data validity for the excel file
    ''' override the sub to get the faster processing
    ''' </summary>
    ''' <param name="StatusLabel"></param>
    ''' <param name="SDS1">sqlDataSource Control represnt database table</param>
    ''' <param name="LV1">a listview control to display data, to get DataKeys by reference to it</param>
    ''' <param name="filepath">excel file path</param>
    ''' <param name="filename">excel file name</param>
    ''' <param name="flextsn">excel file suffix</param>
    Public Overrides Function dataDeletedFromDatabaseTablePerExcelData(ByRef StatusLabel As Label, ByRef SDS1 As SqlDataSource, ByRef LV1 As ListView, ByVal filepath As String, ByVal filename As String, ByVal flextsn As String, ByRef continueOrnot As Boolean) As String

        continueOrnot = False

        Dim msgRtrn As New StringBuilder

        Dim continue1 As Boolean = True
        Dim excelconnectionstr As String
        Dim keyname00 As String = LV1.DataKeyNames(0)

        If flextsn = ".xls" Then
            excelconnectionstr = String.Format("provider=Microsoft.Jet.OLEDB.4.0; Data Source='{0}';" & "Extended Properties='Excel 8.0;HDR=YES;IMEX=1'", filepath & filename)
        Else
            excelconnectionstr = String.Format("provider=Microsoft.ACE.OLEDB.12.0; Data Source='{0}';" & "Extended Properties='Excel 12.0 Xml;HDR=YES;IMEX=1'", filepath & filename)
        End If

        Dim connexcl As New System.Data.OleDb.OleDbConnection(excelconnectionstr)
        connexcl.Open()

        Dim frstSheetName As String = "Sheet1$"


        'Try
        '    ' Get the name of the first worksheet:
        '    Dim dbSchema As DataTable = connexcl.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, Nothing)

        '    If (dbSchema Is Nothing OrElse dbSchema.Rows.Count < 1) Then
        '    End If

        '    frstSheetName = dbSchema.Rows(0)("TABLE_NAME").ToString()


        'Catch ex As Exception
        '    msgRtrn.AppendLine("<div style='color:red;'>" & ex.Message & "</div>")
        '    continue1 = False

        'End Try

        Dim excelslctsql As String = SDS1.SelectCommand.ToString
        excelslctsql = excelslctsql.Substring(0, excelslctsql.IndexOf(" FROM "))
        Dim dtADPexcel As OleDbDataAdapter = New OleDbDataAdapter(excelslctsql & " FROM [" & frstSheetName & "A1:" & rightmostColumn() & maxRowNumber.ToString() & "] WHERE (" & keyname00 & " Is Not Null)", connexcl)
        Dim dataDS1 As DataSet = New DataSet
        Try
            dtADPexcel.Fill(dataDS1, "update")
        Catch ex As Exception

            continue1 = False
            'msgPopUP("Maybe you did not get the right excel file! Or columns missed, or row number exceeds " & maxRowNumber.ToString(), StatusLabel)
            msgRtrn.AppendLine("<div style='color:red;'>" & ex.Message & " Maybe you did not get the right excel file! row number limit is " & maxRowNumber.ToString() & "</div>")
        End Try

        dtADPexcel.Dispose()

        If connexcl IsNot Nothing Then
            connexcl.Close()
            connexcl.Dispose()
        End If
        'connexcl = Nothing


        If continue1 Then

            Dim dltdetail As DataTable = dataDS1.Tables("update")

            Try

                Dim query1 = (From order1 In dltdetail.AsEnumerable() _
                            Select order1.Field(Of String)(keyname00)).Distinct()
                'check if there are duplicate records in the update file

                If continue1 Then

                    If query1.Count < dltdetail.Rows.Count Then
                        continue1 = False
                        'msgPopUP("Duplicate records exist in your update file.", StatusLabel)
                        msgRtrn.AppendLine("<div style='color:red;'> Duplicate records exist in your update file. </div>")
                    End If
                End If

            Catch ex As Exception
                continue1 = False
                msgRtrn.AppendLine("<div style='color:red;'>" & ex.Message & "</div>")
            End Try


            If continue1 Then

                Dim msstring As String = String.Empty
                'check data validity
                msstring = dataValidityCheck(dataDS1.Tables("update"))

                If Not String.IsNullOrEmpty(msstring) Then
                    continue1 = False
                    'msgPopUP(msstring, StatusLabel)
                    msgRtrn.AppendLine("<div style='color:red;'>" & msstring & "</div>")
                End If
            End If



        End If


        If continue1 Then

            Dim connstr As String = ConfigurationManager.ConnectionStrings(dbConnectionName).ProviderName & ConfigurationManager.ConnectionStrings(dbConnectionName).ConnectionString
            Dim conn As OleDbConnection = New OleDbConnection(connstr)
            conn.Open()

            Dim dtAdapter1 As OleDbDataAdapter = New OleDbDataAdapter("SELECT * FROM Esch_Na_tbl_orders ", conn)



            dtAdapter1.DeleteCommand = New OleDbCommand(SDS1.DeleteCommand, conn)
            Dim paramenter1 As OleDbParameter = dtAdapter1.DeleteCommand.Parameters.Add("@txt_order_key", OleDbType.VarChar, 22)
            paramenter1.SourceColumn = keyname00
            paramenter1.SourceVersion = DataRowVersion.Original

            Try

                dtAdapter1.Fill(dataDS1, "tobeupdated")

                Dim howManyRecordsDeleted As Integer = 0

                For Each r As DataRow In dataDS1.Tables("tobeupdated").Rows
                    r.Delete()
                    howManyRecordsDeleted += 1
                Next

                'dtAdapter1.UpdateBatchSize = 512
                dtAdapter1.Update(dataDS1, "tobeupdated")

                msgRtrn.AppendLine("During action on overwriting upon excel file, the number of deleted records is " & howManyRecordsDeleted)

                continueOrnot = True

            Catch ex As Exception
                continueOrnot = False
                msgRtrn.AppendLine("<div style='color:red;'>" & ex.Message & "</div>")
            End Try



            dtAdapter1.Dispose()

            If Not (conn.State = ConnectionState.Closed) Then
                conn.Close()
            End If
            conn.Dispose()

        End If


        dataDS1.Dispose()



        Return msgRtrn.ToString

    End Function


    ''' <summary>
    ''' Insert new record to database table from excel file
    ''' override the function to get the faster processing
    ''' </summary>
    ''' <param name="StatusLabel"></param>
    ''' <param name="SDS1">sqlDataSource Control represnt database table</param>
    ''' <param name="LV1">a listview control to display data, to get DataKeys by reference to it</param>
    ''' <param name="filepath">excel file path</param>
    ''' <param name="filename">excel file name</param>
    ''' <param name="flextsn">excel file suffix</param>
    ''' <returns>return warning message or others</returns>
    ''' 
    Public Overrides Function dataInsertionToDatabaseTableFromExcel(ByRef StatusLabel As Label, ByRef SDS1 As SqlDataSource, ByRef LV1 As ListView, ByVal filepath As String, ByVal filename As String, ByVal flextsn As String, ByRef continueOrnot As Boolean) As String

        Dim msgRtrn As New StringBuilder

        'Dim warningMessage As StringBuilder = New StringBuilder()
        Dim continue1 As Boolean = True
        Dim howManyRecordsInserted As Integer = 0
        Dim excelconnectionstr As String
        Dim keyname00 As String = LV1.DataKeyNames(0)
        If flextsn = ".xls" Then
            excelconnectionstr = String.Format("provider=Microsoft.Jet.OLEDB.4.0; Data Source='{0}';" & "Extended Properties='Excel 8.0;HDR=YES;IMEX=1'", filepath & filename)
        Else
            excelconnectionstr = String.Format("provider=Microsoft.ACE.OLEDB.12.0; Data Source='{0}';" & "Extended Properties='Excel 12.0 Xml;HDR=YES;IMEX=1'", filepath & filename)
        End If

        Dim connexcl As New System.Data.OleDb.OleDbConnection(excelconnectionstr)
        connexcl.Open()

        Dim frstSheetName As String = "Sheet1$"


        'Try
        '     Get the name of the first worksheet:
        '    Dim dbSchema As DataTable = connexcl.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, Nothing)

        '    If (dbSchema Is Nothing OrElse dbSchema.Rows.Count < 1) Then
        '    End If

        '    frstSheetName = dbSchema.Rows(0)("TABLE_NAME").ToString()

        'Catch ex As Exception
        '    msgRtrn.AppendLine("<div style='color:red;'>" & ex.Message & "</div>")
        '    continue1 = False
        'Finally

        'End Try

        Dim excelslctsql As String = SDS1.SelectCommand.ToString
        excelslctsql = excelslctsql.Substring(0, excelslctsql.IndexOf(" FROM "))
        Dim dtADPexcel As OleDbDataAdapter = New OleDbDataAdapter(excelslctsql & " FROM [" & frstSheetName & "A1:" & rightmostColumn() & maxRowNumber.ToString() & "] WHERE (" & keyname00 & " Is Not Null)", connexcl)
        Dim dataDS1 As DataSet = New DataSet



        Dim db As String = ConfigurationManager.ConnectionStrings(dbConnectionName).ConnectionString
        Dim connstr As String = ConfigurationManager.ConnectionStrings(dbConnectionName).ProviderName & db
        Dim conn As OleDbConnection = New OleDbConnection(connstr)
        conn.Open()

        Dim dtAdapter1 As OleDbDataAdapter = New OleDbDataAdapter(SDS1.SelectCommand, conn)

        Dim cmdbAccessCmdBuilder As New OleDbCommandBuilder(dtAdapter1)
        dtAdapter1.InsertCommand = cmdbAccessCmdBuilder.GetInsertCommand()

        If continue1 Then
            Try

                dtAdapter1.Fill(dataDS1, "tobeupdated")
                Dim keys(1) As DataColumn
                keys(0) = dataDS1.Tables("tobeupdated").Columns(keyname00)
                dataDS1.Tables("tobeupdated").PrimaryKey = keys



            Catch ex As Exception
                continue1 = False
                msgRtrn.AppendLine("<div style='color:red;'> Maybe there are duplicate records you are inserting. " & ex.Message & "</div>")
            End Try

        End If

        If continue1 Then
            Try
                'insert records
                dtADPexcel.FillLoadOption = LoadOption.Upsert
                dtADPexcel.Fill(dataDS1, "tobeupdated")



                'do some routine calculation
                Dim hasException As Boolean = False

                Dim changedRows() As DataRow = dataDS1.Tables("tobeupdated").Select("int_line_no <> " & CInt(valueOf("intDummyLine")), Nothing, System.Data.DataViewRowState.Added)
                msgRtrn.AppendLine(UpdateOrderCompletionPercentage1(conn, changedRows))
                'msgRtrn.AppendLine(calculateRSD1(conn, changedRows))
                msgRtrn.AppendLine(finishTime_exPlantDate_Span1(conn, changedRows, hasException))

                howManyRecordsInserted = dataDS1.Tables("tobeupdated").Select(Nothing, Nothing, System.Data.DataViewRowState.Added).Count


                If hasException Then
                    'exceptionR.Visible = True
                    CacheInsert("Oexception", 1)
                End If

                'considering additional days to add on original ex-plant date
                msgRtrn.AppendLine(additionDaysOnExplantDate(conn, dataDS1.Tables("tobeupdated")))


                'dtAdapter1.UpdateBatchSize = 512
                dtAdapter1.Update(dataDS1, "tobeupdated")
                msgRtrn.AppendLine("Action of insertion upon excel file is completed. The number of inserted records is " & howManyRecordsInserted)

            Catch ex As Exception
                msgRtrn.AppendLine("<div style='color:red;'> Something wrong when inserting. " & ex.Message & "</div>")
            End Try


        End If

        cmdbAccessCmdBuilder.Dispose()

        dataDS1.Dispose()

        dtAdapter1.Dispose()


        dtADPexcel.Dispose()

        If connexcl IsNot Nothing Then
            connexcl.Close()
            connexcl.Dispose()
        End If
        'connexcl = Nothing

        If Not (conn.State = ConnectionState.Closed) Then
            conn.Close()
        End If
        conn.Dispose()




        Return msgRtrn.ToString

    End Function





    ''' <summary>
    ''' check the data validity for the excel file
    ''' override the sub to get the faster processing
    ''' </summary>
    ''' <param name="StatusLabel"></param>
    ''' <param name="SDS1">sqlDataSource Control represnt database table</param>
    ''' <param name="LV1">a listview control to display data, to get DataKeys by reference to it</param>
    ''' <param name="filepath">excel file path</param>
    ''' <param name="filename">excel file name</param>
    ''' <param name="flextsn">excel file suffix</param>
    Public Function dataDeletedFromDatabaseTablePerExcelDataForOverwrite(ByRef StatusLabel As Label, ByRef SDS1 As SqlDataSource, ByRef LV1 As ListView, ByVal filepath As String, ByVal filename As String, ByVal flextsn As String, ByRef continueOrnot As Boolean) As String

        continueOrnot = False

        Dim msgRtrn As New StringBuilder

        Dim continue1 As Boolean = True
        Dim excelconnectionstr As String
        Dim keyname00 As String = LV1.DataKeyNames(0)

        If flextsn = ".xls" Then
            excelconnectionstr = String.Format("provider=Microsoft.Jet.OLEDB.4.0; Data Source='{0}';" & "Extended Properties='Excel 8.0;HDR=YES;IMEX=1'", filepath & filename)
        Else
            excelconnectionstr = String.Format("provider=Microsoft.ACE.OLEDB.12.0; Data Source='{0}';" & "Extended Properties='Excel 12.0 Xml;HDR=YES;IMEX=1'", filepath & filename)
        End If

        Using connexcl As New System.Data.OleDb.OleDbConnection(excelconnectionstr)

            'connexcl.Open()

            Dim frstSheetName As String = "Sheet1$"


            'Try
            '    ' Get the name of the first worksheet:
            '    Dim dbSchema As DataTable = connexcl.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, Nothing)

            '    frstSheetName = dbSchema.Rows(0)("TABLE_NAME").ToString()

            'Catch ex As Exception
            '    msgRtrn.AppendLine("<div style='color:red;'>" & ex.Message & "</div>")
            '    continue1 = False

            'End Try


            Dim excelslctsql As String = SDS1.SelectCommand.ToString
            excelslctsql = excelslctsql.Substring(0, excelslctsql.IndexOf(" FROM "))
            Using excleOleCommand As OleDbCommand = New OleDbCommand(excelslctsql & " FROM [" & frstSheetName & "A1:" & rightmostColumn() & maxRowNumber.ToString() & "] WHERE (" & keyname00 & " Is Not Null)", connexcl)
                Using dtADPexcel As OleDbDataAdapter = New OleDbDataAdapter(excleOleCommand)
                    Dim dataDS1 As DataSet = New DataSet
                    Try
                        dtADPexcel.Fill(dataDS1, "update")
                    Catch ex As Exception
                        'Page.ClientScript.RegisterClientScriptBlock(Me.GetType(), "script1", "<script language='javascript'>alert('Maybe you did not get the right excel file! Or columns missed, or row number exceeds 15001 .');</script>", False)
                        continue1 = False
                        'msgPopUP("Maybe you did not get the right excel file! Or columns missed, or row number exceeds " & maxRowNumber.ToString(), StatusLabel)
                        msgRtrn.AppendLine("<div style='color:red;'>" & ex.Message & " Maybe you did not get the right excel file! row number limit is " & maxRowNumber.ToString() & "</div>")
                    End Try



                    If continue1 Then

                        Dim dltdetail As DataTable = dataDS1.Tables("update")

                        Try
                            Dim query1 = (From order1 In dltdetail.AsEnumerable() _
                                        Select order1.Field(Of String)(keyname00)).Distinct()
                            'check if there are duplicate records in the update file

                            If continue1 Then

                                If query1.Count < dltdetail.Rows.Count Then
                                    continue1 = False
                                    'msgPopUP("Duplicate records exist in your update file.", StatusLabel)
                                    msgRtrn.AppendLine("<div style='color:red;'> Duplicate records exist in your update file. </div>")
                                End If
                            End If

                        Catch ex As Exception
                            continue1 = False
                            msgRtrn.AppendLine("<div style='color:red;'>" & ex.Message & "</div>")
                        End Try

                        If continue1 Then

                            Dim msstring As String = String.Empty
                            'check data validity
                            msstring = dataValidityCheck(dataDS1.Tables("update"))

                            If Not String.IsNullOrEmpty(msstring) Then
                                continue1 = False
                                'msgPopUP(msstring, StatusLabel)
                                msgRtrn.AppendLine("<div style='color:red;'>" & msstring & "</div>")
                            End If
                        End If


                    End If


                    If continue1 Then

                        Dim connstr As String = ConfigurationManager.ConnectionStrings(dbConnectionName).ProviderName & ConfigurationManager.ConnectionStrings(dbConnectionName).ConnectionString
                        Dim conn As OleDbConnection = New OleDbConnection(connstr)
                        conn.Open()

                        Dim dtAdapter1 As OleDbDataAdapter = New OleDbDataAdapter("SELECT * FROM Esch_Na_tbl_orders ", conn)


                        dtAdapter1.DeleteCommand = New OleDbCommand(SDS1.DeleteCommand, conn)
                        Dim paramenter1 As OleDbParameter = dtAdapter1.DeleteCommand.Parameters.Add("@txt_order_key", OleDbType.VarChar, 22)
                        paramenter1.SourceColumn = keyname00
                        paramenter1.SourceVersion = DataRowVersion.Original

                        Try


                            dtAdapter1.Fill(dataDS1, "tobeupdated")

                            Dim howManyRecordsDeleted As Integer = 0

                            For Each r As DataRow In dataDS1.Tables("tobeupdated").Rows
                                r.Delete()
                                howManyRecordsDeleted += 1
                            Next
                            'dtAdapter1.UpdateBatchSize = 512
                            dtAdapter1.Update(dataDS1, "tobeupdated")

                            msgRtrn.AppendLine("During action on overwriting upon excel file, the number of deleted records is " & howManyRecordsDeleted)
                            continueOrnot = True

                        Catch ex As Exception
                            continueOrnot = False
                            msgRtrn.AppendLine("<div style='color:red;'>" & ex.Message & "</div>")
                        End Try


                        dtAdapter1.Dispose()

                        If Not (conn.State = ConnectionState.Closed) Then
                            conn.Close()
                        End If
                        conn.Dispose()

                    End If

                    dataDS1.Dispose()

                End Using
            End Using
        End Using

        Return msgRtrn.ToString

    End Function


    ''' <summary>
    ''' Insert new record to database table from excel file
    ''' override the function to get the faster processing
    ''' </summary>
    ''' <param name="StatusLabel"></param>
    ''' <param name="SDS1">sqlDataSource Control represnt database table</param>
    ''' <param name="LV1">a listview control to display data, to get DataKeys by reference to it</param>
    ''' <param name="filepath">excel file path</param>
    ''' <param name="filename">excel file name</param>
    ''' <param name="flextsn">excel file suffix</param>
    ''' <returns>return warning message or others</returns>
    ''' 
    Public Function dataInsertionToDatabaseTableFromExcelForOverwrite(ByRef StatusLabel As Label, ByRef SDS1 As SqlDataSource, ByRef LV1 As ListView, ByVal filepath As String, ByVal filename As String, ByVal flextsn As String) As String

        Dim msgRtrn As New StringBuilder

        'Dim warningMessage As StringBuilder = New StringBuilder()
        Dim continue1 As Boolean = True
        Dim howManyRecordsInserted As Integer = 0
        Dim excelconnectionstr As String
        Dim keyname00 As String = LV1.DataKeyNames(0)
        If flextsn = ".xls" Then
            excelconnectionstr = String.Format("provider=Microsoft.Jet.OLEDB.4.0; Data Source='{0}';" & "Extended Properties='Excel 8.0;HDR=YES;IMEX=1'", filepath & filename)
        Else
            excelconnectionstr = String.Format("provider=Microsoft.ACE.OLEDB.12.0; Data Source='{0}';" & "Extended Properties='Excel 12.0 Xml;HDR=YES;IMEX=1'", filepath & filename)
        End If

        Using connexcl As New System.Data.OleDb.OleDbConnection(excelconnectionstr)
            connexcl.Open()

            Dim frstSheetName As String = "Sheet1$"


            Dim excelslctsql As String = SDS1.SelectCommand.ToString
            excelslctsql = excelslctsql.Substring(0, excelslctsql.IndexOf(" FROM "))
            Using excleOleCommand As OleDbCommand = New OleDbCommand(excelslctsql & " FROM [" & frstSheetName & "A1:" & rightmostColumn() & maxRowNumber.ToString() & "] WHERE (" & keyname00 & " Is Not Null)", connexcl)
                Using dtADPexcel As OleDbDataAdapter = New OleDbDataAdapter(excleOleCommand)
                    Dim dataDS1 As DataSet = New DataSet



                    Dim db As String = ConfigurationManager.ConnectionStrings(dbConnectionName).ConnectionString
                    Dim connstr As String = ConfigurationManager.ConnectionStrings(dbConnectionName).ProviderName & db
                    Dim conn As OleDbConnection = New OleDbConnection(connstr)
                    conn.Open()

                    Dim dtAdapter1 As OleDbDataAdapter = New OleDbDataAdapter(SDS1.SelectCommand, conn)

                    Dim cmdbAccessCmdBuilder As New OleDbCommandBuilder(dtAdapter1)
                    dtAdapter1.InsertCommand = cmdbAccessCmdBuilder.GetInsertCommand()

                    If continue1 Then
                        Try

                            dtAdapter1.Fill(dataDS1, "tobeupdated")
                            Dim keys(1) As DataColumn
                            keys(0) = dataDS1.Tables("tobeupdated").Columns(keyname00)
                            dataDS1.Tables("tobeupdated").PrimaryKey = keys



                        Catch ex As Exception
                            continue1 = False
                            msgRtrn.AppendLine("<div style='color:red;'> Maybe there are duplicate records you are inserting. " & ex.Message & "</div>")
                        End Try

                    End If

                    If continue1 Then
                        Try
                            'insert records
                            dtADPexcel.FillLoadOption = LoadOption.Upsert
                            dtADPexcel.Fill(dataDS1, "tobeupdated")



                            'do some routine calculation
                            Dim hasException As Boolean = False

                            Dim changedRows() As DataRow = dataDS1.Tables("tobeupdated").Select("int_line_no <> " & CInt(valueOf("intDummyLine")), Nothing, System.Data.DataViewRowState.Added)
                            msgRtrn.AppendLine(UpdateOrderCompletionPercentage1(conn, changedRows))

                            msgRtrn.AppendLine(finishTime_exPlantDate_Span1(conn, changedRows, hasException))

                            howManyRecordsInserted = dataDS1.Tables("tobeupdated").Select(Nothing, Nothing, System.Data.DataViewRowState.Added).Count


                            If hasException Then
                                CacheInsert("Oexception", 1)
                            End If

                            'considering additional days to add on original ex-plant date
                            msgRtrn.AppendLine(additionDaysOnExplantDate(conn, dataDS1.Tables("tobeupdated")))
							'FANAR+ GO-LIVE, on Nov.9.2016  need UPDATE these information according to order status, if they are 'NEWnew'
							msgRtrn.AppendLine(PackageOrMiscellaneous(conn))
							msgRtrn.AppendLine(DoAssignScrewDieAndFDA())

                            'dtAdapter1.UpdateBatchSize = 512
                            dtAdapter1.Update(dataDS1, "tobeupdated")
                            msgRtrn.AppendLine("Action of insertion upon excel file is completed. The number of inserted records is " & howManyRecordsInserted)

                        Catch ex As Exception
                            msgRtrn.AppendLine("<div style='color:red;'> Something wrong when inserting. " & ex.Message & "</div>")
                        End Try


                    End If

                    cmdbAccessCmdBuilder.Dispose()

                    dataDS1.Dispose()

                    dtAdapter1.Dispose()


                    If Not (conn.State = ConnectionState.Closed) Then
                        conn.Close()
                    End If
                    conn.Dispose()

                End Using
            End Using
        End Using




        Return msgRtrn.ToString

    End Function




    ''' <summary>
    ''' Insert new record to database table from excel file
    ''' override the function to get the faster processing
    ''' </summary>
    ''' <param name="StatusLabel"></param>
    ''' <param name="SDS1">sqlDataSource Control represnt database table</param>
    ''' <param name="LV1">a listview control to display data, to get DataKeys by reference to it</param>
    ''' <param name="filepath">excel file path</param>
    ''' <param name="filename">excel file name</param>
    ''' <param name="flextsn">excel file suffix</param>
    ''' <returns>return warning message or others</returns>
    ''' 
    Public Overrides Function dataInsertionToDatabaseTableFromExcel(ByRef StatusLabel As Label, ByRef SDS1 As SqlDataSource, ByRef LV1 As ListView, ByVal filepath As String, ByVal filename As String, ByVal flextsn As String) As String

        Dim msgRtrn As New StringBuilder

        'Dim warningMessage As StringBuilder = New StringBuilder()
        Dim continue1 As Boolean = True
        Dim howManyRecordsInserted As Integer = 0
        Dim excelconnectionstr As String
        Dim keyname00 As String = LV1.DataKeyNames(0)

        If flextsn = ".xls" Then
            excelconnectionstr = String.Format("provider=Microsoft.Jet.OLEDB.4.0; Data Source='{0}';" & "Extended Properties='Excel 8.0;HDR=YES;IMEX=1'", filepath & filename)
        Else
            excelconnectionstr = String.Format("provider=Microsoft.ACE.OLEDB.12.0; Data Source='{0}';" & "Extended Properties='Excel 12.0 Xml;HDR=YES;IMEX=1'", filepath & filename)
        End If

        Dim connexcl As New System.Data.OleDb.OleDbConnection(excelconnectionstr)
        connexcl.Open()

        Dim frstSheetName As String = "Sheet1$"


        'Try
        '    ' Get the name of the first worksheet:
        '    Dim dbSchema As DataTable = connexcl.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, Nothing)

        '    If (dbSchema Is Nothing OrElse dbSchema.Rows.Count < 1) Then
        '    End If

        '    frstSheetName = dbSchema.Rows(0)("TABLE_NAME").ToString()

        'Catch ex As Exception
        '    msgRtrn.AppendLine("<div style='color:red;'>" & ex.Message & "</div>")
        '    continue1 = False

        'End Try

        Dim excelslctsql As String = SDS1.SelectCommand.ToString
        excelslctsql = excelslctsql.Substring(0, excelslctsql.IndexOf(" FROM "))
        Dim dtADPexcel As OleDbDataAdapter = New OleDbDataAdapter(excelslctsql & " FROM [" & frstSheetName & "A1:" & rightmostColumn() & maxRowNumber.ToString() & "] WHERE (" & keyname00 & " Is Not Null)", connexcl)
        'Dim dtADPexcel As OleDbDataAdapter = New OleDbDataAdapter(excelslctsql & " FROM [" & frstSheetName & "] ", connexcl)
        Dim dataDS1 As DataSet = New DataSet
        Try
            dtADPexcel.Fill(dataDS1, "update")

        Catch ex As Exception
            continue1 = False
            'warningMessage.Append("Maybe you did not get the right excel file! Or columns missed, or inserted row number exceeds " & maxRowNumber.ToString() & "<br />")
            msgRtrn.AppendLine("<div style='color:red;'>" & ex.Message & " Maybe you did not get the right excel file! row number limit is " & maxRowNumber.ToString() & "</div>")
        End Try


        If continue1 Then

            Dim msstring As String = String.Empty
            'check data validity
            msstring = dataValidityCheck(dataDS1.Tables("update"))

            If Not String.IsNullOrEmpty(msstring) Then
                continue1 = False
                'msgPopUP(msstring, StatusLabel)
                msgRtrn.AppendLine("<div style='color:red;'>" & msstring & "</div>")
            End If
        End If



        Dim connstr As String = ConfigurationManager.ConnectionStrings(dbConnectionName).ProviderName & ConfigurationManager.ConnectionStrings(dbConnectionName).ConnectionString
        Dim conn As OleDbConnection = New OleDbConnection(connstr)
        conn.Open()

        Dim dtAdapter1 As OleDbDataAdapter = New OleDbDataAdapter(SDS1.SelectCommand, conn)

        Dim cmdbAccessCmdBuilder As New OleDbCommandBuilder(dtAdapter1)
        dtAdapter1.InsertCommand = cmdbAccessCmdBuilder.GetInsertCommand()

        If continue1 Then
            Try

                dtAdapter1.Fill(dataDS1, "tobeupdated")
                Dim keys(1) As DataColumn
                keys(0) = dataDS1.Tables("tobeupdated").Columns(keyname00)
                dataDS1.Tables("tobeupdated").PrimaryKey = keys


                'check if there is record in both excel file and database
                Dim orders1 As DataTable = dataDS1.Tables("tobeupdated")
                Dim insertdetail As DataTable = dataDS1.Tables("update")



                'check if there are duplicate records in the update file
                If continue1 Then
                    Dim query1 = (From order1 In insertdetail.AsEnumerable() _
                            Select order1.Field(Of String)(keyname00)).Distinct()
                    If Not (query1.Count = insertdetail.Rows.Count) Then
                        continue1 = False
                        'msgPopUP("Duplicate rows exist in your update file.", StatusLabel)
                        msgRtrn.AppendLine("<div style='color:red;'>" & "Duplicate rows exist in your update file. " & "</div>")
                    End If
                End If


                If continue1 Then

                    Dim listFrmTableExcel = From order1 In insertdetail.AsEnumerable()
                                            Select order1.Field(Of String)(keyname00)

                    Dim listToTable = From order1 In orders1.AsEnumerable()
                                            Select order1.Field(Of String)(keyname00)

                    Dim orderKeylist = From a In listFrmTableExcel Where Not listToTable.Contains(a)
                                        Select a


                    For Each a As String In orderKeylist
                        Dim toupdateRows() As DataRow = insertdetail.Select("txt_order_key = '" & a & "'")
                        dataDS1.Tables("tobeupdated").Rows.Add(toupdateRows(0).ItemArray)
                    Next



                End If



            Catch ex As Exception
                continue1 = False
                'msgPopUP("Maybe there are duplicate records you are inserting.", StatusLabel)
                'Page.ClientScript.RegisterClientScriptBlock(Me.GetType(), "script1", "<script language='javascript'>alert('Database maybe do not have the records you are asking to delete.');</script>", False)
                msgRtrn.AppendLine("<div style='color:red;'> Maybe there are duplicate records you are inserting. " & ex.Message & "</div>")
            End Try

        End If

        If continue1 Then
            Try
                'insert records
                'dtADPexcel.FillLoadOption = LoadOption.Upsert
                'dtADPexcel.Fill(dataDS1, "tobeupdated")



                'do some routine calculation
                Dim hasException As Boolean = False

                Dim changedRows() As DataRow = dataDS1.Tables("tobeupdated").Select("int_line_no <> " & CInt(valueOf("intDummyLine")), Nothing, System.Data.DataViewRowState.Added)
                msgRtrn.AppendLine(UpdateOrderCompletionPercentage1(conn, changedRows))
                'msgRtrn.AppendLine(calculateRSD1(conn, changedRows))
                msgRtrn.AppendLine(finishTime_exPlantDate_Span1(conn, changedRows, hasException))

                'howManyRecordsInserted = changedRows.Count
                howManyRecordsInserted = dataDS1.Tables("tobeupdated").Select(Nothing, Nothing, System.Data.DataViewRowState.Added).Count

                If hasException Then
                    CacheInsert("Oexception", 1)
                End If

                'considering additional days to add on original ex-plant date
                msgRtrn.AppendLine(additionDaysOnExplantDate(conn, dataDS1.Tables("tobeupdated")))
				
				'FANAR+ GO-LIVE, on Nov.9.2016  need UPDATE these information according to order status, if they are 'NEWnew'
				msgRtrn.AppendLine(PackageOrMiscellaneous(conn))
				msgRtrn.AppendLine(DoAssignScrewDieAndFDA())

                'dtAdapter1.UpdateBatchSize = 512
                dtAdapter1.Update(dataDS1, "tobeupdated")
                msgRtrn.AppendLine("Action of insertion upon excel file is completed. The number of inserted records is " & howManyRecordsInserted)

            Catch ex As Exception
                'msgPopUP("Something wrong when inserting. " & ex.Message, StatusLabel)
                msgRtrn.AppendLine("<div style='color:red;'> Something wrong when inserting. " & ex.Message & "</div>")
            End Try


        End If

        cmdbAccessCmdBuilder.Dispose()

        dataDS1.Dispose()

        dtAdapter1.Dispose()


        dtADPexcel.Dispose()

        If Not (conn.State = ConnectionState.Closed) Then
            conn.Close()
        End If
        conn.Dispose()

        If connexcl IsNot Nothing Then
            connexcl.Close()
            connexcl.Dispose()
        End If
        'connexcl = Nothing

        Return msgRtrn.ToString

    End Function






    ''' <summary>
    ''' to check data validity before further processing 
    ''' </summary>
    Public Overrides Function dataValidityCheck(ByRef dtTbl As DataTable) As String



        Dim msstring As String = String.Empty

        Try

            For Each rowexcel As DataRow In dtTbl.Rows
                If Not (IsNumeric(rowexcel("flt_order_qty")) AndAlso IsNumeric(rowexcel("flt_unallocate_qty")) AndAlso IsNumeric(rowexcel("planned_production_qty"))) Then
                    msstring = "There is something wrong with row txt_order_key=" & rowexcel("txt_order_key") & "! Maybe flt_order_qty,flt_unallocate_qty or planned_production_qty are not valid numbers."
                    Exit For
                End If

                If Not (DBNull.Value.Equals(rowexcel("txt_lot_no")) OrElse String.IsNullOrEmpty(rowexcel("txt_lot_no")) OrElse rowexcel("txt_lot_no").ToString.Length = 8 OrElse rowexcel("txt_lot_no").ToString.Length = 9 ) Then
                    msstring = "'There is something wrong with row txt_order_key=" & rowexcel("txt_order_key") & "! Maybe  txt_lot_no is not 8 digits."
                    Exit For
                End If

                If DBNull.Value.Equals(rowexcel("int_line_no")) OrElse Not (IsNumeric(rowexcel("int_line_no").ToString()) AndAlso CInt(rowexcel("int_line_no")) >= 1 AndAlso CInt(rowexcel("int_line_no")) <= 333) Then
                    msstring = "'There is something wrong with row txt_order_key=" & rowexcel("txt_order_key") & "! Maybe int_line_no  is not valid number."
                    Exit For
                End If


                If Not (IsDate(rowexcel("dat_etd")) AndAlso IsDate(rowexcel("dat_rdd")) AndAlso IsDate(rowexcel("dat_start_date"))) Then
                    msstring = "There is something wrong with row txt_order_key=" & rowexcel("txt_order_key") & "! Maybe dat_rdd,dat_etd,dat_start_date are not valid dates ."
                    Exit For
                End If


            Next

        Catch
            msstring = "There is something wrong with the data validity."
        End Try

        Return msstring

    End Function

    Protected Sub LV1_ItemDeleting(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.ListViewDeleteEventArgs) Handles LV1.ItemDeleting

        Dim messageText As String = String.Empty

        Dim item As ListViewItem = LV1.Items(e.ItemIndex)

        If Not lineListOwnedByUser.Contains("'" & CType(item.FindControl("int_line_noLabel"), Label).Text & "'") Then
            messageText = " This line is not owned by current user (" & userIden() & ")!"

        End If


        If Not String.IsNullOrEmpty(messageText) Then
            e.Cancel = True
            Message.Text = messageText
        End If

    End Sub





    ''' <summary>
    ''' You can do some checking on data validity; if some exception is thrown out, the updating would be cancelled
    ''' </summary>
    Protected Sub LV1_ItemUpdating(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.ListViewUpdateEventArgs) Handles LV1.ItemUpdating


        Dim messageText As StringBuilder = New StringBuilder()



        'check date related data
        If IsDate(e.NewValues("dat_etd")) AndAlso IsDate(e.NewValues("dat_rdd")) AndAlso IsDate(e.NewValues("dat_start_date")) Then
        Else
            messageText.Append(" Date is needed in field dat_etd,dat_rdd or dat_start_time <br />")
        End If

        'check number related.
        If IsNumeric(e.NewValues("flt_order_qty")) AndAlso IsNumeric(e.NewValues("planned_production_qty")) AndAlso IsNumeric(e.NewValues("flt_unallocate_qty")) Then
        Else
            messageText.Append(" Number is needed in field flt_order_qty,planned_production_qty or flt_unallocate_qty <br />")
        End If

        'check txt_lot_no
        If String.IsNullOrEmpty(e.NewValues("txt_lot_no")) OrElse (Regex.IsMatch(e.NewValues("txt_lot_no").ToString, "\d{8,8}")) Then
        Else
            messageText.Append(" The value in field int_line_no should be 8 digits <br />")
        End If

        'check int
        If String.IsNullOrEmpty(e.NewValues("int_line_no")) Then
            messageText.Append(" Empty value is not allowed in field int_line_no <br />")
        Else
            If IsNumeric(e.NewValues("int_line_no")) Then
                Dim LineFound As Boolean = False
                Dim newLine As Integer = CInt(e.NewValues("int_line_no"))
                For Each line As Integer In arrayOfLines()
                    If newLine = line Then
                        LineFound = True
                        Exit For
                    End If
                Next

                If (Not LineFound) AndAlso (newLine <> CInt(valueOf("intDummyLine"))) Then
                    messageText.Append(newLine & " is not a valid production line. <br />")
                End If
            Else
                messageText.Append(e.NewValues("int_line_no") & " is not a valid production line. <br />")
            End If

        End If

        If Not lineListOwnedByUser.Contains("'" & e.OldValues("int_line_no").ToString & "'") Then
            messageText.Append(" This line is not owned by current user (" & userIden() & ")!")
        End If


        If messageText.Length > 0 Then
            e.Cancel = True

        Else

            messageText.AppendLine(oneTimeRoutinePlanning(Nothing, e)) 'd

        End If

        Message.Text = messageText.ToString()


    End Sub



    ''' <summary>
    ''' Click to view the exception report
    ''' </summary>
    Protected Sub exception_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles exceptionR.Click
        CacheRemove("Oexception")
        Response.Redirect("~/Makerelated/exception.aspx")
    End Sub



    Protected Function oneTimeRoutinePlanning(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.ListViewUpdateEventArgs) As String

        Dim msgRtrn As New StringBuilder

        Dim connstr As String = ConfigurationManager.ConnectionStrings(dbConnectionName).ProviderName & ConfigurationManager.ConnectionStrings(dbConnectionName).ConnectionString
        'applying some planning rules after importing data to database
        Dim conn As OleDbConnection = New OleDbConnection(connstr)


        Dim connParam As OleDbConnection = New OleDbConnection(ConfigurationManager.ConnectionStrings(dbConnForParam).ProviderName & ConfigurationManager.ConnectionStrings(dbConnForParam).ConnectionString)

        Try




            Dim filterCondition As String = " WHERE txt_line_no = '" & e.NewValues("int_line_no") & "'"

            '1,RSD ========================
            If True Then

                Dim dtUpdateTo As OleDbDataAdapter = New OleDbDataAdapter("SELECT txt_order_key,txt_item_no,dat_rdd,dat_etd,txt_currency,txt_destination,txt_ship_method,txt_end_user FROM Esch_Na_tbl_orders WHERE txt_order_key = '" & e.OldValues("txt_order_no") & "-" & e.OldValues("txt_order_line_no") & "'", conn)

                Dim dtTable As DataTable = New DataTable

                dtUpdateTo.Fill(dtTable)

                Dim r() As DataRow = dtTable.Select(Nothing)
                If r.Count > 0 Then

                    For Each a As DataColumn In dtTable.Columns
                        If Not String.IsNullOrEmpty(e.NewValues(a.ColumnName)) Then
                            r(0).Item(a.ColumnName) = e.NewValues(a.ColumnName)
                        End If
                    Next

                End If


                msgRtrn.AppendLine(calculateRSD1(conn, r))
                e.NewValues("dat_etd") = r(0).Item("dat_etd")


                dtTable.Dispose()
                dtUpdateTo.Dispose()


            End If

            '2,UpdateOrderCompletionPercentage1 ========================
            If True Then

                Dim dtUpdateTo As OleDbDataAdapter = New OleDbDataAdapter("SELECT txt_order_key,flt_actual_qty_man,flt_actual_completed,planned_production_qty FROM Esch_Na_tbl_orders WHERE txt_order_key = '" & e.OldValues("txt_order_no") & "-" & e.OldValues("txt_order_line_no") & "'", conn)

                Dim dtTable As DataTable = New DataTable

                dtUpdateTo.Fill(dtTable)

                Dim r() As DataRow = dtTable.Select(Nothing)
                If r.Count > 0 Then

                    For Each a As DataColumn In dtTable.Columns
                        If Not String.IsNullOrEmpty(e.NewValues(a.ColumnName)) Then
                            r(0).Item(a.ColumnName) = e.NewValues(a.ColumnName)
                        End If
                    Next

                End If


                msgRtrn.AppendLine(UpdateOrderCompletionPercentage1(conn, r))
                e.NewValues("flt_actual_completed") = r(0).Item("flt_actual_completed")


                dtTable.Dispose()
                dtUpdateTo.Dispose()


            End If


            '3,finishTime_exPlantDate_Span1 ========================
            If True Then

                Dim dtUpdateTo As OleDbDataAdapter = New OleDbDataAdapter("SELECT * FROM Esch_Na_tbl_orders WHERE txt_order_key = '" & e.OldValues("txt_order_no") & "-" & e.OldValues("txt_order_line_no") & "'", conn)

                Dim dtTable As DataTable = New DataTable

                dtUpdateTo.Fill(dtTable)

                Dim r() As DataRow = dtTable.Select(Nothing)
                If r.Count > 0 Then

                    For Each a As DataColumn In dtTable.Columns
                        If Not String.IsNullOrEmpty(e.NewValues(a.ColumnName)) Then
                            r(0).Item(a.ColumnName) = e.NewValues(a.ColumnName)
                        End If
                    Next

                End If

                Dim hasException As Boolean = False
                msgRtrn.AppendLine(finishTime_exPlantDate_Span1(conn, r, hasException))

                If hasException Then
                    'exceptionR.Visible = True
                    CacheInsert("Oexception", 1)
                End If
                'considering additional days to add on original ex-plant date
                msgRtrn.AppendLine(additionDaysOnExplantDate(conn, dtTable, DataViewRowState.ModifiedCurrent))

                e.NewValues("flt_working_hours") = r(0).Item("flt_working_hours")
                e.NewValues("int_change_over_time") = r(0).Item("int_change_over_time")
				'e.NewValues("dat_finish_date") = r(0).Item("dat_finish_date")
				'FANAR+ GO-LIVE, on Nov.9.2016  need calculate start time based on finished time, backward calculation
                e.NewValues("dat_new_explant") = r(0).Item("dat_new_explant")
                e.NewValues("int_span") = r(0).Item("int_span")

  
                dtTable.Dispose()
                dtUpdateTo.Dispose()


            End If

        Catch ex As Exception
            msgRtrn.AppendLine("<div style='color:red;'>" & ex.Message & "</div>")
        End Try


        connParam.Dispose()
        conn.Dispose()

        Return msgRtrn.ToString


    End Function


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




    Protected Sub btchCrtn_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btchCrtn.Click
        Response.Redirect("~/interface/batchCreationAndUpload.aspx")
    End Sub

    Protected Sub mtiOE_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles mtiOE.Click
        Response.Redirect("~/dragDrop/normalOP/MTI.aspx")
    End Sub

    Protected Sub cmbnPrdctn_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles cmbnPrdctn.Click
        Response.Redirect("~/dragDrop/normalOP/combineProduction.aspx")
    End Sub






    Protected Function deleteSome(Optional ByVal condition As String = " true ") As String
        Dim orderLines As Integer

        Dim connstr As String = ConfigurationManager.ConnectionStrings(dbConnectionName).ProviderName & ConfigurationManager.ConnectionStrings(dbConnectionName).ConnectionString
        Dim conn As OleDbConnection = New OleDbConnection(connstr)
        orderLines = deleteUponCondition(conn, condition)

        conn.Dispose()

        Return "Accordign to the filter condition, the number of deleted  order lines is " & orderLines

    End Function


    Protected Sub delCS_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles delCS.Click

        Dim condition As String = SDS1.FilterExpression
        If Not String.IsNullOrEmpty(condition) Then
            condition = "True And " & SDS1.FilterExpression
        Else
            condition = " True "
        End If

        Dim msg As String = deleteSome(condition)

        clrfltr1_Click(Nothing, System.EventArgs.Empty) 'clear the filter for LV1 to show all the data

        msgPopUP(msg, StatusLabel, False, False)

    End Sub


    Protected Sub Page_PreInit(sender As Object, e As EventArgs) Handles Me.PreInit
        'if for shanghai plant, go to shanghai page because shanghai page has different layout for this page
        If valueOf("strOrderDetailStyle").ToUpper.StartsWith("LIKE_SHANGHAI") Then
            Response.Redirect("OrderDetail2.aspx")
        End If
    End Sub
End Class

