Imports System.IO
Imports System.Data.OleDb
Imports System.Data
Imports orderRelatedPlanning


Partial Class interface_downloadOrders
    Inherits orderRelatedPlanning



    'Inherits manyOperationsBeforeResoveBackToDatabase

    ' IMPorders_Click -> process_text_file      -> planningActionsAfterInterfaceFilesGot
    '                     |=> dataInTextToTable


    ''' <summary>
    ''' import batch status update information from OPM ftp server to local folder and consider to share the same files with S$FS department
    ''' read the data in the txt file and use the data to update relevant data table in access database
    ''' </summary>
    Protected Sub IMPorders_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles IMPorders.Click

        '

        Dim msgRtrn As New StringBuilder
        Dim continues As Boolean = True

        Dim userName As String = lockKeyTable(priority.ImportNewOrderOrBatch)
        If Not String.IsNullOrEmpty(userName) Then
            continues = False
            msgRtrn.Append("<div style='color:red;font-size:150%;'>" & userName & " is using the order detail table" & "</div>")
        End If

        If continues Then
            msgRtrn.Append("<div style = 'color:red;'>" & preCheckBeforeOrdersImportation(continues) & "</div>")
        End If


        Dim localpath As String = ConfigurationManager.AppSettings("EDIfilesFolder")
        If Not localpath.EndsWith("\") Then localpath &= "\"
        localpath &= "interfaceData\"
        If Not Directory.Exists(localpath) Then
            Try
                Directory.CreateDirectory(localpath)
            Catch
                'msgPopUP("You might not have right to create folder: " & localpath, Label1)
                continues = False
                'errMessage.Append("You might not have right to create folder: " & localpath)
                msgRtrn.AppendLine("<div style = 'color:red;'>" & "You might not have right to create folder: " & localpath & "</div>")
            End Try
        End If




        'get new&revision order status files from ftp server 
        Dim localpathsubfolder As String = localpath & "NewRevisionOrder\"
        Dim workingDirectory As String = valueOf("strOpenOrderPath")
 

        'get user name and password and ftp server IP to log on to server to get files by ftp method
        Dim organization As String = valueOf("strOrganization")
        Dim prefix As String = valueOf("strOpenOrderPrefix")
        Dim serverUri As String = valueOf("strFTPserverIP")
        If Not serverUri.StartsWith("ftp://") Then serverUri = "ftp://" & serverUri
        If Not serverUri.EndsWith("/") Then serverUri &= "/"

        If Not workingDirectory.EndsWith("/") Then workingDirectory &= "/"


        Dim olist As List(Of String) = New List(Of String)



        Dim ftpoperation As FTPcls = New FTPcls(valueOf("strFTP_ID"), valueOf("strFTP_PW"), serverUri & workingDirectory)
        Dim errMsg As String = String.Empty
        olist = ftpoperation.FileListInDirContains(prefix & "." & organization, errMsg)

        If Not String.IsNullOrEmpty(errMsg) Then
            'errMessage.Append("<br />" & serverUri & workingDirectory & errMsg)
            msgRtrn.AppendLine("<div style = 'color:red;'>" & "You might not have right to create folder: " & serverUri & workingDirectory & errMsg & "</div>")
        End If

        If continues Then

            olist.Sort(StringComparer.Ordinal)

            If olist.Count > 1 Then
                'msge.Append("<br /> " & olist.Count & " files are downloaded from ftp server " & serverUri & workingDirectory & " .<br />")
                msgRtrn.AppendLine("<div style = 'color:black;'>" & olist.Count & " files are downloaded from ftp server " & serverUri & workingDirectory & "</div>")
            Else
                'msge.Append("<br /> " & olist.Count & " file is downloaded from ftp server " & serverUri & workingDirectory & " .<br />")
                msgRtrn.AppendLine("<div style = 'color:black;'>" & olist.Count & " file is downloaded from ftp server " & serverUri & workingDirectory & "</div>")
            End If

            'down the file and delete the file on the ftp server
            For Each filename As String In olist
                If ftpoperation.DownFileFrmServer(localpath & filename, filename).ToLower() = "true" Then
                    ftpoperation.DeleteFileOnServer(filename)
                End If

                'msge.Append(filename & "<br />")
                msgRtrn.AppendLine("<div style = 'color:black;'>" & filename & "</div>")
            Next

        End If




        'process these files further===========
        If continues Then
            Dim errormsg As String = String.Empty
            msgRtrn.AppendLine(process_text_file(errormsg, continues))
            msgRtrn.AppendLine("<div style = 'color:red;'>" & errormsg & "</div>")
        End If




        If continues Then
            'finally archive these files
            Try
                If Not Directory.Exists(localpathsubfolder) Then Directory.CreateDirectory(localpathsubfolder)
                'archive those files downloaded from the first ftp server
                For Each file2 In olist
                    File.Copy(localpath & file2, localpathsubfolder & file2, True)
                    File.Delete(localpath & file2)
                Next


            Catch ex As Exception
                continues = False
                msgRtrn.AppendLine("<div style = 'color:red;'>" & "Archive: " & ex.Message & "</div>")
            End Try
        End If


        Dim msgIncAll As String = msgRtrn.ToString

        If continues Then
            'execute several planning action
            planningActionsAfterInterfaceFilesGot(msgIncAll)
        End If

        If String.IsNullOrEmpty(userName) Then
            unlockKeyTable(priority.ImportNewOrderOrBatch)
        End If

        Label1.Text = msgIncAll


    End Sub


    ''' <summary>
    ''' process those order data information in text files
    ''' import text file one by one in the order from the oldest to the newest
    ''' </summary>
    ''' <returns></returns>
    Private Function process_text_file(ByRef errormsg As String, ByRef continues As Boolean) As String

        Dim msge As System.Text.StringBuilder = New System.Text.StringBuilder()

        Dim begintime As Date = Date.Now

        Dim processingTxt As String = "order.txt"
        Dim tablename As String = "Esch_Sh_Esch_Sh_tbl_orders_new_revision_from_OPM"
        Dim localpath As String = ConfigurationManager.AppSettings("EDIfilesFolder")
        If Not localpath.EndsWith("\") Then localpath &= "\"
        localpath &= "interfaceData\"
        If Not Directory.Exists(localpath) Then
            Try
                Directory.CreateDirectory(localpath)
            Catch
                'msgPopUP("You might not have right to create folder: " & localpath, Label1)
                errormsg &= "You might not have right to create folder: " & localpath
                Return String.Empty
            End Try
        End If

        Dim files = From file In Directory.EnumerateFiles(localpath) Where file.ToLower().Contains(valueOf("strOpenOrderPrefix").ToLower()) Order By file Ascending

        If files.Count = 0 Then
            continues = False
            errormsg &= "<div style='color:red;'>You do not have one EDI file.</div>"
        End If


        Dim db As String = ConfigurationManager.ConnectionStrings(dbConnectionName).ConnectionString
        Dim connstr As String = ConfigurationManager.ConnectionStrings(dbConnectionName).ProviderName & db
        Dim conn As OleDbConnection = New OleDbConnection(connstr)

        conn.Open()

        Dim command As New OleDbCommand("SELECT txtFieldName FROM Esch_Sh_tbl_interface_mapping WHERE intImportOrderMapping > 0 ORDER BY intImportOrderMapping", conn)
        Dim reader As OleDbDataReader = command.ExecuteReader()
        Dim mapping1() As String = {String.Empty}, icount As Integer = 0
        While reader.Read()
            ReDim Preserve mapping1(icount)
            mapping1(icount) = reader("txtFieldName")
            icount += 1
        End While
        reader.Close()

        'clear the content remaining in the table Esch_Sh_Esch_Sh_tbl_orders_new_revision_from_OPM
        Try
            command.CommandText = "DELETE * FROM " & tablename
            command.ExecuteNonQuery()
        Catch ex As Exception
            msge.Append("<br /><div style='color:red'>Delete records from table " & tablename & " -- " & ex.Message & "</div>")
        End Try

        command.Dispose()


        'create and edit file schema.ini in the same folder
        If Not File.Exists(localpath & "schema.ini") Then
            Dim filstr As StreamWriter = New StreamWriter(localpath & "schema.ini", False) 'overwritten if the schema.ini file exists
            filstr.WriteLine("[" & processingTxt & "]")
            filstr.WriteLine("TEXTDELIMITER=none")
            filstr.WriteLine("ColNameHeader=False")
            'filstr.WriteLine("CharacterSet=UTF-8")
            filstr.WriteLine("Format=Delimited(@)")
            filstr.WriteLine("DateTimeFormat=YYMMDD")
            For i1 As Integer = 0 To (mapping1.Count - 1)
                If mapping1(i1).ToLower().IndexOf("dat") = 0 Then
                    filstr.WriteLine("Col" & (i1 + 1) & "=" & mapping1(i1) & " DateTime")
                Else
                    filstr.WriteLine("Col" & (i1 + 1) & "=" & mapping1(i1) & " Text")
                End If
            Next
            filstr.Close()
            filstr.Dispose()
        End If




        Dim dtAdapter1 As OleDbDataAdapter = New OleDbDataAdapter("SELECT " & String.Join(" , ", mapping1) & " , txt_order_key , txt_grade , txt_color " & " FROM " & tablename, conn)
        'Dim dtAdapter1 As OleDbDataAdapter = New OleDbDataAdapter("SELECT * FROM " & tablename, conn)
        Dim cmdbAccessCmdBuilder As New OleDbCommandBuilder(dtAdapter1)
        dtAdapter1.InsertCommand = cmdbAccessCmdBuilder.GetInsertCommand()
        dtAdapter1.UpdateCommand = cmdbAccessCmdBuilder.GetUpdateCommand()

        Dim dataDS1 As DataSet = New DataSet

        Try
            dtAdapter1.Fill(dataDS1, "ToTable")

            'add primary keys
            Dim keys(2) As DataColumn
            keys(0) = dataDS1.Tables("ToTable").Columns("txt_order_no")
            keys(1) = dataDS1.Tables("ToTable").Columns("txt_order_line_no")
            dataDS1.Tables("ToTable").PrimaryKey = keys


        Catch ex As Exception
            msge.Append("<br /><div style='color:red'>Data from Access " & " -- " & ex.Message & "</div>")
            continues = False
        End Try

        'transfer data from text file to dataset
        For Each file1 In files
            File.Copy(file1, localpath & processingTxt, True)
            msge.Append(dataInTextToTable(localpath, processingTxt, dataDS1))
        Next

        'fill data for columns txt_order_key,txt_grade,txt_color
        'delete some orders do not belong to Engineering Resin
        Dim resinRows1() As DataRow = dataDS1.Tables("ToTable").Select(" txt_gl_class  In ('" & valueOf("strProductCategoryExcluded").Replace(",", "','") & "')")
        For i5 As Integer = 0 To (resinRows1.Count - 1)
            resinRows1(i5).Delete()
        Next

        'split txt_item_no into txt_grade and txt_color
        Dim resinRows() As DataRow = dataDS1.Tables("ToTable").Select(Nothing, Nothing, DataViewRowState.CurrentRows)
        Dim i3 As Integer = 0
        Try
            For i3 = 0 To (resinRows.Count - 1)
                resinRows(i3)("txt_order_key") = resinRows(i3)("txt_order_no") & "-" & resinRows(i3)("txt_order_line_no")
                'Dim grade_color() As String = resinRows(i3)("txt_item_no").Split("-".ToCharArray(), StringSplitOptions.None)
                Dim i4 As Integer = resinRows(i3)("txt_item_no").ToString().IndexOf("-")
                If i4 > 0 Then
                    resinRows(i3)("txt_grade") = resinRows(i3)("txt_item_no").ToString().Substring(0, i4)
                    resinRows(i3)("txt_color") = resinRows(i3)("txt_item_no").ToString().Substring(i4 + 1)
                Else
                    resinRows(i3).Delete()
                End If

            Next

        Catch ex As Exception
            errormsg &= "<br /> error happened when item ==> grade&color " & resinRows(i3)("txt_item_no").ToString() & " -- " & ex.Message
            continues = False
            'msge.Append("<br /><div style='color:red'>error happened when item ==> grade&color " & resinRows(i3)("txt_item_no").ToString() & " -- " & ex.Message & "</div>")
        End Try


        Try
            dtAdapter1.Update(dataDS1, "ToTable")

        Catch ex As Exception
            errormsg &= "<br /> Data to Access " & " -- " & ex.Message
            continues = False
            'msge.Append("<br /><div style='color:red'>Data to Access " & " -- " & ex.Message & "</div>")
        End Try

        cmdbAccessCmdBuilder.Dispose()
        dataDS1.Dispose()

        dtAdapter1.Dispose()

        If Not (conn.State = ConnectionState.Closed) Then
            conn.Close()
        End If
        conn.Dispose()

        Dim endtime As Date = Date.Now
        Dim elapsedTicks As Long = endtime.Ticks - begintime.Ticks
        Dim elapsedSpan As New TimeSpan(elapsedTicks)
        msge.Append(" <br /> Processing EDI files:  " & elapsedSpan.Hours & " hours," & elapsedSpan.Minutes & " minutes," & elapsedSpan.Seconds & " seconds.")

        Return msge.ToString()

    End Function



    ''' <summary>
    ''' use schema.ini to transfer data from txt file to specific table in access database
    ''' </summary>
    Private Function dataInTextToTable(ByVal localpath As String, ByVal filename As String, ByRef dataDS1 As DataSet) As String

        Dim msge As System.Text.StringBuilder = New System.Text.StringBuilder()

        Dim strSql As String = "SELECT * FROM [" & filename & "]"

        If localpath.EndsWith("\") Then localpath = localpath.Substring(0, localpath.Length - 1)

        Dim strCSVConnString As String = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & localpath & ";" & "Extended Properties='text;HDR=NO;'"

        ' load the data from text to DataTable 
        Using oleda As OleDbDataAdapter = New OleDbDataAdapter(strSql, strCSVConnString)

            Try
                oleda.FillLoadOption = LoadOption.Upsert
                oleda.Fill(dataDS1, "ToTable")

            Catch ex As Exception
                msge.Append("<br /> <div style='color:red'>" & filename & " -- " & ex.Message & "</div>")
            End Try
        End Using

        Return msge.ToString()

    End Function


    ''' <summary>
    ''' import data from txt file to specific table in access database
    ''' </summary>
    Private Function dataFromTxtToTableForOrder(ByVal localfile As String, ByVal tablename As String, ByVal mapping() As String, ByRef conn As OleDbConnection, Optional ByVal separator As String = "@") As String
        dataFromTxtToTableForOrder = "true"
        'open the table
        Dim dtAdapter1 As OleDbDataAdapter = New OleDbDataAdapter("SELECT " & String.Join(" , ", mapping) & " FROM " & tablename, conn)
        Dim cmdbAccessCmdBuilder As New OleDbCommandBuilder(dtAdapter1)
        dtAdapter1.InsertCommand = cmdbAccessCmdBuilder.GetInsertCommand()

        Dim dataDS1 As DataSet = New DataSet

        Try
            dtAdapter1.Fill(dataDS1, "ToTable")

        Catch ex As Exception
            dataFromTxtToTableForOrder = tablename & " -- " & ex.Message
        End Try

        'open the txt file and read data from it
        Dim seperatorstring() As Char = separator.ToCharArray()
        Dim intCount As Integer = mapping.Count
        Dim dtTable As DataTable = dataDS1.Tables("ToTable")
        Dim sr As StreamReader

        Try
            ' Create an instance of StreamReader to read from a file.
            sr = New StreamReader(localfile)
            Dim line As String
            line = sr.ReadLine()
            While (line IsNot Nothing)
                Dim result() As String = line.Split(seperatorstring, StringSplitOptions.None)
                If result.Count = intCount Then
                    Dim dtRow() As DataRow = dtTable.Select("txt_order_no = '" & result(3) & "' And txt_order_line_no ='" & result(4) & "'")
                    If dtRow.Count = 0 Then
                        Dim newdtRow As DataRow = dtTable.NewRow()
                        For i1 As Integer = 0 To intCount - 1
                            If Not String.IsNullOrEmpty(result(i1)) Then
                                If mapping(i1).IndexOf("dat") = 0 Then
                                    newdtRow(mapping(i1)) = New Date("20" & result(i1).Substring(0, 2), result(i1).Substring(2, 2), result(i1).Substring(4, 2))
                                Else
                                    newdtRow(mapping(i1)) = result(i1)
                                End If
                            End If
                        Next
                        dtTable.Rows.Add(newdtRow)
                    End If
                End If

                line = sr.ReadLine()
            End While
            sr.Close()
            dtAdapter1.Update(dataDS1, "ToTable")
        Catch E As Exception
            ' Let the user know what went wrong.
            sr.Close()

        End Try

        dtTable.Dispose()
        dataDS1.Dispose()
        cmdbAccessCmdBuilder.Dispose()

        dtAdapter1.Dispose()


    End Function



    Protected Overrides Sub checkAuthorityForPage()
        'if not logged in, then go to login page
        If accesslevel() > 2 AndAlso (Request.Cookies("userInfo") Is Nothing OrElse (CType(Request.Cookies("userInfo")("level"), Integer) < accesslevel())) Then

            Response.Redirect("~/usermanage/login.aspx?oldurl=" & IIf(String.IsNullOrEmpty(Request.Params("login")), Request.Url.ToString, String.Empty))
        Else
            'If CacheFrom("sso") Is Nothing Then
            '    'Response.Redirect("~/usermanage/loginBySSO.aspx?oldurl=" & Request.Url.ToString)

            'End If
        End If
    End Sub



    Private Sub planningActionsAfterInterfaceFilesGot(ByRef message As String)


        Dim errMessage As String = message



        Try

            Dim status1 As StringBuilder = New StringBuilder()
            Dim begintime As Date = Date.Now, endtime As Date



            'endtime = Date.Now
            'Dim elapsedTicks As Long = endtime.Ticks - begintime.Ticks
            'Dim elapsedSpan As New TimeSpan(elapsedTicks)
            'status1.Append(" <br /> " & elapsedSpan.Days & " days," & elapsedSpan.Hours & " hours," & elapsedSpan.Minutes & " minutes," & elapsedSpan.Seconds & " seconds.<br />")

            Dim showExceptionReport As Boolean = False

            Dim connstr As String = ConfigurationManager.ConnectionStrings(dbConnectionName).ProviderName & ConfigurationManager.ConnectionStrings(dbConnectionName).ConnectionString
            Using conn As OleDbConnection = New OleDbConnection(connstr)
                conn.Open()

                NewRevisionOrderDataToMainOrderTable(conn, errMessage)


                errMessage &= calculateRSD(conn)


                errMessage &= Planned_production_qty(conn)


                assignLineToNewOrder(conn, errMessage)

                finishTime_exPlantDate_Span(conn, showExceptionReport)

                Preparation_For_Automatic_scheduling(conn, errMessage)


                MOQappliedToNewOrder(conn, errMessage)

                get_Orders_Start_Time(conn, errMessage)


                scheduleFreeSampleStartTime(conn, errMessage)


                paymentTermsCheck(conn, errMessage)


                errMessage &= finishTime_exPlantDate_Span(conn, showExceptionReport)


                errMessage &= AssignScrewDieAndFDA(conn)


                errMessage &= PackageOrMiscellaneous(conn)


                resolveLeadtimeBackToTable(conn, errMessage)

                conn.Close()
            End Using

            endtime = Date.Now
            Dim elapsedTicks As Long = endtime.Ticks - begintime.Ticks
            Dim elapsedSpan As TimeSpan = New TimeSpan(elapsedTicks)
            status1.Append(errMessage)
            status1.Append(" <br />Several frequent planning actions from NewRevisionOrderDataToMainOrderTable to  resolveLeadtimeBackToTable :" & elapsedSpan.Hours & " hours," & elapsedSpan.Minutes & " minutes," & elapsedSpan.Seconds & " seconds.<br />")
            message = status1.ToString()

        Catch ex As Exception
            message &= "<div style='color:red;'>" & ex.Message & "</div>"
        End Try

  

    End Sub




   
    Protected Sub listEDI_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles listEDI.Click
        Response.Redirect("orderFilesEDI.aspx")
    End Sub
End Class
