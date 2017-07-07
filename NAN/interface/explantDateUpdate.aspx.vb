Imports System.IO
Imports System.Text
Imports System.Data
Imports System.Data.SqlClient
Imports System.Globalization
Imports Microsoft.VisualBasic
Imports System.Linq
Imports System.Threading



Partial Class interface_explantDateUpdate
    Inherits basepage1

    Public Delegate Sub asychrSub(ByRef PGPTbatchAndOrder2 As PGPTbatchAndOrder) 'to run an asynchronous function,to delete those records which have the mark of screw pull earlier than 3 months ago

    ''' <summary>
    ''' initialize the time to the morning of last working day 
    ''' </summary>
    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        If Not Page.IsPostBack Then
            Dim startpoint As DateTime = DateTime.Today.AddYears(-1).Date

            'date time format based on  culture  of en-US
            Dim originalCulture As CultureInfo = Thread.CurrentThread.CurrentCulture
            Thread.CurrentThread.CurrentCulture = New CultureInfo("en-US")
            'Specific format for date and time
            txtStartPoint.Text = startpoint.ToShortDateString

            'date time format based on original culture
            Thread.CurrentThread.CurrentCulture = originalCulture


        End If

        message.Text = String.Empty

    End Sub


    ''' <summary>
    ''' generate explant EDI file and ftp the EDI file to OPM server
    ''' EDI file format as follows.  @orgnization code@order no@order no line@new explant date('d-mmm-yy')@RSD('d-mmm-yy')
    ''' </summary>
    Protected Sub explantToOPM_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles explantToOPM.Click


        Dim errAndNormalMessage As StringBuilder = New StringBuilder()

        'delete all the records in the table
        Dim connstr As String = ConfigurationManager.ConnectionStrings(dbConnectionName).ConnectionString
        Dim conn As SqlConnection = New SqlConnection(connstr)

        Dim continueToUpdate As Boolean = True

        Dim userName As String = lockKeyTable(priority.UploadExplantDate)
        If Not String.IsNullOrEmpty(userName) Then
            continueToUpdate = False
            errAndNormalMessage.Append("<div style='color:red;font-size:150%;'>" & userName & " is using the order detail table" & "</div>")
        End If

        Dim dtUpdateTo As SqlDataAdapter = New SqlDataAdapter("SELECT * FROM Esch_Na_tbl_explant", conn)
        Dim cmdbAccessCmdBuilder As New SqlCommandBuilder(dtUpdateTo)
        dtUpdateTo.InsertCommand = cmdbAccessCmdBuilder.GetInsertCommand()
        dtUpdateTo.DeleteCommand = cmdbAccessCmdBuilder.GetDeleteCommand()


        Dim dtTableUpdateTo As DataTable = New DataTable()
        dtUpdateTo.Fill(dtTableUpdateTo)

        Dim keys(0) As DataColumn
        keys(0) = dtTableUpdateTo.Columns("txt_order_key")
        dtTableUpdateTo.PrimaryKey = keys


        For Each a As DataRow In dtTableUpdateTo.Rows
            a.Delete()  'firstly delete all the records in the table
        Next


        Dim dummyLine As Integer = CInt(valueOf("intDummyLine"))
        Dim theSameDayOfNextYear As DateTime = DateTime.Today.AddYears(1)
        Dim tmStamp As DateTime = Today


        Dim sqlWhereClause As StringBuilder = New StringBuilder(" WHERE (txt_order_type = 'MTO') And (int_status_key <> 'invoiced') And (int_status_key <> 'cancelled') And (dat_new_explant >= " & dateSeparator & CDate(txtStartPoint.Text) & dateSeparator & ") And (CAST(int_line_no as VARCHAR(5)) <> '" & valueOf("intDummyLine") & "')")


        Dim dtFrom0 As SqlDataAdapter = New SqlDataAdapter("SELECT txt_orgn_code,int_status_key,int_line_no,txt_order_key,dat_etd,dat_new_explant FROM Esch_Na_tbl_orders " & sqlWhereClause.ToString(), conn)
        Dim cmdbAccessCmdBuilder0 As New SqlCommandBuilder(dtFrom0)
        dtFrom0.UpdateCommand = cmdbAccessCmdBuilder0.GetUpdateCommand()

        Dim dtTableFrom0 As DataTable = New DataTable()
        dtFrom0.Fill(dtTableFrom0)
        Dim keys0(0) As DataColumn
        keys0(0) = dtTableFrom0.Columns("txt_order_key")
        dtTableFrom0.PrimaryKey = keys0

        'how many records to be inserted into table
        Dim recordsCount As Integer = dtTableFrom0.Rows.Count


        'decide how to upload explant date, if dat_etd is later than actual ex-plant date, then new explant date is equal to dat_etd instead
        If Not String.IsNullOrEmpty(valueOf("bnlLateOneOfRSDandExPlant")) AndAlso CBool(valueOf("bnlLateOneOfRSDandExPlant")) Then
            For Each etdAndExPlant As DataRow In dtTableFrom0.Rows
                If CDate(etdAndExPlant.Item("dat_etd")).CompareTo(CDate(etdAndExPlant.Item("dat_new_explant"))) > 0 Then
                    etdAndExPlant.Item("dat_new_explant") = etdAndExPlant.Item("dat_etd")
                End If
            Next
        End If



        'connect to explant history table
        Dim dtFrom1 As SqlDataAdapter = New SqlDataAdapter("SELECT * FROM Esch_Na_Esch_Na_tbl_explant_history ", conn)
        Dim cmdbCmdBuilder As New SqlCommandBuilder(dtFrom1)
        dtFrom1.InsertCommand = cmdbCmdBuilder.GetInsertCommand()
        dtFrom1.UpdateCommand = cmdbCmdBuilder.GetUpdateCommand()
        dtFrom1.DeleteCommand = cmdbCmdBuilder.GetDeleteCommand()




        Dim dtTableFrom1 As DataTable = New DataTable()
        dtFrom1.Fill(dtTableFrom1)


        Dim keys1(0) As DataColumn
        keys1(0) = dtTableFrom1.Columns("txt_order_key")
        dtTableFrom1.PrimaryKey = keys1


        For Each r As DataRow In dtTableFrom1.Select("dte_timestamp <= " & dateSeparator & tmStamp.AddMonths(-2) & dateSeparator & "")
            r.Delete()
        Next




        Dim orgnization As String = valueOf("strOrganization")
        Dim pathOfEDI As String = ConfigurationManager.AppSettings("EDIfilesFolder") & "\"
        'If pathOfEDI.EndsWith("\\") Then pathOfEDI = pathOfEDI.Replace("\\", "\")
        If Not pathOfEDI.EndsWith("\") Then pathOfEDI &= "\"
        Dim fileNameOfEDI As String = valueOf("strExplantPrefix") & "." & orgnization & "." & String.Format("{0:yyyyMMddHHmmss}" & ".CSV", System.DateTime.Now)
        Dim pathAndFileNameOfEDI As String = pathOfEDI & fileNameOfEDI

        Dim ExplantdateReallyChanged As Boolean = True 'if there is a real change on ex-plant date, change order's status to 'old' after upload explant date data

        Try

            If recordsCount > 0 Then



                For Each dtrow0 As DataRow In dtTableFrom0.Rows 'get order line data from table Esch_Na_tbl_orders

                    Dim dtHistory() As DataRow = dtTableFrom1.Select("txt_order_key = '" & dtrow0.Item("txt_order_key") & "'")


                    If dtHistory.Count > 0 Then
                        Try
                            If (Not DBNull.Value.Equals(dtHistory(0).Item("dat_new_explant"))) AndAlso dtHistory(0).Item("dat_new_explant") = dtrow0.Item("dat_new_explant") AndAlso dtHistory(0).Item("dat_etd") = dtrow0.Item("dat_etd") Then
                            Else
                                ExplantdateReallyChanged = True 'deafault as if there is a change on ex-plant date
                                Dim newRow As DataRow = dtTableUpdateTo.NewRow()
                                newRow.Item("txt_orgn_code") = dtrow0.Item("txt_orgn_code")
                                newRow.Item("txt_order_key") = dtrow0.Item("txt_order_key")

                                If dtrow0.Item("int_line_no") = dummyLine Then 'explant is null if this order is arranged in dummy production line
                                    newRow.Item("dat_new_explant") = DBNull.Value
                                    'dtHistory(0).Item("dat_new_explant") = ""
                                Else
                                    If CDate(dtrow0.Item("dat_new_explant")).CompareTo(theSameDayOfNextYear) > 0 Then  'if explant date is one unreasonably late, say later than the same day of next year
                                        newRow.Item("dat_new_explant") = theSameDayOfNextYear
                                        dtHistory(0).Item("dat_new_explant") = theSameDayOfNextYear
                                    Else
                                        newRow.Item("dat_new_explant") = dtrow0.Item("dat_new_explant")

                                        If (Not DBNull.Value.Equals(dtHistory(0).Item("dat_new_explant"))) AndAlso dtHistory(0).Item("dat_new_explant") = dtrow0.Item("dat_new_explant") Then
                                            ExplantdateReallyChanged = False  'in this case , there is no real change on ex-plant date
                                        Else
                                            dtHistory(0).Item("dat_new_explant") = dtrow0.Item("dat_new_explant")
                                        End If

                                    End If
                                End If


                                newRow.Item("dat_etd") = dtrow0.Item("dat_etd")
                                dtHistory(0).Item("dat_etd") = dtrow0.Item("dat_etd")
                                dtHistory(0).Item("dte_timestamp") = tmStamp


                                dtTableUpdateTo.Rows.Add(newRow)

                                If ExplantdateReallyChanged Then
                                    Dim old_int_status_key As String = dtrow0.Item("int_status_key")
                                    If Not DBNull.Value.Equals(newRow.Item("dat_new_explant")) Then
                                        If old_int_status_key.IndexOf("2") <> -1 Then
                                            If CDate(dtrow0.Item("dat_new_explant")).CompareTo(CDate(dtrow0.Item("dat_etd"))) <= 0 Then
                                                dtrow0.Item("int_status_key") = old_int_status_key.Replace("2", "0")
                                            Else
                                                dtrow0.Item("int_status_key") = old_int_status_key.Replace("2", "1")
                                            End If

                                        Else
                                            If old_int_status_key.IndexOf("1") <> -1 Then
                                                If CDate(dtrow0.Item("dat_new_explant")).CompareTo(CDate(dtrow0.Item("dat_etd"))) <= 0 Then
                                                    dtrow0.Item("int_status_key") = old_int_status_key.Replace("1", "0")
                                                Else
                                                    dtrow0.Item("int_status_key") = "old"  'change order's status to 'old' .
                                                End If
                                            End If                                          
                                        End If
                                    End If

                                End If


                                End If

                        Catch ex As Exception
                            errAndNormalMessage.AppendLine("<div style='color:red;'>There is historical explant record for " & dtrow0.Item("txt_order_key") & " " & ex.Message & "</div>")
                        End Try
                    Else 'no record exist before,need create new record both Esch_Na_tbl_explant and Esch_Na_Esch_Na_tbl_explant_history
                        Try
                            Dim newRow As DataRow = dtTableUpdateTo.NewRow()
                            Dim newHistory As DataRow = dtTableFrom1.NewRow()

                            newRow.Item("txt_orgn_code") = dtrow0.Item("txt_orgn_code")
                            newRow.Item("txt_order_key") = dtrow0.Item("txt_order_key")

                            newHistory.Item("txt_order_key") = dtrow0.Item("txt_order_key")
                            newHistory.Item("dte_timestamp") = tmStamp

                            If dtrow0.Item("int_line_no") = dummyLine Then
                                newRow.Item("dat_new_explant") = DBNull.Value
                                newHistory.Item("dat_new_explant") = DBNull.Value
                            Else
                                newRow.Item("dat_new_explant") = dtrow0.Item("dat_new_explant")
                                newHistory.Item("dat_new_explant") = dtrow0.Item("dat_new_explant")
                            End If

                            newRow.Item("dat_etd") = dtrow0.Item("dat_etd")
                            newHistory.Item("dat_etd") = dtrow0.Item("dat_etd")

                            dtTableUpdateTo.Rows.Add(newRow)
                            dtTableFrom1.Rows.Add(newHistory)


                            Dim old_int_status_key As String = dtrow0.Item("int_status_key")
                            'If old_int_status_key.IndexOf("2") <> -1 Then
                            '    dtrow0.Item("int_status_key") = old_int_status_key.Replace("2", "1")
                            'Else
                            '    dtrow0.Item("int_status_key") = "old"  'change order's status to 'old' .
                            'End If

                            If Not DBNull.Value.Equals(newRow.Item("dat_new_explant")) Then
                                If old_int_status_key.IndexOf("2") <> -1 Then
                                    If CDate(dtrow0.Item("dat_new_explant")).CompareTo(CDate(dtrow0.Item("dat_etd"))) <= 0 Then
                                        dtrow0.Item("int_status_key") = old_int_status_key.Replace("2", "0")
                                    Else
                                        dtrow0.Item("int_status_key") = old_int_status_key.Replace("2", "1")
                                    End If

                                Else
                                    If old_int_status_key.IndexOf("1") <> -1 Then
                                        If CDate(dtrow0.Item("dat_new_explant")).CompareTo(CDate(dtrow0.Item("dat_etd"))) <= 0 Then
                                            dtrow0.Item("int_status_key") = old_int_status_key.Replace("1", "0")
                                        Else
                                            dtrow0.Item("int_status_key") = "old"  'change order's status to 'old' .
                                        End If
                                    End If
                                End If
                            End If

                        Catch ex As Exception
                            errAndNormalMessage.AppendLine("<div style='color:red;'>There is no historical explant record for " & dtrow0.Item("txt_order_key") & " " & ex.Message & "</div>")
                        End Try

                    End If



                Next




                'extract all the data from Esch_Na_tbl_explant to convert to a string in order to write it to a file

                Dim orgnizationInEDIfile As String = String.Empty
                Dim orderNoInEDIfile As String = String.Empty
                Dim orderNoLineInEDIfile As String = String.Empty
                Dim explantInEDIfile As String = String.Empty
                Dim etdInEDIfile As String = String.Empty

                Dim listOfExplantLines As StringBuilder = New StringBuilder

                Dim originalCulture As CultureInfo = Thread.CurrentThread.CurrentCulture
                Thread.CurrentThread.CurrentCulture = New CultureInfo("en-US")

                For Each r As DataRow In dtTableUpdateTo.Select(Nothing)
                    'use to write to EDI file
                    orgnizationInEDIfile = r.Item("txt_orgn_code")
                    orderNoInEDIfile = r.Item("txt_order_key").ToString.Split("-".ToCharArray())(0)
                    orderNoLineInEDIfile = r.Item("txt_order_key").ToString.Split("-".ToCharArray())(1)

                    If Not DBNull.Value.Equals(r.Item("dat_new_explant")) Then
                        explantInEDIfile = String.Format("{0:d-MMM-yy}", CDate(r.Item("dat_new_explant")))
                    Else
                        explantInEDIfile = String.Empty
                    End If

                    If Not DBNull.Value.Equals(r.Item("dat_etd")) Then
                        etdInEDIfile = String.Format("{0:d-MMM-yy}", CDate(r.Item("dat_etd")))
                    Else
                        etdInEDIfile = String.Empty
                    End If

                    listOfExplantLines.Append(orgnizationInEDIfile & "," & orderNoInEDIfile & "," & orderNoLineInEDIfile & "," & explantInEDIfile & "," & etdInEDIfile & vbLf)

                Next



                'write all the data of explant to a file
                If listOfExplantLines.Length > 0 Then

                    Using outFile As New StreamWriter(pathAndFileNameOfEDI)
                        outFile.Write(listOfExplantLines.ToString)
                    End Using

                Else
                    continueToUpdate = False 'no record
                End If


                errAndNormalMessage.AppendLine("The number of order lines which have update explant date is " & dtTableUpdateTo.Select(Nothing).Count)

                'date time format based on original culture
                Thread.CurrentThread.CurrentCulture = originalCulture


            Else
                continueToUpdate = False 'no record
                errAndNormalMessage.AppendLine("The number of order lines which have update explant date is 0")
            End If

        Catch ex As Exception
            continueToUpdate = False
            errAndNormalMessage.AppendLine("<div style='color:red;'> " & ex.Message & "</div>")
        End Try

        'write data back to table Esch_Na_tbl_explant
        If continueToUpdate Then
            dtUpdateTo.Update(dtTableUpdateTo)

        Else
            If File.Exists(pathAndFileNameOfEDI) Then File.Delete(pathAndFileNameOfEDI)

        End If


        'ftp EDI file to OPM server
        Dim localpath As String = ConfigurationManager.AppSettings("EDIfilesFolder") & "\"
        If Not localpath.EndsWith("\") Then localpath &= "\"
        localpath &= "interfaceData\"

        If continueToUpdate Then

            If Not Directory.Exists(localpath) Then
                Try
                    Directory.CreateDirectory(localpath)
                Catch ex As Exception
                    'msgPopUP("You might not have right to create folder: " & localpath, Literal1)
                    continueToUpdate = False
                    errAndNormalMessage.AppendLine("<div style='color:red;'> You might not have right to create folder: " & ex.Message & "</div>")
                End Try
            End If
        End If

        If continueToUpdate Then
            'send file to OPM ftp server via ftp
            Dim localpathsubfolder As String = localpath & "ExplantDate\"
            Dim workingDirectoryAtServer As String = valueOf("strExplantPath")

            'get user name and password and ftp server IP to log on to server to get files by ftp method
            Dim organization As String = valueOf("strOrganization")
            Dim serverUri As String = valueOf("strFTPserverIP")
            If Not serverUri.StartsWith("ftp://") Then serverUri = "ftp://" & serverUri
            If Not serverUri.EndsWith("/") Then serverUri &= "/"

            If Not workingDirectoryAtServer.EndsWith("/") Then workingDirectoryAtServer &= "/"

            Dim debugPrefix As String = String.Empty
            If serverUri.IndexOf("127.0.0.1") > 0 Then
                debugPrefix = valueOf("strCompanyDomain") & "\"
            End If

            Dim ftpoperation As FTPcls = New FTPcls(debugPrefix & valueOf("strFTP_ID"), valueOf("strFTP_PW"), serverUri & workingDirectoryAtServer)

            Dim returnMessageAfterFTPoperation As String = ftpoperation.UploadFileToServer(pathAndFileNameOfEDI, fileNameOfEDI).ToLower()

            If returnMessageAfterFTPoperation = "true" Then
                'backup the EDI file by moving the file to one sub folder
                If File.Exists(pathAndFileNameOfEDI) Then
                    File.Copy(pathAndFileNameOfEDI, localpathsubfolder & fileNameOfEDI, True)
                    File.Delete(pathAndFileNameOfEDI)
                End If

                errAndNormalMessage.AppendLine("<div style='color:black;'> " & pathAndFileNameOfEDI & " has been sucessfully ftp to server " & serverUri & workingDirectoryAtServer & "</div>")

            Else
                continueToUpdate = False
                errAndNormalMessage.AppendLine("<div style='color:red;'> " & returnMessageAfterFTPoperation & "</div>")
            End If

        End If




        'make this update as the part of table explant history
        If continueToUpdate Then

            Try

                dtFrom1.Update(dtTableFrom1) 'write data back to Esch_Na_Esch_Na_tbl_explant_history
            Catch ex As Exception
                continueToUpdate = False
                errAndNormalMessage.AppendLine("<div style='color:red;'> update explant historical : " & ex.Message & "</div>")
            End Try

        End If

        'write the changes on order status back to Esch_Na_tbl_orders
        If continueToUpdate Then

            Try
                For Each etdAndExPlant As DataRow In dtTableFrom0.Rows
                    etdAndExPlant.Item("dat_new_explant") = etdAndExPlant.Item("dat_new_explant", DataRowVersion.Original)
                Next
                dtFrom0.Update(dtTableFrom0) 'write data back to Esch_Na_Esch_Na_tbl_explant_history
            Catch ex As Exception
                errAndNormalMessage.AppendLine("<div style='color:red;'> update order status back to order details table : " & ex.Message & "</div>")
            End Try

        End If




        dtTableFrom0.Dispose()
        cmdbAccessCmdBuilder0.Dispose()

        dtTableFrom1.Dispose()
        cmdbCmdBuilder.Dispose()

        dtTableUpdateTo.Dispose()
        cmdbAccessCmdBuilder.Dispose()

        dtFrom0.Dispose()
        dtFrom1.Dispose()

        dtUpdateTo.Dispose()


        If String.IsNullOrEmpty(userName) Then
            unlockKeyTable(priority.UploadExplantDate)
        End If


        If errAndNormalMessage.Length > 0 Then
            msgPopUP(errAndNormalMessage.ToString, message, False, False)
        End If


        'trigger an asychronous action to do some routine tasks
        asyAction()


    End Sub


    'trigger a asynchronous action to do some routine works
    Public Sub asyAction()

        Dim PGPTbatchAndOrder2 As New PGPTbatchAndOrder()
        PGPTbatchAndOrder2.organization = valueOf("strInvalidOrgn")
        PGPTbatchAndOrder2.strFTP_ID = valueOf("strFTP_ID")
        PGPTbatchAndOrder2.strFTP_PW = valueOf("strFTP_PW")
        PGPTbatchAndOrder2.strBatchStatusUpdatePrefix = valueOf("strBatchStatusUpdatePrefix")
        PGPTbatchAndOrder2.strBatchStatusUpdatePath = valueOf("strBatchStatusUpdatePath")
        PGPTbatchAndOrder2.strFTPserverIP = valueOf("strFTPserverIP")
        PGPTbatchAndOrder2.strOpenOrderPrefix = valueOf("strOpenOrderPrefix")
        PGPTbatchAndOrder2.strOpenOrderPath = valueOf("strOpenOrderPath")

        Dim asySubroutine As asychrSub
        asySubroutine = New asychrSub(AddressOf doSomething)
        asySubroutine.BeginInvoke(PGPTbatchAndOrder2, Nothing, Nothing)
    End Sub


    Public Sub doSomething(ByRef PGPTbatchAndOrder2 As PGPTbatchAndOrder)

        Dim doSomethings As New routineMaintenance(Cache, PGPTbatchAndOrder2)
        doSomethings.routineMaintenance()

    End Sub


    Protected Sub listEDI_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles listEDI.Click
        Response.Redirect("explantFilesEDI.aspx")
    End Sub
End Class

