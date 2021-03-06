﻿Imports System.IO
Imports System.Data.SqlClient
Imports System.Data
Imports System.Globalization
Imports System.Threading


Partial Class interface_downloadBatchstatus
    Inherits FrequentPlanActions


    'import batch status update information from OPM ftp server to local folder and consider to share the same files with S$FS department
    'read the data in the txt file and use the data to update relevant data table in access database
    Protected Sub batchstatus_FromOPMftpServer(ByRef ftpMsg As String)

        Dim errMessage As StringBuilder = New StringBuilder()
        Dim msge As System.Text.StringBuilder = New System.Text.StringBuilder()
        Dim continues As Boolean = True

        Dim userName As String = lockKeyTable(priority.ImportNewOrderOrBatch)
        If Not String.IsNullOrEmpty(userName) Then
            continues = False
            errMessage.Append("<div style='color:red;font-size:150%;'>" & userName & " is using the order detail table" & "</div>")
        End If

        Try



            Dim localpath As String = ConfigurationManager.AppSettings("EDIfilesFolder")
            If Not localpath.EndsWith("\") Then localpath &= "\"
            localpath &= "interfaceData\"
            If Not Directory.Exists(localpath) Then Directory.CreateDirectory(localpath)

            'get batch update status files from ftp server and also copy them to S&FS SBU folder for their use
            Dim localpathsubfolder As String = localpath & "BatchStatus\"
            Dim workingDirectory As String = valueOf("strBatchStatusUpdatePath")
            Dim sfsFolder As String = valueOf("strBatchstatusResinToSFS")
            Dim resinFolder As String = valueOf("strBatchstatusResinFromSFSdwnld")

            'get user name and password and ftp server IP to log on to server to get files by ftp method
            Dim organization As String = valueOf("strOrganization")
            Dim prefix As String = valueOf("strBatchStatusUpdatePrefix")
            Dim serverUri As String = valueOf("strFTPserverIP")

            If Not serverUri.StartsWith("ftp://") Then serverUri = "ftp://" & serverUri
            If Not serverUri.EndsWith("/") Then serverUri &= "/"

            If Not workingDirectory.EndsWith("/") Then workingDirectory &= "/"

            Dim ftpoperation As FTPcls = New FTPcls(valueOf("strFTP_ID"), valueOf("strFTP_PW"), serverUri & workingDirectory)

            Dim errMsg As String = String.Empty
            Dim olist As List(Of String) = ftpoperation.FileListInDirContains(prefix & "." & organization, errMsg)

            If Not String.IsNullOrEmpty(errMsg) Then ' if no error message return, then continue further processing
                'continues = False  'even let it continue would not have impact on the later process
                errMessage.Append("<br />" & errMsg)
            End If

            olist.Sort(StringComparer.Ordinal)


            If olist.Count > 1 Then
                msge.Append("<br /> " & olist.Count & " files are downloaded from ftp server " & serverUri & workingDirectory & " .<br />")
            Else
                msge.Append("<br /> " & olist.Count & " file is downloaded from ftp server " & serverUri & workingDirectory & " .<br />")
            End If

            'down the file and delete the file on the server
            For Each filename As String In olist
                If ftpoperation.DownFileFrmServer(localpath & filename, filename).ToLower() = "true" Then
                    ftpoperation.DeleteFileOnServer(filename)
                End If

                msge.Append(filename & "<br />")
            Next




            'process these files further===========
            If continues Then
                Dim err1 As String = String.Empty
                batchStatusFromOPMtoTable(err1)
                errMessage.Append(err1)
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
                    errMessage.Append("<br />" & ex.Message)
                End Try
            End If



            'pass the information of batch status from table Esch_Na_tbl_BatchSts_from_OPM  to table Esch_Na_tbl_orders
            If continues Then
                Dim err2 As String = String.Empty
                msge.Append(batchStatusFromTableTo_Esch_Na_tbl_orders(err2))
                errMessage.Append("<br />" & err2)
            End If

        Catch ex As Exception
            errMessage.Append("<br />" & ex.Message)
        End Try

        If String.IsNullOrEmpty(userName) Then
            unlockKeyTable(priority.ImportNewOrderOrBatch)
        End If

        ftpMsg = "Batch status from OPM ftp server:<br /><div style='color:black;'> " & msge.ToString() & "</div>" & "<div style='color:red;'> " & errMessage.ToString() & "</div>"

    End Sub



    ''' <summary>
    ''' after download text files about batch status, extracting data from these files and import them into access table Esch_Na_tbl_BatchSts_from_OPM 
    ''' </summary>
    ''' <remarks></remarks>
    Private Sub batchStatusFromOPMtoTable(ByRef errorM As String)
        Dim localpath As String = ConfigurationManager.AppSettings("EDIfilesFolder")
        If Not localpath.EndsWith("\") Then localpath &= "\"
        localpath &= "interfaceData\"
        If Not Directory.Exists(localpath) Then Directory.CreateDirectory(localpath)

        Dim files = From file In Directory.EnumerateFiles(localpath) Where file.ToLower().Contains(valueOf("strBatchStatusUpdatePrefix").ToLower()) Order By file Descending
        'Dim mapping1() As String = {"txt_lot_no", "txt_batch_status", "txt_actual_line_no", "dat_actual_start", "dat_actual_finish", "flt_actual_qty"}


        Dim connstr As String = ConfigurationManager.ConnectionStrings(dbConnectionName).ConnectionString
        Dim conn As SqlConnection = New SqlConnection(connstr)

        conn.Open()

        Dim command As New SqlCommand("SELECT txtFieldName FROM Esch_Na_tbl_interface_mapping WHERE intImpBatchstatusMapping > 0 ORDER BY intImpBatchstatusMapping", conn)
        Dim reader As SqlDataReader = command.ExecuteReader()
        Dim mapping1() As String = {String.Empty}, icount As Integer = 0
        Dim btchstsTable As String = "Esch_Na_tbl_BatchSts_from_OPM"
        While reader.Read()
            ReDim Preserve mapping1(icount)
            mapping1(icount) = reader("txtFieldName")
            icount += 1
        End While
        reader.Close()

        'empty the table to accept new data
        command.CommandText = "DELETE  FROM " & btchstsTable
        command.ExecuteNonQuery()
        command.Dispose()

        'transfer batch status data to table  from text files one by one
        For Each file1 In files
            errorM &= dataFromTxtToTable(file1, btchstsTable, mapping1, conn)
        Next

        conn.Close()

    End Sub


    'import data from txt file to specific table in access database
    Private Function dataFromTxtToTable(ByVal localfile As String, ByVal tablename As String, ByVal mapping() As String, ByRef conn As SqlConnection, Optional ByVal separator As String = "@") As String
        'dataFromTxtToTable = String.Empty

        Dim msg As StringBuilder = New StringBuilder()

        'open the table
        Dim dtAdapter1 As SqlDataAdapter = New SqlDataAdapter("SELECT " & String.Join(" , ", mapping) & " FROM " & tablename, conn)
        Dim cmdbAccessCmdBuilder As New SqlCommandBuilder(dtAdapter1)
        dtAdapter1.InsertCommand = cmdbAccessCmdBuilder.GetInsertCommand()


        Dim dataDS1 As DataSet = New DataSet

        Try
            dtAdapter1.Fill(dataDS1, "ToTable")
            Dim keys(0) As DataColumn
            keys(0) = dataDS1.Tables("ToTable").Columns("txt_lot_no")
            dataDS1.Tables("ToTable").PrimaryKey = keys

        Catch ex As Exception
            msg.Append("<br /> " & tablename & " -- " & ex.Message)
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
                    Dim dtRow() As DataRow = dtTable.Select("txt_lot_no = '" & result(0) & "'")  ' if batch status exists in the latest file, do not need add it. The trick is that we import batch status files from the latest to the oldest
                    If dtRow.Count = 0 Then
                        Dim newdtRow As DataRow = dtTable.NewRow()
                        For i1 As Integer = 0 To intCount - 1
                            If Not String.IsNullOrEmpty(result(i1)) Then
                                If mapping(i1).IndexOf("dat") = 0 Then
                                    newdtRow(mapping(i1)) = New Date(result(i1).Substring(0, 4), result(i1).Substring(4, 2), result(i1).Substring(6, 2), result(i1).Substring(8, 2), result(i1).Substring(10, 2), result(i1).Substring(12, 2))
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
            ' Let the user know what goes wrong.
            msg.Append("<br /> " & E.Message)
            sr.Close()
        Finally

        End Try

        Return msg.ToString

    End Function


    ''' <summary>
    ''' transfer batch status from table to table Esch_Na_tbl_BatchSts_from_OPM  to table Esch_Na_tbl_orders
    ''' </summary>
    ''' <param name="errorMsg"></param>
    ''' <returns></returns>
    Public Function batchStatusFromTableTo_Esch_Na_tbl_orders(ByRef errorMsg As String) As String
        Dim msgNormal As StringBuilder = New StringBuilder()

        'get the meaning for each number of batch status
        Dim connstr1 As String = ConfigurationManager.ConnectionStrings(dbConnForParam).ConnectionString
        Dim connParam As SqlConnection = New SqlConnection(connstr1)

        Dim batchMeaning As New Dictionary(Of String, String)
        Dim command As SqlCommand = New SqlCommand("SELECT [txt_batch_status],[txt_meaning]  FROM  [Esch_Na_tbl_batch_status_meaning] ", connParam)
        Dim reader As SqlDataReader

        Try
            connParam.Open()
            reader = command.ExecuteReader()

            While reader.Read()
                batchMeaning.Add(reader("txt_batch_status"), reader("txt_meaning"))
            End While

        Finally
            reader.Close()
            command.Dispose()
            connParam.Dispose()
        End Try


        Dim connstr As String = ConfigurationManager.ConnectionStrings(dbConnectionName).ConnectionString
        Dim conn As SqlConnection = New SqlConnection(connstr)

        Dim dtUpdateFrom As SqlDataAdapter = New SqlDataAdapter("Select * From Esch_Na_tbl_BatchSts_from_OPM ", conn)

        Dim dtUpdateTo As SqlDataAdapter = New SqlDataAdapter("SELECT txt_order_key,txt_lot_no,txt_batch_status,txt_actual_line_no,dat_actual_start,dat_actual_finish,flt_actual_qty FROM  Esch_Na_tbl_orders  WHERE (txt_lot_no Is Not Null) And ( int_status_key <> 'cancelled' )   ORDER BY  txt_lot_no ASC ", conn)
        Dim cmdbAccessCmdBuilder As New SqlCommandBuilder(dtUpdateTo)
        dtUpdateTo.UpdateCommand = cmdbAccessCmdBuilder.GetUpdateCommand()
        dtUpdateTo.InsertCommand = cmdbAccessCmdBuilder.GetInsertCommand()
        Dim dsAccess As DataSet = New DataSet

        Try
            dtUpdateTo.Fill(dsAccess, "UpdateTo")
            Dim keys(0) As DataColumn
            keys(0) = dsAccess.Tables("UpdateTo").Columns("txt_order_key")
            dsAccess.Tables("UpdateTo").PrimaryKey = keys

            dtUpdateFrom.Fill(dsAccess, "batchStausFromOPM")

        Catch ex As Exception
            errorMsg &= ex.Message
        Finally


        End Try


        For Each b As DataRow In dsAccess.Tables("batchStausFromOPM").Rows

            Dim orderLines() As DataRow = dsAccess.Tables("UpdateTo").Select("txt_lot_no = '" & b.Item("txt_lot_no") & "'")
            For Each oL As DataRow In orderLines
                oL.Item("txt_batch_status") = batchMeaning(b.Item("txt_batch_status"))
                oL.Item("txt_actual_line_no") = b.Item("txt_actual_line_no")
                oL.Item("dat_actual_start") = b.Item("dat_actual_start")
                oL.Item("dat_actual_finish") = b.Item("dat_actual_finish")
                oL.Item("flt_actual_qty") = b.Item("flt_actual_qty")
            Next

        Next

        msgNormal.Append("The number of updated Lots is " & dsAccess.Tables("batchStausFromOPM").Rows.Count)


        Try
            dtUpdateTo.Update(dsAccess, "UpdateTo")  'resolve changes back to database
        Catch ex As Exception
            errorMsg &= ex.Message
        End Try



        dtUpdateTo.Dispose()
        dtUpdateFrom.Dispose()
        dsAccess.Dispose()
        cmdbAccessCmdBuilder.Dispose()
        conn.Dispose()


        Return msgNormal.ToString

    End Function



    ''' <summary>
    ''' use to get batch status from QA e-Color system, QA online guys input real time Lot status. The data from QA is more reliable and more real time
    ''' </summary>
    ''' <remarks></remarks>
    Protected Sub GetBatchStatusFromQA_eColor_SQLserver(ByVal startPoint As DateTime, ByRef errMsg As String, ByRef nrmlMsg As String)



        Try

            Dim now1 As DateTime = DateTime.Now()

            Dim connstrSQL As String = ConfigurationManager.ConnectionStrings("eColor").ProviderName & ConfigurationManager.ConnectionStrings("eColor").ConnectionString
            Dim connSQL As SqlConnection = New SqlConnection(connstrSQL)
            'date format is a little different when using SQL clause to connect to SQL server database
            'advance start time by 5 days to get enough batch information from SQL server in case that some batches were running through 5 days
            Dim dtUpdateFrom As SqlDataAdapter = New SqlDataAdapter("Select On_Line_Data1.Lot_No AS actualLotNo, min([Time]) AS earliestStart,min([Line]) AS actualLine From ((SELECT Lot_No,Time FROM On_Line_Data WHERE (Lot_No Is Not Null)  And Time >'" & startPoint.AddDays(-14) & "') AS On_Line_Data1 Left Join Col_Cor_Hdr On Col_Cor_Hdr.Lot_No = On_Line_Data1.Lot_No ) WHERE ([Line] Is Not Null) Group by On_Line_Data1.Lot_No", connSQL)

            Dim db As String = ConfigurationManager.ConnectionStrings(dbConnectionName).ConnectionString
            Dim connstr As String = db
            Dim conn As SqlConnection = New SqlConnection(connstr)
            Dim dtUpdateTo As SqlDataAdapter = New SqlDataAdapter("SELECT txt_order_key,int_line_no,txt_lot_no,dat_start_date,dat_finish_date,planned_production_qty,dat_start_from_qa,flt_actual_completed,flt_actual_qty_man FROM Esch_Na_tbl_orders WHERE  ( txt_lot_no Is Not Null )  And ( int_status_key <> 'cancelled' ) And ( (dat_start_date between " & dateSeparator & startPoint & dateSeparator & " And " & dateSeparator & now1.AddDays(3) & dateSeparator & ")  Or (dat_finish_date  between " & dateSeparator & startPoint & dateSeparator & " And " & dateSeparator & now1.AddDays(3) & dateSeparator & ") )  ORDER BY txt_lot_no ASC,dat_start_date ASC", conn)
            Dim cmdbAccessCmdBuilder As New SqlCommandBuilder(dtUpdateTo)
            dtUpdateTo.UpdateCommand = cmdbAccessCmdBuilder.GetUpdateCommand()
            dtUpdateTo.InsertCommand = cmdbAccessCmdBuilder.GetInsertCommand()
            Dim dsAccess As DataSet = New DataSet
            Dim continues As Boolean = True



            dtUpdateTo.Fill(dsAccess, "UpdateTo")
            Dim keys(0) As DataColumn
            keys(0) = dsAccess.Tables("UpdateTo").Columns("txt_order_key")
            dsAccess.Tables("UpdateTo").PrimaryKey = keys

            dtUpdateFrom.Fill(dsAccess, "eColor")



            Dim howManyLots As Integer = 0
            Dim howManyOrderLines As Integer = 0

            Dim currentLotNo As String = String.Empty
            Dim currentLine As String = valueOf("intDummyLine")
            Dim howManyMinutesDelayedOrAdvanced As Long = 0

            Dim eColorRows() As DataRow

            Try
                For Each a As DataRow In dsAccess.Tables("UpdateTo").Rows

                    If dsAccess.Tables("eColor").Columns("actualLotNo").DataType = GetType(String) Then 'to judge whether actualLotNo is a string or not
                        eColorRows = dsAccess.Tables("eColor").Select("actualLotNo = '" & a.Item("txt_lot_no") & "'")
                    Else
                        eColorRows = dsAccess.Tables("eColor").Select("actualLotNo = " & a.Item("txt_lot_no"))
                    End If

                    If eColorRows.Count > 0 Then
                        If currentLotNo <> CStr(a.Item("txt_lot_no")) Then

                            currentLotNo = CStr(a.Item("txt_lot_no"))
                            currentLine = CStr(eColorRows(0).Item("actualLine"))

                            howManyLots += 1

                            If IsDate(eColorRows(0).Item("earliestStart")) Then
                                'howManyMinutesDelayedOrAdvanced = DateDiff(DateInterval.Minute, CDate(eColorRows(0).Item("earliestStart")), CDate(a.Item("dat_start_date")))
                                howManyMinutesDelayedOrAdvanced = DateDiff(DateInterval.Minute, CDate(a.Item("dat_start_date")), CDate(eColorRows(0).Item("earliestStart")))
                            Else
                                'howManyMinutesDelayedOrAdvanced = DateDiff(DateInterval.Minute, now1, CDate(a.Item("dat_start_date")))
                                howManyMinutesDelayedOrAdvanced = DateDiff(DateInterval.Minute, CDate(a.Item("dat_start_date")), now1)
                                If howManyMinutesDelayedOrAdvanced < 0 Then howManyMinutesDelayedOrAdvanced = 0 'if order's start time later than now, keep start time unchanged
                            End If

                        End If

                        If IsDate(eColorRows(0).Item("earliestStart")) Then a.Item("dat_start_from_qa") = CDate(eColorRows(0).Item("earliestStart"))

                        a.Item("dat_start_date") = CDate(a.Item("dat_start_date")).AddMinutes(howManyMinutesDelayedOrAdvanced)
                        a.Item("dat_finish_date") = CDate(a.Item("dat_finish_date")).AddMinutes(howManyMinutesDelayedOrAdvanced)
                        a.Item("int_line_no") = CInt(currentLine)

                        If DateDiff(DateInterval.Minute, CDate(a.Item("dat_start_date")), now1) > 0 Then
                            If DateDiff(DateInterval.Minute, CDate(a.Item("dat_finish_date")), now1) > 0 Then   'dat_start_date  dat_finish_date now1
                                a.Item("flt_actual_completed") = 100
                            Else   'dat_start_date  now1 dat_finish_date 
                                a.Item("flt_actual_completed") = CInt(100 * DateDiff(DateInterval.Minute, CDate(a.Item("dat_start_date")), now1) / DateDiff(DateInterval.Minute, CDate(a.Item("dat_start_date")), CDate(a.Item("dat_finish_date"))))
                            End If
                        Else   'now1  dat_start_date  dat_finish_date 
                            a.Item("flt_actual_completed") = 0
                        End If

                        If a.Item("planned_production_qty") = 0 Then
                            a.Item("flt_actual_completed") = 100
                        End If

                        a.Item("flt_actual_qty_man") = a.Item("planned_production_qty") / 100 * a.Item("flt_actual_completed")

                    Else

                        howManyMinutesDelayedOrAdvanced = DateDiff(DateInterval.Minute, CDate(a.Item("dat_start_date")), now1)
                        If howManyMinutesDelayedOrAdvanced < 0 Then howManyMinutesDelayedOrAdvanced = 0 'if order's start time later than now, keep start time unchanged
                        a.Item("dat_start_date") = CDate(a.Item("dat_start_date")).AddMinutes(howManyMinutesDelayedOrAdvanced)
                        a.Item("dat_finish_date") = CDate(a.Item("dat_finish_date")).AddMinutes(howManyMinutesDelayedOrAdvanced)

                    End If

                    howManyOrderLines += 1

                Next

            Catch ex1 As Exception
                errMsg &= ex1.Message
            End Try

            Try
                dtUpdateTo.Update(dsAccess, "UpdateTo")  'resolve changes back to database

                nrmlMsg = "The number of updated lots is " & howManyLots & " <br /> The number of updated order lines is " & howManyOrderLines

            Catch ex2 As Exception
                errMsg &= ex2.Message

            End Try

            dsAccess.Dispose()
            dtUpdateTo.Dispose()
            dtUpdateFrom.Dispose()



        Catch ex3 As Exception
            errMsg &= ex3.Message
        End Try

        If errMsg.Length > 0 Then errMsg = "Error happened when downloading data from QA eColor server: " & errMsg



    End Sub



    ''' <summary>
    ''' initialize the time to the morning of last working day 
    ''' </summary>
    ''' 
    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        If Not Page.IsPostBack Then
            Dim startpoint As DateTime = DateTime.Today.AddDays(-1)
            Select Case startpoint.DayOfWeek
                Case DayOfWeek.Sunday
                    startpoint = startpoint.AddDays(-2)
                Case DayOfWeek.Saturday
                    startpoint = startpoint.AddDays(-1)
            End Select

            'date time format based on  culture  of en-US
            Dim originalCulture As CultureInfo = Thread.CurrentThread.CurrentCulture
            Thread.CurrentThread.CurrentCulture = New CultureInfo("en-US")
            'Specific format for date and time
            txtStartPoint.Text = startpoint.ToShortDateString

            'date time format based on original culture
            Thread.CurrentThread.CurrentCulture = originalCulture

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
            ddlHour1.SelectedIndex = 0


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
            ddlMinute1.SelectedIndex = 0


        End If

        If Not valueOf("strOrganization").ToUpper.StartsWith("PGNA") Then
            UpdatePanel1.Visible = False
            fromQAandFTP.Text = "Batch status from OPM ftp server"
        End If



    End Sub

    Protected Sub fromQAandFTP_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles fromQAandFTP.Click

        If True Then


            Try
                Dim startPoint As DateTime = CDate(txtStartPoint.Text).AddHours(ddlHour1.SelectedIndex).AddMinutes(ddlMinute1.SelectedIndex)
                Dim errMsg As String = String.Empty, nrmlMsg As String = String.Empty, ftpMsg As String = String.Empty

                If valueOf("strOrganization").ToUpper.StartsWith("PGNA") Then  'enable the function for Nansha plant only
                    GetBatchStatusFromQA_eColor_SQLserver(startPoint, errMsg, nrmlMsg)
                    Label11.Text = "Batch status data from QA eColor database:<br /><div style='color:black;'>" & nrmlMsg & "</div><br /><div style='color:red;'>" & errMsg & "</div><br />"
                End If


                batchstatus_FromOPMftpServer(ftpMsg)

                Label11.Text &= ftpMsg


            Catch ex As Exception
                Label11.Text &= "<br /><div style='color:red;'>" & ex.Message & "</div>"

            End Try



        End If

    End Sub


    Protected Overrides Sub checkAuthorityForPage()
        Return   '
        'if not logged in, then go to login page
        If accesslevel() > 2 AndAlso (Request.Cookies("userInfo") Is Nothing OrElse (CType(Request.Cookies("userInfo")("level"), Integer) < accesslevel())) Then

            Response.Redirect("~/usermanage/login.aspx?oldurl=" & IIf(String.IsNullOrEmpty(Request.Params("login")), Request.Url.ToString, String.Empty))
        Else
            If CacheFrom("sso") Is Nothing Then
                Response.Redirect("~/usermanage/loginBySSO.aspx?oldurl=" & Request.Url.ToString)

            End If
        End If
    End Sub

    Protected Sub listEDI_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles listEDI.Click
        Response.Redirect("batchStatusFilesEDI.aspx")
    End Sub
End Class
