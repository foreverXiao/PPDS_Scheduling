Imports Microsoft.VisualBasic
Imports System.Data
Imports System.IO
Imports System.Text
Imports System.Globalization
Imports System.Linq
Imports System.Runtime.InteropServices
Imports System.Diagnostics
Imports System.Data.Odbc.OdbcConnection
Imports System.Data.OleDb
Imports System.Data.SqlClient


Partial Public Class InteracWithExcel
    Inherits basepage1


    Private _maxNumberOfColumn As String = "U"
    Private _maxNumberOfRow As Integer = 5001


    ''' <summary>
    ''' how many columns of data to be used when importing from or exporting to excel file
    ''' </summary>
    Public Property rightmostColumn() As String
        Get
            Return _maxNumberOfColumn
        End Get
        Set(ByVal value As String)
            _maxNumberOfColumn = value
        End Set
    End Property

    ''' <summary>
    ''' how many rows of data to be used when importing from or exporting to excel file
    ''' </summary>
    Public Property maxRowNumber() As Integer
        Get
            Return _maxNumberOfRow
        End Get
        Set(ByVal value As Integer)
            _maxNumberOfRow = value
        End Set
    End Property

    ''' <summary>
    ''' based on the integer number, come up with a corresponding column name in the excel sheet
    ''' </summary>
    ''' <param name="Col">input a integer to show how many columns in the database table</param>
    Public Function integerToExcelColumnName(ByVal Col As Integer) As String
        ' Col is the present column, not the number of cols
        Const A As Integer = 65    'ASCII value for capital A
        Dim sCol As String
        Dim iRemain As Integer
        ' THIS ALGORITHM ONLY WORKS UP TO ZZ. It fails on AAA
        If Col > 701 Then
            Return "ZZ"
        End If
        If Col <= 25 Then
            sCol = Chr(A + Col)
        Else
            iRemain = Int((Col / 26)) - 1
            sCol = Chr(A + iRemain) & integerToExcelColumnName(Col Mod 26)
        End If

        Return sCol

    End Function


    ''' <summary>
    ''' get data from sqldatasource and a excel is created to store the data, and generated excel is expected for downloading
    ''' </summary>
    ''' <param name="SDS1">Data is to be extracted from this dataSource and put into excel file</param>
    ''' <returns>return a excel file name</returns>
    Public Overloads Function generateExcel(ByRef SDS1 As SqlDataSource, ByVal filePath As String, ByRef rtrnMessage As String) As String

        Dim args As DataSourceSelectArguments = New DataSourceSelectArguments()
        Dim view1 As DataView = CType(SDS1.Select(args), DataView)
        Dim table1 As DataTable = view1.ToTable()


        Return generateExcel(rtrnMessage, table1, filePath)


    End Function


    ''' <summary>
    ''' get data from dataTable and a excel is created to store the data, and generated excel is expected for downloading
    ''' </summary>
    ''' <param name="msgRtrn">to show the warning message or others</param>
    ''' <returns>return a excel file name</returns>
    Public Overloads Function generateExcel(ByRef msgRtrn As String, ByRef dtTable As DataTable, ByVal filePath As String) As String

        If filePath.EndsWith("\") Then
            filePath = filePath.Remove(filePath.Length - 1, 1)
        End If

        Dim excelfilename As String = pageTitleReference() & userIden() & "DWLD"
        Dim errorMessage As StringBuilder = New StringBuilder()


        Dim xlApp As New Object
        Dim xlBook As New Object
        Dim startTimeBeforexlApp As DateTime
        Dim startTimeAfterxlApp As DateTime

        Try
            startTimeBeforexlApp = DateTime.Now()
            xlApp = Server.CreateObject("Excel.Application")
            startTimeAfterxlApp = DateTime.Now()

            xlApp.visible = False 'make excel application open not visible
            xlApp.DisplayAlerts = False

            xlBook = xlApp.Workbooks.Add()

        Catch ex1 As Exception
            errorMessage.Append("R1" & ex1.Message)
            msgRtrn &= "<div style='color:red;'>Failed to open excel application." & errorMessage.ToString & "</div>"
            Return String.Empty
        End Try

        Dim xclFileExtsn As String
        xclFileExtsn = IIf(CInt(xlApp.version) >= 12, ".xlsx", ".xls") 'see which version of excel you installed in the PC along with IIS (excel 2007 above)



        Try


            Dim rCount As Integer = dtTable.Rows.Count
            Dim cCount As Integer = dtTable.Columns.Count

            Dim dataArray(rCount + 1, cCount) As Object

            With xlBook.Worksheets(1)
                .Cells.Font.Name = "Arial"
                .Cells.Font.Size = 9
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

            errorMessage.Append("R2" & ex2.Message)
            msgRtrn &= "<div style='color:red;'>Failed to create records." & errorMessage.ToString & "</div>"
        Finally

        End Try




        Dim FileName As System.IO.FileInfo, FileNameExist As Boolean = True
        Try
            FileName = New System.IO.FileInfo(filePath & "\" & excelfilename & xclFileExtsn)
        Catch ex3 As Exception
            errorMessage.Append("R3" & ex3.Message)
            msgRtrn &= "<div style='color:red;'>Failed to open an excel file." & errorMessage.ToString & "</div>"
            FileNameExist = False
        End Try

        If FileNameExist Then
            Try
                File.Delete(filePath & "\" & excelfilename & xclFileExtsn)
            Catch ex4 As Exception
                errorMessage.Append("R4" & ex4.Message)
                'Page.ClientScript.RegisterClientScriptBlock(Me.GetType(), "script1", "<script language='javascript'>alert('The file is in use by other program.');</script>", False)
                msgRtrn &= "<div style='color:red;'>The file is in use by other program." & errorMessage.ToString & "</div>"
                Return "The file is in use by other program"
            End Try
        End If

        Try

            Try
                'xlBook.SaveAs(filePath & "\" & excelfilename & xclFileExtsn)
                xlBook.SaveCopyAs(filePath & "\" & excelfilename & xclFileExtsn) 'on server 2008, method of SaveAs is not supported
            Catch ex5 As Exception
                errorMessage.Append("G5" & ex5.Message)
                msgRtrn &= "<div style='color:red;'>Failed to save an excel file." & errorMessage.ToString & "</div>"
            End Try

            xlBook.close()
            xlBook = Nothing
            xlApp.Quit()
            xlApp = Nothing
            'GC.Collect()

        Catch ex6 As Exception
            errorMessage.Append("G6" & ex6.Message)
            msgRtrn &= "<div style='color:red;'>Failed attempt to close excel file." & errorMessage.ToString & "</div>"
            Try
                Marshal.FinalReleaseComObject(xlBook)
                xlBook = Nothing
                Marshal.FinalReleaseComObject(xlApp)
                xlApp = Nothing

                GC.Collect()
                GC.WaitForPendingFinalizers()

            Catch ex As Exception  'below process is a little rude,immediately kill the process
                Try
                    For Each p As Process In Process.GetProcessesByName("Excel")
                        If p.StartTime >= startTimeBeforexlApp AndAlso p.StartTime <= startTimeAfterxlApp Then
                            p.Kill()
                            Exit For
                        End If

                    Next
                Catch ex9 As Exception
                    errorMessage.Append("G9" & ex9.Message)
                    msgRtrn &= "<div style='color:red;'>Failed  to close excel file finally." & errorMessage.ToString & "</div>"
                End Try

            End Try
        End Try



        Return excelfilename & xclFileExtsn

    End Function




    ''' <summary>
    ''' Use to get one excel file from sqlDataSource control
    ''' </summary>
    ''' <param name="SDS1"></param>
    ''' <param name="StatusLabel"></param>
    ''' <remarks></remarks>
    Public Sub downloadExcelFileFromSqlDataSource(ByRef SDS1 As SqlDataSource, ByRef StatusLabel As Label, ByRef hiddenBT As WebControls.HiddenField)

        Dim offsetTimeZone As Integer = -480

        If IsNumeric(hiddenBT.Value) Then
            offsetTimeZone = CInt(hiddenBT.Value)
        End If

        Dim errorMessage As StringBuilder = New StringBuilder()

        'File Path and File Name
        Dim filePath As String = ConfigurationManager.AppSettings("excelFolder")


        If Not Directory.Exists(filePath) Then
            Try
                Directory.CreateDirectory(filePath)
            Catch ex1 As Exception
                errorMessage.Append("P1" & ex1.Message)
                msgPopUP("You might not have right to create folder: " & filePath & errorMessage.ToString, StatusLabel)
                Return
            End Try
        End If

        If filePath.EndsWith("\") Then filePath = filePath.Remove(filePath.Length - 1, 1)
        Dim ExcelReturnMessage As String
        Dim _downloadfilename As String = generateExcel(SDS1, filePath, ExcelReturnMessage)
        Dim fileExtsn As String = ".xlsx"

        errorMessage.Append(ExcelReturnMessage) 'get the error message from function generateExcel

        Dim FileName As System.IO.FileInfo, FileNameExist As Boolean = True
        Dim myFile As FileStream
        Dim _BinaryReader As BinaryReader

        If _downloadfilename.IndexOf(fileExtsn) < 0 Then fileExtsn = ".xls"

        Try
            FileName = New System.IO.FileInfo(filePath & "\" & _downloadfilename)
            myFile = New FileStream(filePath & "\" & _downloadfilename, FileMode.Open, FileAccess.Read, FileShare.Read)
            'Reads file as binary values
            _BinaryReader = New BinaryReader(myFile)
        Catch ex2 As Exception
            errorMessage.Append("P2" & ex2.Message)
            FileNameExist = False
        End Try



        If FileNameExist Then
            Try

                Dim startBytes As Long = 0
                Dim lastUpdateTimeStamp As String = File.GetLastWriteTimeUtc(filePath).ToString("r")
                Dim _EncodedData = HttpUtility.UrlEncode(_downloadfilename, Encoding.UTF8) & lastUpdateTimeStamp

                Response.Clear()
                Response.Buffer = False

                Response.AddHeader("Accept-Ranges", "bytes")
                Response.AppendHeader("ETag", "'" & _EncodedData & "'")
                Response.AppendHeader("Last-Modified", lastUpdateTimeStamp)
                Response.ContentType = "application/octet-stream"
                Response.AddHeader("Content-Disposition", "attachment;filename=" & String.Format(pageTitleReference() & "-{0:yyyy-MM-dd-HH-mm-ss}" & fileExtsn, System.DateTime.Now.ToUniversalTime.AddMinutes(-offsetTimeZone)))
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
                Response.Flush()

            Catch ex As Exception
                errorMessage.Append("P3" & ex.Message)
                msgPopUP("Download status: can not download the file. The following error occured: " & errorMessage.ToString, StatusLabel)
            Finally
                _BinaryReader.Close()
                myFile.Close()
            End Try

            StatusLabel.ForeColor = Drawing.Color.Black
            StatusLabel.Text = "Download status: File is OK now!" & errorMessage.ToString()

        Else
            msgPopUP("Download status: File is NOT available now!" & errorMessage.ToString(), StatusLabel)

        End If

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
    Public Overridable Function fileToServer(ByRef filepath As String, ByRef filename As String, ByRef flextsn As String, ByVal FileUpload1 As FileUpload, ByRef StatusLabel As String) As Boolean
        fileToServer = True



        If FileUpload1.HasFile Then
            'filepath = Server.MapPath("~/excelFiles") & "\"
            filepath = ConfigurationManager.AppSettings("excelFolder") & "\"

            If Not Directory.Exists(filepath) Then
                Try
                    Directory.CreateDirectory(filepath)
                Catch
                    'msgPopUP("You might not have right to create folder: " & filepath, StatusLabel)
                    StatusLabel = "<div style='color:red;'>" & "You might not have right to create folder: " & filepath & "</div>"
                    Return False
                End Try
            End If


            'If filepath.EndsWith("\\") Then filepath = filepath.Replace("\\", "\")
            If Not filepath.EndsWith("\") Then filepath &= "\"

            filename = FileUpload1.FileName
            flextsn = ".xls"
            Dim fileSize As Integer = FileUpload1.PostedFile.ContentLength

            Dim continue1 As Boolean = True

            'make a judgement to see if the file is a excel file
            If filename.ToLower().IndexOf(".xlsx") >= 0 Then
                flextsn = ".xlsx"
            Else
                If filename.ToLower().IndexOf(".xls") < 0 Then
                    'msgPopUP("Only excel files are accepted!", StatusLabel, True, False)
                    StatusLabel &= "<div style='color:red;'>" & "Only excel files are accepted!" & "</div>"
                    Return False
                End If
            End If




            filename = pageTitleReference() & userIden() & "UPLD" & flextsn

            Try
                If fileSize < 100 * 1024 * 1024 Then
                    FileUpload1.SaveAs(filepath & filename)

                    'msgPopUP("File uploaded to server successfully!", StatusLabel, False, False)
                    StatusLabel &= "File uploaded to server successfully!"
                Else
                    StatusLabel &= "<div style='color:red;'>" & "File size can not exceed 100MB!" & "</div>"
                    Return False
                End If
            Catch ex As Exception

                'msgPopUP("The file could not be uploaded. The following error occured: " & ex.Message, StatusLabel, True, False)
                StatusLabel &= "<div style='color:red;'>" & "The file could not be uploaded. The following error occured: " & ex.Message & "</div>"
                Return False
            End Try

        Else
            Return False
        End If


    End Function



    ''' <summary>
    ''' Insert new record to database table from excel file
    ''' </summary>
    ''' <param name="StatusLabel"></param>
    ''' <param name="SDS1">sqlDataSource Control represnt database table</param>
    ''' <param name="LV1">a listview control to display data, to get DataKeys by reference to it</param>
    ''' <param name="filepath">excel file path</param>
    ''' <param name="filename">excel file name</param>
    ''' <param name="flextsn">excel file suffix</param>
    ''' <returns>return warning message or others</returns>
    Public Overridable Function dataInsertionToDatabaseTableFromExcel(ByRef StatusLabel As Label, ByRef SDS1 As SqlDataSource, ByRef LV1 As ListView, ByVal filepath As String, ByVal filename As String, ByVal flextsn As String) As String

        Dim msgRtrn As New StringBuilder

        Dim continue1 As Boolean = True
        Dim howManyRecordsInserted As Integer = 0


        If Not Directory.Exists(filepath) Then
            Try
                Directory.CreateDirectory(filepath)
            Catch
                continue1 = False
                Return "<div style='color:red;'>" & "You might not have right to create folder: " & filepath & "</div>"

            End Try
        End If



        'Dim warningMessage As StringBuilder = New StringBuilder()
        Dim excelconnectionstr As String
        If flextsn = ".xls" Then
            excelconnectionstr = String.Format("provider=Microsoft.Jet.OleDb.4.0; Data Source='{0}';" & "Extended Properties='Excel 8.0;HDR=YES;IMEX=1'", filepath & filename)
        Else
            excelconnectionstr = String.Format("provider=Microsoft.ACE.OleDb.12.0; Data Source='{0}';" & "Extended Properties='Excel 12.0 Xml;HDR=YES;IMEX=1'", filepath & filename)
        End If

        Dim connexcl As New System.Data.OleDb.OleDbConnection(excelconnectionstr)
        connexcl.Open()

        Dim frstSheetName As String = "Sheet1$"


        'Try
        '    ' Get the name of the first worksheet:
        '    Dim dbSchema As DataTable = connexcl.GetSqlSchemaTable(SqlSchemaGuid.Tables, Nothing)

        '    If (dbSchema Is Nothing OrElse dbSchema.Rows.Count < 1) Then
        '    End If

        '    frstSheetName = dbSchema.Rows(0)("TABLE_NAME").ToString()

        'Catch ex As Exception
        '    msgRtrn.AppendLine("<div style='color:red;'>" & ex.Message & "</div>")
        '    continue1 = False

        'End Try

        Dim excelslctsql As String = SDS1.SelectCommand.ToString
        excelslctsql = excelslctsql.Substring(0, excelslctsql.IndexOf(" FROM "))
        Dim dtADPexcel As OleDbDataAdapter = New OleDbDataAdapter(excelslctsql & " FROM [" & frstSheetName & "A1:" & rightmostColumn() & maxRowNumber.ToString() & "] ", connexcl)
        Dim dataDS1 As DataSet = New DataSet



        Try
            dtADPexcel.Fill(dataDS1, "update")
        Catch ex As Exception
            continue1 = False
            'warningMessage.Append("Maybe you did not get the right excel file! Or columns missed, or inserted row number exceeds " & maxRowNumber.ToString() & "<br />")
            msgRtrn.AppendLine("<div style='color:red;'>" & ex.Message & "Maybe you did not get the right excel file! Row number limit is " & maxRowNumber.ToString() & "</div>")
        End Try




        If continue1 Then

            Dim msstring As String = String.Empty
            'check data validity
            msstring = dataValidityCheck(dataDS1.Tables("update"))

            If Not String.IsNullOrEmpty(msstring) Then
                continue1 = False
                'msgPopUP(msstring, StatusLabel, True)
                msgRtrn.AppendLine("<div style='color:red;'>" & msstring & "</div>")
            End If
        End If



        Dim connstr As String = SDS1.ConnectionString
        Dim conn As SqlConnection = New SqlConnection(connstr)
        'conn.Open()
        Dim dtAdapter1 As SqlDataAdapter = New SqlDataAdapter(SDS1.SelectCommand.ToString(), conn)


        Dim cmdbAccessCmdBuilder As New SqlCommandBuilder(dtAdapter1)
        dtAdapter1.InsertCommand = cmdbAccessCmdBuilder.GetInsertCommand()
        'dtAdapter1.UpdateCommand = cmdbAccessCmdBuilder.GetUpdateCommand()




        If continue1 Then


            dtAdapter1.Fill(dataDS1, "tobeupdated")

            Dim primaryColumnsCount As Integer = LV1.DataKeyNames.Count

            Dim filterSQL As String = String.Empty
            Dim keys(primaryColumnsCount - 1) As DataColumn
            Dim keyNames(primaryColumnsCount - 1) As String
            'Dim childKeys(primaryColumnsCount - 1) As DataColumn
            For i As Integer = 0 To primaryColumnsCount - 1
                keys(i) = dataDS1.Tables("tobeupdated").Columns(LV1.DataKeyNames(i))
                'childKeys(i) = dataDS1.Tables("update").Columns(LV1.DataKeyNames(i))

                keyNames(i) = LV1.DataKeyNames(i)
                If dataDS1.Tables("tobeupdated").Columns(keyNames(i)).DataType = GetType(String) Then
                    filterSQL &= keyNames(i) & " = '@" & i & "' And "
                Else
                    filterSQL &= keyNames(i) & " = @" & i & " And "
                End If
            Next
            dataDS1.Tables("tobeupdated").PrimaryKey = keys
            filterSQL = filterSQL.Remove(filterSQL.Length - 4, 4)


            'Dim filterSQL As String = String.Empty
            'Dim keyCount As Integer = LV1.DataKeyNames.Count
            'Dim keyNames(keyCount - 1) As String
            'For i As Integer = 0 To keyCount - 1
            '    'keyNames(i) = LV1.DataKeyNames(i)
            '    If dataDS1.Tables("tobeupdated").Columns(keyNames(i)).DataType = GetType(String) Then
            '        filterSQL &= keyNames(i) & " = '@" & i & "' And "
            '    Else
            '        filterSQL &= keyNames(i) & " = @" & i & " And "
            '    End If
            'Next
            'filterSQL = filterSQL.Remove(filterSQL.Length - 4, 4)





            'check if there are duplicate records in the update file
            If continue1 Then

                Dim sortedColumns As String = String.Empty
                For i As Integer = 0 To primaryColumnsCount - 1
                    sortedColumns &= keyNames(i) & " ASC,"
                Next
                sortedColumns = sortedColumns.Remove(sortedColumns.Length - 1, 1)

                'firstly sort the data for further process
                Dim sortedRow() As DataRow = dataDS1.Tables("update").Select(Nothing, sortedColumns)

                'to see if there two nearby rows have the same value for all the key columns
                Dim completeEqual As Boolean

                For i As Integer = 1 To sortedRow.Count - 1
                    completeEqual = True
                    For j As Integer = 0 To primaryColumnsCount - 1
                        If sortedRow(i).Item(keyNames(j)) = sortedRow(i - 1).Item(keyNames(j)) Then
                        Else
                            completeEqual = False
                        End If
                    Next

                    If completeEqual Then
                        Dim rowInformation As StringBuilder = New StringBuilder
                        For k As Integer = 0 To primaryColumnsCount - 1
                            rowInformation.Append(sortedRow(i).Item(keyNames(k)) & ",")
                        Next
                        rowInformation.Remove(rowInformation.Length - 1, 1)
                        continue1 = False
                        'warningMessage.Append("Based on the combination of " & String.Join(",", keyNames) & " ,duplicate records exist in your excel file:" & rowInformation.ToString() & "<br />")
                        msgRtrn.AppendLine("<div style='color:red;'>" & "Based on the combination of " & String.Join(",", keyNames) & " ,duplicate records exist in your excel file:" & rowInformation.ToString() & "</div>")
                        Exit For
                    End If

                Next

            End If



            'to see if the record you are inserting from the excel file does exist in the database

            If continue1 Then
                Dim filterClause As String

                Try
                    For Each updateRow As DataRow In dataDS1.Tables("update").Rows
                        filterClause = filterSQL
                        For i As Integer = 0 To primaryColumnsCount - 1
                            filterClause = filterClause.Replace("@" & i, updateRow.Item(keyNames(i)))
                        Next

                        Dim toupdateRows() As DataRow = dataDS1.Tables("tobeupdated").Select(filterClause)

                        If toupdateRows.Count = 0 Then

                            dataDS1.Tables("tobeupdated").Rows.Add(updateRow.ItemArray)
                            howManyRecordsInserted += 1
                        End If

                    Next

                Catch ex As Exception
                    'warningMessage.Append(ex.Message & " : related record in Excel : " & filterClause)
                    msgRtrn.AppendLine("<div style='color:red;'>" & ex.Message & " : related record in Excel : " & filterClause & "</div>")
                End Try

            End If

        End If


        If continue1 Then
            Try
                'dtAdapter1.UpdateBatchSize = 512
                dtAdapter1.Update(dataDS1, "tobeupdated")
                'msgPopUP("Action of insertion upon excel file is completed. The number of inserted records is " & howManyRecordsInserted, StatusLabel, False, False)
                msgRtrn.AppendLine("Action of insertion upon excel file is completed. The number of inserted records is " & howManyRecordsInserted)
            Catch ex As Exception
                'warningMessage.Append("Something wrong when inserting.<br />")
                msgRtrn.AppendLine("<div style='color:red;'>" & "Something wrong when inserting." & ex.Message & "</div>")
            End Try


        End If

        dataDS1.Dispose()

        cmdbAccessCmdBuilder.Dispose()

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


        Return msgRtrn.ToString

    End Function



    ''' <summary>
    ''' Insert new record to database table from excel file
    ''' </summary>
    ''' <param name="StatusLabel"></param>
    ''' <param name="SDS1">sqlDataSource Control represnt database table</param>
    ''' <param name="LV1">a listview control to display data, to get DataKeys by reference to it</param>
    ''' <param name="filepath">excel file path</param>
    ''' <param name="filename">excel file name</param>
    ''' <param name="flextsn">excel file suffix</param>
    ''' <param name="continueOrnot">out parameter</param>
    ''' <returns>return warning message or others</returns>
    Public Overridable Function dataInsertionToDatabaseTableFromExcel(ByRef StatusLabel As Label, ByRef SDS1 As SqlDataSource, ByRef LV1 As ListView, ByVal filepath As String, ByVal filename As String, ByVal flextsn As String, ByRef continueOrnot As Boolean) As String

        Dim msgRtrn As New StringBuilder

        Dim continue1 As Boolean = True
        Dim howManyRecordsInserted As Integer = 0


        If Not Directory.Exists(filepath) Then
            Try
                Directory.CreateDirectory(filepath)
            Catch
                continue1 = False
                Return "<div style='color:red;'>" & "You might not have right to create folder: " & filepath & "</div>"

            End Try
        End If



        'Dim warningMessage As StringBuilder = New StringBuilder()
        Dim excelconnectionstr As String
        If flextsn = ".xls" Then
            excelconnectionstr = String.Format("provider=Microsoft.Jet.OleDb.4.0; Data Source='{0}';" & "Extended Properties='Excel 8.0;HDR=YES;IMEX=1'", filepath & filename)
        Else
            excelconnectionstr = String.Format("provider=Microsoft.ACE.OleDb.12.0; Data Source='{0}';" & "Extended Properties='Excel 12.0 Xml;HDR=YES;IMEX=1'", filepath & filename)
        End If

        Dim connexcl As New System.Data.OleDb.OleDbConnection(excelconnectionstr)
        connexcl.Open()

        Dim frstSheetName As String = "Sheet1$"


        'Try
        '    ' Get the name of the first worksheet:
        '    Dim dbSchema As DataTable = connexcl.GetSqlSchemaTable(SqlSchemaGuid.Tables, Nothing)

        '    If (dbSchema Is Nothing OrElse dbSchema.Rows.Count < 1) Then
        '    End If

        '    frstSheetName = dbSchema.Rows(0)("TABLE_NAME").ToString()

        'Catch ex As Exception
        '    msgRtrn.AppendLine("<div style='color:red;'>" & ex.Message & "</div>")
        '    continue1 = False

        'End Try

        Dim excelslctsql As String = SDS1.SelectCommand.ToString
        excelslctsql = excelslctsql.Substring(0, excelslctsql.IndexOf(" FROM "))
        Dim dtADPexcel As OleDbDataAdapter = New OleDbDataAdapter(excelslctsql & " FROM [" & frstSheetName & "A1:" & rightmostColumn() & maxRowNumber.ToString() & "] ", connexcl)
        Dim dataDS1 As DataSet = New DataSet



        Try
            dtADPexcel.Fill(dataDS1, "update")
        Catch ex As Exception
            continue1 = False
            'warningMessage.Append("Maybe you did not get the right excel file! Or columns missed, or inserted row number exceeds " & maxRowNumber.ToString() & "<br />")
            msgRtrn.AppendLine("<div style='color:red;'>" & ex.Message & "Maybe you did not get the right excel file! Row number limit is " & maxRowNumber.ToString() & "</div>")
        End Try




        'If continue1 Then

        '    Dim msstring As String = String.Empty
        '    'check data validity
        '    msstring = dataValidityCheck(dataDS1.Tables("update"))

        '    If Not String.IsNullOrEmpty(msstring) Then
        '        continue1 = False
        '        'msgPopUP(msstring, StatusLabel, True)
        '        msgRtrn.AppendLine("<div style='color:red;'>" & msstring & "</div>")
        '    End If
        'End If



        Dim connstr As String = SDS1.ConnectionString
        Dim conn As SqlConnection = New SqlConnection(connstr)
        'conn.Open()
        Dim dtAdapter1 As SqlDataAdapter = New SqlDataAdapter(SDS1.SelectCommand.ToString(), conn)


        Dim cmdbAccessCmdBuilder As New SqlCommandBuilder(dtAdapter1)
        dtAdapter1.InsertCommand = cmdbAccessCmdBuilder.GetInsertCommand()
        'dtAdapter1.UpdateCommand = cmdbAccessCmdBuilder.GetUpdateCommand()




        If continue1 Then


            dtAdapter1.Fill(dataDS1, "tobeupdated")

            Dim primaryColumnsCount As Integer = LV1.DataKeyNames.Count

            'Dim filterSQL As String = String.Empty
            Dim keys(primaryColumnsCount - 1) As DataColumn
            'Dim keyNames(primaryColumnsCount - 1) As String
            'Dim childKeys(primaryColumnsCount - 1) As DataColumn
            For i As Integer = 0 To primaryColumnsCount - 1
                keys(i) = dataDS1.Tables("tobeupdated").Columns(LV1.DataKeyNames(i))
                'childKeys(i) = dataDS1.Tables("update").Columns(LV1.DataKeyNames(i))

                'keyNames(i) = LV1.DataKeyNames(i)
                'If dataDS1.Tables("tobeupdated").Columns(keyNames(i)).DataType = GetType(String) Then
                '    filterSQL &= keyNames(i) & " = '@" & i & "' And "
                'Else
                '    filterSQL &= keyNames(i) & " = @" & i & " And "
                'End If
            Next
            dataDS1.Tables("tobeupdated").PrimaryKey = keys
            'filterSQL = filterSQL.Remove(filterSQL.Length - 4, 4)

        End If


        If continue1 Then
            Try

                'insert records
                dtADPexcel.FillLoadOption = LoadOption.Upsert
                dtADPexcel.Fill(dataDS1, "tobeupdated")

                howManyRecordsInserted = dataDS1.Tables("tobeupdated").Select(Nothing, Nothing, System.Data.DataViewRowState.Added).Count

                'dtAdapter1.UpdateBatchSize = 512
                dtAdapter1.Update(dataDS1, "tobeupdated")

                msgRtrn.AppendLine("Action of insertion upon excel file is completed. The number of inserted records is " & howManyRecordsInserted)
            Catch ex As Exception
                'warningMessage.Append("Something wrong when inserting.<br />")
                msgRtrn.AppendLine("<div style='color:red;'>" & "Something wrong when inserting." & ex.Message & "</div>")
            End Try


        End If

        dataDS1.Dispose()

        cmdbAccessCmdBuilder.Dispose()

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



        Return msgRtrn.ToString

    End Function






    ''' <summary>
    ''' update data in database table based on the input from excel file
    ''' </summary>
    ''' <param name="StatusLabel"></param>
    ''' <param name="SDS1">sqlDataSource Control represnt database table</param>
    ''' <param name="LV1">a listview control to display data, to get DataKeys by reference to it</param>
    ''' <param name="filepath">excel file path</param>
    ''' <param name="filename">excel file name</param>
    ''' <param name="flextsn">excel file suffix</param>
    Public Overridable Function dataUpdatedToDatabaseTablePerExcelData(ByRef StatusLabel As Label, ByRef SDS1 As SqlDataSource, ByRef LV1 As ListView, ByVal filepath As String, ByVal filename As String, ByVal flextsn As String) As String

        Dim msgRtrn As New StringBuilder

        Dim continue1 As Boolean = True
        Dim howManyRecordsUpdated As Integer = 0

        If Not Directory.Exists(filepath) Then
            Try
                Directory.CreateDirectory(filepath)
            Catch ex As Exception
                continue1 = False
                'msgPopUP("You might not have right to create folder: " & filepath, StatusLabel)
                Return "<div style='color:red;'>" & ex.Message & ". You might not have right to create folder: " & filepath & "</div>"
            End Try
        End If



        Dim excelconnectionstr As String
        If flextsn = ".xls" Then
            excelconnectionstr = String.Format("provider=Microsoft.Jet.OleDb.4.0; Data Source='{0}';" & "Extended Properties='Excel 8.0;HDR=YES;IMEX=1'", filepath & filename)
        Else
            excelconnectionstr = String.Format("provider=Microsoft.ACE.OleDb.12.0; Data Source='{0}';" & "Extended Properties='Excel 12.0 Xml;HDR=YES;IMEX=1'", filepath & filename)
        End If

        Dim connexcl As New System.Data.Oledb.OledbConnection(excelconnectionstr)
        connexcl.Open()

        Dim frstSheetName As String = "Sheet1$"

        'Try
        '    ' Get the name of the first worksheet:
        '    Dim dbSchema As DataTable = connexcl.GetSqlSchemaTable(SqlSchemaGuid.Tables, Nothing)

        '    If (dbSchema Is Nothing OrElse dbSchema.Rows.Count < 1) Then
        '    End If

        '    frstSheetName = dbSchema.Rows(0)("TABLE_NAME").ToString()

        'Catch ex As Exception
        '    msgRtrn.AppendLine("<div style='color:red;'>" & ex.Message & "</div>")
        '    continue1 = False

        'End Try

        Dim excelslctsql As String = SDS1.SelectCommand.ToString
        excelslctsql = excelslctsql.Substring(0, excelslctsql.IndexOf(" FROM "))
        Dim dtADPexcel As OleDbDataAdapter = New OleDbDataAdapter(excelslctsql & " FROM [" & frstSheetName & "A1:" & rightmostColumn() & maxRowNumber.ToString() & "] ", connexcl)
        'Dim dtADPexcel As SqlDataAdapter = New SqlDataAdapter("Select * From [" & frstSheetName & "A1:BT15001] ", connexcl)
        Dim dataDS1 As DataSet = New DataSet
        Try
            dtADPexcel.Fill(dataDS1, "update")
        Catch ex As Exception
            'Page.ClientScript.RegisterClientScriptBlock(Me.GetType(), "script1", "<script language='javascript'>alert('Maybe you did not get the right excel file! row number limit is '" & maxRowNumber.ToString() & ");</script>", False)
            continue1 = False
            msgRtrn.AppendLine("<div style='color:red;'>" & ex.Message & " Maybe you did not get the right excel file! row number limit is " & maxRowNumber.ToString() & "</div>")
        End Try


        'Dim db As String = ConfigurationManager.ConnectionStrings(dbConnectionName).ConnectionString
        'Dim connstr As String = ConfigurationManager.ConnectionStrings(dbConnectionName).ProviderName & db
        Dim connstr As String = SDS1.ConnectionString
        Dim conn As SqlConnection = New SqlConnection(connstr)
        'conn.Open()
        Dim dtAdapter1 As SqlDataAdapter = New SqlDataAdapter(SDS1.SelectCommand.ToString(), conn)



        'use aunto command generation mechnism to generate standard insert SQL clause
        Dim cmdbAccessCmdBuilder As New SqlCommandBuilder(dtAdapter1)
        dtAdapter1.UpdateCommand = cmdbAccessCmdBuilder.GetUpdateCommand()


        Try


            dtAdapter1.Fill(dataDS1, "tobeupdated")

            Dim primaryColumnsCount As Integer = LV1.DataKeyNames.Count

            Dim keys(primaryColumnsCount - 1) As DataColumn
            Dim childKeys(primaryColumnsCount - 1) As DataColumn

            Try
                For i As Integer = 0 To primaryColumnsCount - 1
                    keys(i) = dataDS1.Tables("tobeupdated").Columns(LV1.DataKeyNames(i))
                    childKeys(i) = dataDS1.Tables("update").Columns(LV1.DataKeyNames(i))
                Next

                dataDS1.Tables("tobeupdated").PrimaryKey = keys

            Catch ex As Exception
                continue1 = False
                msgRtrn.AppendLine("<div style='color:red;'>" & ex.Message & " Key columns among Excel file and database table do not match." & "</div>")
            End Try





            ''try to avoid inserting new records when going to do updating operation
            'If False Then
            '    Try
            '        Dim drel As DataRelation = New DataRelation("ExcltoAccss", keys, childKeys)
            '        dataDS1.Relations.Add(drel)

            '    Catch ex As Exception
            '        continue1 = False
            '        'msgPopUP("There some new rows in the update excel file.Updating only, not allow inserting new records.", StatusLabel)
            '        msgRtrn.AppendLine("<div style='color:red;'>" & ex.Message & "There some new rows in the update excel file.Updating only, not allow inserting new records." & "</div>")
            '    End Try
            'End If




            If continue1 Then

                Dim msstring As String = String.Empty
                'check data validity
                msstring = dataValidityCheck(dataDS1.Tables("update"))

                If Not String.IsNullOrEmpty(msstring) Then
                    continue1 = False
                    'msgPopUP(msstring, StatusLabel, True)
                    msgRtrn.AppendLine("<div style='color:red;'>" & msstring & "</div>")
                End If
            End If

            If continue1 Then
                Try
                    dtADPexcel.FillLoadOption = LoadOption.Upsert
                    dtADPexcel.Fill(dataDS1, "tobeupdated")

                    'delete those newly added because this function is for updating
                    For Each a As DataRow In dataDS1.Tables("tobeupdated").Select(Nothing, Nothing, System.Data.DataViewRowState.Added)
                        a.Delete()
                    Next

                    howManyRecordsUpdated = dataDS1.Tables("tobeupdated").Select(Nothing, Nothing, System.Data.DataViewRowState.ModifiedCurrent).Count


                    'dtAdapter1.UpdateBatchSize = 512  'access database does not support this
                    dtAdapter1.Update(dataDS1, "tobeupdated")
                    'msgPopUP("Action of update upon the excel file is completed. The number of updated records is " & howManyRecordsUpdated, StatusLabel, False, False)
                    msgRtrn.AppendLine("Action of update upon the excel file is completed. The number of updated records is " & howManyRecordsUpdated)
                Catch ex As Exception
                    'msgPopUP("Something wrong with the update operation.", StatusLabel, True, True)
                    msgRtrn.AppendLine("<div style='color:red;'>" & "Something wrong with the update operation." & ex.Message & "</div>")
                End Try

            End If

        Catch ex As Exception
            msgRtrn.AppendLine("<div style='color:red;'>" & ex.Message & "</div>")
        End Try

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


        Return msgRtrn.ToString

    End Function





    ''' <summary>
    ''' delete records in database table based on the input from excel file
    ''' </summary>
    ''' <param name="StatusLabel"></param>
    ''' <param name="SDS1">sqlDataSource Control represnt database table</param>
    ''' <param name="LV1">a listview control to display data, to get DataKeys by reference to it</param>
    ''' <param name="filepath">excel file path</param>
    ''' <param name="filename">excel file name</param>
    ''' <param name="flextsn">excel file suffix</param>
    Public Overridable Function dataDeletedFromDatabaseTablePerExcelData(ByRef StatusLabel As Label, ByRef SDS1 As SqlDataSource, ByRef LV1 As ListView, ByVal filepath As String, ByVal filename As String, ByVal flextsn As String) As String


        Dim msgRtrn As New StringBuilder

        Dim continue1 As Boolean = True
        Dim howManyRecordsDeleted As Integer = 0

        If Not Directory.Exists(filepath) Then
            Try
                Directory.CreateDirectory(filepath)
            Catch ex As Exception
                continue1 = False
                'msgPopUP("You might not have right to create folder: " & filepath, StatusLabel)
                Return "<div style='color:red;'>" & ex.Message & ". You might not have right to create folder: " & filepath & "</div>"
            End Try
        End If

        Dim excelconnectionstr As String
        If flextsn = ".xls" Then
            excelconnectionstr = String.Format("provider=Microsoft.Jet.OleDb.4.0; Data Source='{0}';" & "Extended Properties='Excel 8.0;HDR=YES;IMEX=1'", filepath & filename)
        Else
            excelconnectionstr = String.Format("provider=Microsoft.ACE.OleDb.12.0; Data Source='{0}';" & "Extended Properties='Excel 12.0 Xml;HDR=YES;IMEX=1'", filepath & filename)
        End If

        Dim connexcl As New System.Data.Oledb.OledbConnection(excelconnectionstr)
        connexcl.Open()

        Dim frstSheetName As String = "Sheet1$"


        'Try
        '    ' Get the name of the first worksheet:
        '    Dim dbSchema As DataTable = connexcl.GetSqlSchemaTable(SqlSchemaGuid.Tables, Nothing)

        '    If (dbSchema Is Nothing OrElse dbSchema.Rows.Count < 1) Then
        '    End If

        '    frstSheetName = dbSchema.Rows(0)("TABLE_NAME").ToString()
        'Catch ex As Exception
        '    msgRtrn.AppendLine("<div style='color:red;'>" & ex.Message & "</div>")
        '    continue1 = False

        'End Try

        Dim excelslctsql As String = SDS1.SelectCommand.ToString
        excelslctsql = excelslctsql.Substring(0, excelslctsql.IndexOf(" FROM "))
        Dim dtADPexcel As OleDbDataAdapter = New OleDbDataAdapter(excelslctsql & " FROM [" & frstSheetName & "A1:" & rightmostColumn() & maxRowNumber.ToString() & "] ", connexcl)
        Dim dataDS1 As DataSet = New DataSet
        Try
            dtADPexcel.Fill(dataDS1, "update")
        Catch ex As Exception
            'Page.ClientScript.RegisterClientScriptBlock(Me.GetType(), "script1", "<script language='javascript'>alert('Maybe you did not get the right excel file! Or columns missed, or row number exceeds 15001 .');</script>", False)
            continue1 = False
            'msgPopUP("Maybe you did not get the right excel file! Or columns missed, or row number exceeds " & maxRowNumber.ToString(), StatusLabel)
            msgRtrn.AppendLine("<div style='color:red;'>" & ex.Message & "Maybe you did not get the right excel file! Or columns missed, or row number exceeds " & maxRowNumber.ToString() & "</div>")
        End Try


        Dim connstr As String = SDS1.ConnectionString
        Dim conn As SqlConnection = New SqlConnection(connstr)
        'conn.Open()
        Dim dtAdapter1 As SqlDataAdapter = New SqlDataAdapter(SDS1.SelectCommand.ToString(), conn)
        'Dim keyname00 As String = LV1.DataKeyNames(0)  'primary key

        'use aunto command generation mechnism to generate standard insert SQL clause
        Dim cmdbAccessCmdBuilder As New SqlCommandBuilder(dtAdapter1)
        dtAdapter1.DeleteCommand = cmdbAccessCmdBuilder.GetDeleteCommand


        If continue1 Then
            Try
                dtAdapter1.Fill(dataDS1, "tobeupdated")

                Dim primaryColumnsCount As Integer = LV1.DataKeyNames.Count
                Dim filterSQL As String = String.Empty
                Dim keys(primaryColumnsCount - 1) As DataColumn
                Dim keyNames(primaryColumnsCount - 1) As String
                'Dim childKeys(primaryColumnsCount - 1) As DataColumn
                For i As Integer = 0 To primaryColumnsCount - 1
                    keys(i) = dataDS1.Tables("tobeupdated").Columns(LV1.DataKeyNames(i))
                    'childKeys(i) = dataDS1.Tables("update").Columns(LV1.DataKeyNames(i))

                    keyNames(i) = LV1.DataKeyNames(i)
                    If dataDS1.Tables("tobeupdated").Columns(keyNames(i)).DataType = GetType(String) Then 'placeholder for key fields
                        filterSQL &= keyNames(i) & " = '@" & i & "' And "
                    Else
                        filterSQL &= keyNames(i) & " = @" & i & " And "
                    End If
                Next
                dataDS1.Tables("tobeupdated").PrimaryKey = keys
                filterSQL = filterSQL.Remove(filterSQL.Length - 4, 4)


                'Dim filterSQL As String = String.Empty
                'Dim keyCount As Integer = LV1.DataKeyNames.Count
                'Dim keyNames(keyCount - 1) As String

                'For i As Integer = 0 To keyCount - 1
                '    keyNames(i) = LV1.DataKeyNames(i)
                '    If dataDS1.Tables("tobeupdated").Columns(i).DataType = GetType(String) Then 'placeholder for key fields
                '        filterSQL &= LV1.DataKeyNames(i) & " = '@" & i & "' And "
                '    Else
                '        filterSQL &= LV1.DataKeyNames(i) & " = @" & i & " And "
                '    End If
                'Next
                'filterSQL = filterSQL.Remove(filterSQL.Length - 4, 4)



                'check if there are duplicate records in the update file
                If continue1 Then

                    Dim sortedColumns As String = String.Empty
                    For i As Integer = 0 To primaryColumnsCount - 1
                        sortedColumns &= keyNames(i) & " ASC,"
                    Next
                    sortedColumns = sortedColumns.Remove(sortedColumns.Length - 1, 1)

                    'firstly sort the data for further process
                    Dim sortedRow() As DataRow = dataDS1.Tables("update").Select(Nothing, sortedColumns)

                    'to see if there two nearby rows have the same value for all the key columns
                    Dim completeEqual As Boolean

                    For i As Integer = 1 To sortedRow.Count - 1
                        completeEqual = True
                        For j As Integer = 0 To primaryColumnsCount - 1
                            If sortedRow(i).Item(keyNames(j)) = sortedRow(i - 1).Item(keyNames(j)) Then
                            Else
                                completeEqual = False
                            End If
                        Next

                        If completeEqual Then
                            Dim rowInformation As String = String.Empty
                            For k As Integer = 0 To primaryColumnsCount - 1
                                rowInformation &= sortedRow(i).Item(keyNames(k)) & ","
                            Next
                            rowInformation = rowInformation.Substring(0, rowInformation.Length - 1)
                            continue1 = False
                            'msgPopUP("Duplicate records exist in your  excel file: " & rowInformation, StatusLabel)
                            msgRtrn.AppendLine("<div style='color:red;'>" & "Duplicate records exist in your  excel file: " & rowInformation & "</div>")
                            Exit For
                        End If

                    Next

                End If


                'mark row status as delete if there is corresponding record found in table "tobeupdated"
                If continue1 Then
                    Dim filterClause As String
                    For Each updateRow As DataRow In dataDS1.Tables("update").Rows
                        filterClause = filterSQL
                        For i As Integer = 0 To primaryColumnsCount - 1
                            filterClause = filterClause.Replace("@" & i, updateRow.Item(keyNames(i)))
                        Next

                        Dim toupdateRows() As DataRow = dataDS1.Tables("tobeupdated").Select(filterClause)

                        If toupdateRows.Count > 0 Then
                            toupdateRows(0).Delete()
                            howManyRecordsDeleted += 1
                            'Else
                            'continue1 = False
                            'msgPopUP("You are asking to delete a record which does not exist in the database." & filterClause, StatusLabel)
                            'Page.ClientScript.RegisterClientScriptBlock(Me.GetType(), "abc", "<script language='javascript'>alert('You are asking to delete a record which does not exist in the database.'" & filterClause & ");</script>", False)
                            'Exit For
                        End If
                    Next

                End If


            Catch ex As Exception
                continue1 = False
                'msgPopUP("Database maybe do not have the records you are asking to delete.", StatusLabel)
                msgRtrn.AppendLine("<div style='color:red;'>" & "Database maybe do not have the records you are asking to delete." & ex.Message & "</div>")
            End Try

        End If

        If continue1 Then
            Try
                'delete records
                'dtAdapter1.UpdateBatchSize = 512
                dtAdapter1.Update(dataDS1, "tobeupdated")
                'msgPopUP("Action on delection upon excel file is completed. The number of deleted records is " & howManyRecordsDeleted, StatusLabel, False, False)
                msgRtrn.AppendLine("Action on delection upon excel file is completed. The number of deleted records is " & howManyRecordsDeleted)
            Catch ex As Exception
                'msgPopUP("Something wrong when deleting.", StatusLabel, False, False)
                msgRtrn.AppendLine("<div style='color:red;'>" & "Something wrong when deleting." & ex.Message & "</div>")
            End Try


        End If

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

        Return msgRtrn.ToString

    End Function


    ''' <summary>
    ''' delete records in database table based on the input from excel file
    ''' </summary>
    ''' <param name="StatusLabel"></param>
    ''' <param name="SDS1">sqlDataSource Control represnt database table</param>
    ''' <param name="LV1">a listview control to display data, to get DataKeys by reference to it</param>
    ''' <param name="filepath">excel file path</param>
    ''' <param name="filename">excel file name</param>
    ''' <param name="flextsn">excel file suffix</param>
    ''' <param name="continueOrnot">out parameter to show whether the next following step should continue or no</param>
    Public Overridable Function dataDeletedFromDatabaseTablePerExcelData(ByRef StatusLabel As Label, ByRef SDS1 As SqlDataSource, ByRef LV1 As ListView, ByVal filepath As String, ByVal filename As String, ByVal flextsn As String, ByRef continueOrnot As Boolean) As String

        continueOrnot = False

        Dim msgRtrn As New StringBuilder

        Dim continue1 As Boolean = True
        Dim howManyRecordsDeleted As Integer = 0

        If Not Directory.Exists(filepath) Then
            Try
                Directory.CreateDirectory(filepath)
            Catch ex As Exception
                continue1 = False
                'msgPopUP("You might not have right to create folder: " & filepath, StatusLabel)
                Return "<div style='color:red;'>" & ex.Message & ". You might not have right to create folder: " & filepath & "</div>"
            End Try
        End If

        Dim excelconnectionstr As String
        If flextsn = ".xls" Then
            excelconnectionstr = String.Format("provider=Microsoft.Jet.OleDb.4.0; Data Source='{0}';" & "Extended Properties='Excel 8.0;HDR=YES;IMEX=1'", filepath & filename)
        Else
            excelconnectionstr = String.Format("provider=Microsoft.ACE.OleDb.12.0; Data Source='{0}';" & "Extended Properties='Excel 12.0 Xml;HDR=YES;IMEX=1'", filepath & filename)
        End If

        Dim connexcl As New System.Data.Oledb.OledbConnection(excelconnectionstr)
        connexcl.Open()

        Dim frstSheetName As String = "Sheet1$"


        'Try
        '    ' Get the name of the first worksheet:
        '    Dim dbSchema As DataTable = connexcl.GetSqlSchemaTable(SqlSchemaGuid.Tables, Nothing)

        '    If (dbSchema Is Nothing OrElse dbSchema.Rows.Count < 1) Then
        '    End If

        '    frstSheetName = dbSchema.Rows(0)("TABLE_NAME").ToString()
        'Catch ex As Exception
        '    msgRtrn.AppendLine("<div style='color:red;'>" & ex.Message & "</div>")
        '    continue1 = False

        'End Try

        Dim excelslctsql As String = SDS1.SelectCommand.ToString
        excelslctsql = excelslctsql.Substring(0, excelslctsql.IndexOf(" FROM "))
        Dim dtADPexcel As OleDbDataAdapter = New OleDbDataAdapter(excelslctsql & " FROM [" & frstSheetName & "A1:" & rightmostColumn() & maxRowNumber.ToString() & "] ", connexcl)
        Dim dataDS1 As DataSet = New DataSet
        Try
            dtADPexcel.Fill(dataDS1, "update")
        Catch ex As Exception
            'Page.ClientScript.RegisterClientScriptBlock(Me.GetType(), "script1", "<script language='javascript'>alert('Maybe you did not get the right excel file! Or columns missed, or row number exceeds 15001 .');</script>", False)
            continue1 = False
            'msgPopUP("Maybe you did not get the right excel file! Or columns missed, or row number exceeds " & maxRowNumber.ToString(), StatusLabel)
            msgRtrn.AppendLine("<div style='color:red;'>" & ex.Message & "Maybe you did not get the right excel file! Or columns missed, or row number exceeds " & maxRowNumber.ToString() & "</div>")
        End Try


        If continue1 Then

            Dim msstring As String = String.Empty
            'check data validity
            msstring = dataValidityCheck(dataDS1.Tables("update"))

            If Not String.IsNullOrEmpty(msstring) Then
                continue1 = False
                'msgPopUP(msstring, StatusLabel, True)
                msgRtrn.AppendLine("<div style='color:red;'>" & msstring & "</div>")
            End If
        End If



        Dim connstr As String = SDS1.ConnectionString
        Dim conn As SqlConnection = New SqlConnection(connstr)
        'conn.Open()
        Dim dtAdapter1 As SqlDataAdapter = New SqlDataAdapter(SDS1.SelectCommand.ToString(), conn)


        'use aunto command generation mechnism to generate standard insert SQL clause
        Dim cmdbAccessCmdBuilder As New SqlCommandBuilder(dtAdapter1)
        dtAdapter1.DeleteCommand = cmdbAccessCmdBuilder.GetDeleteCommand


        If continue1 Then
            Try
                dtAdapter1.Fill(dataDS1, "tobeupdated")

                Dim primaryColumnsCount As Integer = LV1.DataKeyNames.Count
                Dim filterSQL As String = String.Empty
                Dim keys(primaryColumnsCount - 1) As DataColumn
                Dim keyNames(primaryColumnsCount - 1) As String
                'Dim childKeys(primaryColumnsCount - 1) As DataColumn
                For i As Integer = 0 To primaryColumnsCount - 1
                    keys(i) = dataDS1.Tables("tobeupdated").Columns(LV1.DataKeyNames(i))


                    keyNames(i) = LV1.DataKeyNames(i)
                    If dataDS1.Tables("tobeupdated").Columns(keyNames(i)).DataType = GetType(String) Then 'placeholder for key fields
                        filterSQL &= keyNames(i) & " = '@" & i & "' And "
                    Else
                        filterSQL &= keyNames(i) & " = @" & i & " And "
                    End If
                Next
                dataDS1.Tables("tobeupdated").PrimaryKey = keys
                filterSQL = filterSQL.Remove(filterSQL.Length - 4, 4)



                'check if there are duplicate records in the update file
                If continue1 Then

                    Dim sortedColumns As String = String.Empty
                    For i As Integer = 0 To primaryColumnsCount - 1
                        sortedColumns &= keyNames(i) & " ASC,"
                    Next
                    sortedColumns = sortedColumns.Remove(sortedColumns.Length - 1, 1)

                    'firstly sort the data for further process
                    Dim sortedRow() As DataRow = dataDS1.Tables("update").Select(Nothing, sortedColumns)

                    'to see if there two nearby rows have the same value for all the key columns
                    Dim completeEqual As Boolean

                    For i As Integer = 1 To sortedRow.Count - 1
                        completeEqual = True
                        For j As Integer = 0 To primaryColumnsCount - 1
                            If sortedRow(i).Item(keyNames(j)) = sortedRow(i - 1).Item(keyNames(j)) Then
                            Else
                                completeEqual = False
                            End If
                        Next

                        If completeEqual Then
                            Dim rowInformation As String = String.Empty
                            For k As Integer = 0 To primaryColumnsCount - 1
                                rowInformation &= sortedRow(i).Item(keyNames(k)) & ","
                            Next
                            rowInformation = rowInformation.Substring(0, rowInformation.Length - 1)
                            continue1 = False
                            'msgPopUP("Duplicate records exist in your  excel file: " & rowInformation, StatusLabel)
                            msgRtrn.AppendLine("<div style='color:red;'>" & "Duplicate records exist in your  excel file: " & rowInformation & "</div>")
                            Exit For
                        End If

                    Next

                End If





                'mark row status as delete if there is corresponding record found in table "tobeupdated"
                If continue1 Then


                    howManyRecordsDeleted = dataDS1.Tables("tobeupdated").Rows.Count

                    For i As Integer = 0 To howManyRecordsDeleted - 1
                        dataDS1.Tables("tobeupdated").Rows(i).Delete()
                    Next


                End If


            Catch ex As Exception
                continue1 = False
                'msgPopUP("Database maybe do not have the records you are asking to delete.", StatusLabel)
                msgRtrn.AppendLine("<div style='color:red;'>" & "There might be illegal data in your Excel file." & ex.Message & "</div>")
            End Try

        End If

        If continue1 Then
            Try
                'delete records
                'dtAdapter1.UpdateBatchSize = 512
                dtAdapter1.Update(dataDS1, "tobeupdated")
                'msgPopUP("Action on delection upon excel file is completed. The number of deleted records is " & howManyRecordsDeleted, StatusLabel, False, False)
                msgRtrn.AppendLine("Action on delection upon excel file is completed. The number of deleted records is " & howManyRecordsDeleted)
                continueOrnot = True
            Catch ex As Exception
                continueOrnot = False
                msgRtrn.AppendLine("<div style='color:red;'>" & "Something wrong during the first step of  deletion and insertion." & ex.Message & "</div>")
            End Try


        End If

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

        Return msgRtrn.ToString

    End Function




    ''' <summary>
    ''' before inserting,updating or deleting data on dataTable dtTbl, the data validity has to be checked
    ''' </summary>
    ''' <param name="dtTbl">a specific dataTable to be checked</param>
    ''' <returns>return error or warning messages after data checking</returns>
    Public Overridable Function dataValidityCheck(ByRef dtTbl As DataTable) As String

        Return String.Empty

    End Function

    ''' <summary>
    ''' get a full list of column name in table Esch_Na_tbl_orders  in order to do some checking on data validity
    ''' </summary>
    Protected Function getColumnNameOf_Esch_Na_tbl_orders() As String

        Dim connstr As String = ConfigurationManager.ConnectionStrings(dbConnectionName).ConnectionString
        Dim conn As SqlConnection = New SqlConnection(connstr)
        Dim dtUpdateFrom0 As SqlDataAdapter = New SqlDataAdapter("SELECT * FROM Esch_Na_tbl_orders WHERE txt_order_key = 'Null'", conn)
        Dim dtTable As DataTable = New DataTable()
        dtUpdateFrom0.Fill(dtTable)
        Dim columnNames As StringBuilder = New StringBuilder(",")
        For Each a As DataColumn In dtTable.Columns
            columnNames.Append(a.ColumnName & ",")
        Next

        dtTable.Dispose()
        dtUpdateFrom0.Dispose()
        conn.Dispose()

        Return columnNames.ToString()

    End Function


    ''' <summary>
    ''' put all initial works together
    ''' </summary>
    ''' <param name="SDS1"></param>
    ''' <param name="DDL1"></param>
    ''' <param name="DDL2"></param>
    ''' <param name="filtercdtn1"></param>
    ''' <param name="Filter1"></param>
    ''' <param name="Download1"></param>
    ''' <param name="hiddenBT"></param>
    ''' <remarks></remarks>
    Public Overridable Sub pageLoadInitiate(ByRef SDS1 As SqlDataSource, ByRef DDL1 As WebControls.DropDownList, ByRef DDL2 As WebControls.DropDownList, ByRef filtercdtn1 As WebControls.TextBox, ByRef Filter1 As WebControls.Button, ByRef Download1 As WebControls.Button, ByRef hiddenBT As WebControls.HiddenField)

        Dim args As DataSourceSelectArguments = New DataSourceSelectArguments()
        Dim view1 As DataView = CType(SDS1.Select(args), DataView)
        Dim table1 As DataTable = view1.ToTable()


        'decide how many columns and rows if data in table is put into Excel file
        rightmostColumn = integerToExcelColumnName(table1.Columns.Count - 1)


        If Not Page.IsPostBack Then
            'when first time page loading, initiate filter related objects,fill the field with data extracted from relevant table

            initiateFilterControls(DDL1, DDL2, table1)

        Else
            If IsNothing(CacheFrom("fltr" & pageTitleReference())) Then
                CacheInsert("fltr" & pageTitleReference(), " ")
            End If

            SDS1.FilterExpression = CType(CacheFrom("fltr" & pageTitleReference()), String)

        End If

        'press enter key in textbox to initate one click on another button
        filtercdtn1.Attributes.Add("onkeydown", "if(event.which || event.keyCode){if ((event.which == 13) || (event.keyCode == 13)) {document.getElementById('" & Filter1.ClientID & "').click();return false;}} else {return true}; ")
        Download1.Attributes.Add("onclick", "var offset = new Date().getTimezoneOffset();" & hiddenBT.ClientID & ".value = offset;return true;")

    End Sub




    ''' <summary>
    ''' initiation for some dropdownlist controls
    ''' </summary>
    ''' <param name="DDL1"></param>
    ''' <param name="DDL2"></param>
    ''' <param name="table1"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Overridable Function initiateFilterControls(ByRef DDL1 As WebControls.DropDownList, ByRef DDL2 As WebControls.DropDownList, ByRef table1 As DataTable) As String

        CacheInsert("fltr" & pageTitleReference(), " ")

        Dim listdt As DataTable = New DataTable()
        ' Define the columns of the table.
        listdt.Columns.Add(New DataColumn("fieldname", GetType(String)))
        listdt.Columns.Add(New DataColumn("fieldtype", GetType(String)))

        Dim dr As DataRow, i As Integer = 0
        For Each clmn As DataColumn In table1.Columns
            dr = listdt.NewRow()
            dr("fieldname") = clmn.ColumnName
            dr("fieldtype") = clmn.DataType.ToString() & "," & i.ToString()

            listdt.Rows.Add(dr)
            i += 1
        Next

        Dim dv As DataView = New DataView(listdt)

        DDL1.DataSource = dv
        DDL1.DataTextField = "fieldname"
        DDL1.DataValueField = "fieldtype"
        DDL1.DataBind()

        DDLchangeSelection(DDL1, DDL2)

        Return String.Empty
    End Function

    ''' <summary>
    ''' how to proceed after click button
    ''' </summary>
    ''' <param name="filtercdtn1"></param>
    ''' <param name="DDL1"></param>
    ''' <param name="DDL2"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Overridable Function filterClickHnadler(ByRef filtercdtn1 As WebControls.TextBox, ByRef DDL1 As WebControls.DropDownList, ByRef DDL2 As WebControls.DropDownList, ByRef SDS1 As SqlDataSource, ByRef LV1 As WebControls.ListView) As String

        Dim filterText As String = filtercdtn1.Text.Trim()

        Dim userID_filter As String = "fltr" & pageTitleReference()

        Select Case DDL1.SelectedValue.ToLower().Split(",".ToCharArray)(0)

            Case "system.string"
                Select Case DDL2.SelectedItem.ToString()
                    Case "="
                        CacheInsert(userID_filter, DDL1.SelectedItem.ToString() & DDL2.SelectedValue & " '" & filterText & "'")
                    Case Else
                        CacheInsert(userID_filter, DDL1.SelectedItem.ToString() & DDL2.SelectedValue & " '%" & filterText & "%'")
                End Select

            Case "system.int32", "system.int16", "system.int64", "system.integer", "system.int64", "system.single", "system.double", "system.decimal"
                If IsNumeric(filterText) Then
                    CacheInsert(userID_filter, DDL1.SelectedItem.ToString() & DDL2.SelectedValue & filterText)
                Else
                    CacheInsert(userID_filter, " ")
                    filtercdtn1.Text = filterText & "==>invalid number"
                End If

            Case "system.datetime"
                If IsDate(filterText) Then
                    CacheInsert(userID_filter, DDL1.SelectedItem.ToString() & DDL2.SelectedValue & " " & dateSeparator & CDate(filterText) & dateSeparator & "")

                Else
                    CacheInsert(userID_filter, " ")
                    filtercdtn1.Text = filterText & "==>invalid date"
                End If

            Case "system.boolean"
                Try
                    Dim flag As Boolean = Boolean.Parse(filterText)
                    Select Case DDL2.SelectedItem.ToString()
                        Case "="
                            CacheInsert(userID_filter, DDL1.SelectedItem.ToString() & DDL2.SelectedValue & flag)
                        Case "<>"
                            CacheInsert(userID_filter, DDL1.SelectedItem.ToString() & DDL2.SelectedValue & flag)
                    End Select

                Catch ex As Exception
                    CacheInsert(userID_filter, " ")
                    filtercdtn1.Text = filterText & "==>illegal input"
                End Try

            Case Else
                CacheInsert(userID_filter, " ")
        End Select


        SDS1.FilterExpression = CType(CacheFrom(userID_filter), String)
        LV1.DataBind()


        Return String.Empty

    End Function

    ''' <summary>
    ''' 
    ''' </summary>
    ''' <param name="SDS1"></param>
    ''' <param name="LV1"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Overridable Function clrfltrClickHnadler(ByRef SDS1 As SqlDataSource, ByRef LV1 As WebControls.ListView) As String
        CacheInsert("fltr" & pageTitleReference(), " ")
        SDS1.FilterExpression = " "

        LV1.DataBind()

        Return String.Empty
    End Function


    Public Overridable Function UpldUpdate_ClickHnadler(ByRef SDS1 As SqlDataSource, ByRef LV1 As WebControls.ListView, ByRef StatusLabel As WebControls.Label, ByRef FileUpload1 As WebControls.FileUpload) As String
        Dim msg As String = String.Empty


        Dim filepath As String = String.Empty, filename As String = String.Empty, flextsn As String = String.Empty

        If fileToServer(filepath, filename, flextsn, FileUpload1, msg) Then

            msg &= dataUpdatedToDatabaseTablePerExcelData(StatusLabel, SDS1, LV1, filepath, filename, flextsn)

            clrfltrClickHnadler(SDS1, LV1)  'clear the filter for LV1 to show all the data

        Else
            msg &= "<div style='color:red;'>No file was selected.</div>"
        End If


        msgPopUP(msg, StatusLabel, False, False)

        Return String.Empty
    End Function


    Public Overridable Function UpldDel_ClickHnadler(ByRef SDS1 As SqlDataSource, ByRef LV1 As WebControls.ListView, ByRef StatusLabel As WebControls.Label, ByRef FileUpload1 As WebControls.FileUpload) As String
        Dim msg As String = String.Empty


        Dim filepath As String = String.Empty, filename As String = String.Empty, flextsn As String = String.Empty

        Dim continue1 As Boolean = False

        If fileToServer(filepath, filename, flextsn, FileUpload1, msg) Then

            msg &= dataDeletedFromDatabaseTablePerExcelData(StatusLabel, SDS1, LV1, filepath, filename, flextsn, continue1)

            clrfltrClickHnadler(SDS1, LV1)  'clear the filter for LV1 to show all the data

        Else
            msg &= "<div style='color:red;'>No file was selected.</div>"
        End If


        msgPopUP(msg, StatusLabel, False, False)

        Return String.Empty
    End Function



    Public Overridable Function UpldInsrt_ClickHnadler(ByRef SDS1 As SqlDataSource, ByRef LV1 As WebControls.ListView, ByRef StatusLabel As WebControls.Label, ByRef FileUpload1 As WebControls.FileUpload) As String
        Dim msg As String = String.Empty


        Dim filepath As String = String.Empty, filename As String = String.Empty, flextsn As String = String.Empty

        If fileToServer(filepath, filename, flextsn, FileUpload1, msg) Then

            msg &= dataInsertionToDatabaseTableFromExcel(StatusLabel, SDS1, LV1, filepath, filename, flextsn)

            clrfltrClickHnadler(SDS1, LV1)  'clear the filter for LV1 to show all the data

        Else
            msg &= "<div style='color:red;'>No file was selected.</div>"
        End If


        msgPopUP(msg, StatusLabel, False, False)

        Return String.Empty
    End Function



    Public Overridable Function overwrite_ClickHnadler(ByRef SDS1 As SqlDataSource, ByRef LV1 As WebControls.ListView, ByRef StatusLabel As WebControls.Label, ByRef FileUpload1 As WebControls.FileUpload) As String
        Dim msg As String = String.Empty


        Dim filepath As String = String.Empty, filename As String = String.Empty, flextsn As String = String.Empty

        Dim continue1 As Boolean = False

        If fileToServer(filepath, filename, flextsn, FileUpload1, msg) Then

            msg &= dataDeletedFromDatabaseTablePerExcelData(StatusLabel, SDS1, LV1, filepath, filename, flextsn, continue1) & "<br />"
            If continue1 Then
                msg &= dataInsertionToDatabaseTableFromExcel(StatusLabel, SDS1, LV1, filepath, filename, flextsn, continue1)
            End If

            clrfltrClickHnadler(SDS1, LV1)  'clear the filter for LV1 to show all the data

        Else
            msg &= "<div style='color:red;'>No file was selected.</div>"
        End If


        msgPopUP(msg, StatusLabel, False, False)

        Return String.Empty
    End Function



    ''' <summary>
    ''' After change the selection in the first DDL, the collection will change accordingly
    ''' </summary>
    ''' <param name="DDL1"></param>
    ''' <param name="DDL2"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Overridable Function DDLchangeSelection(ByRef DDL1 As WebControls.DropDownList, ByRef DDL2 As WebControls.DropDownList) As String


        Dim operatordt As DataTable = New DataTable()
        operatordt.Columns.Add(New DataColumn("name", GetType(String)))
        operatordt.Columns.Add(New DataColumn("value", GetType(String)))
        Dim opdr As DataRow, opdv As DataView

        Select Case DDL1.SelectedValue.ToLower().Split(",".ToCharArray)(0)
            Case "system.string"
                opdr = operatordt.NewRow
                opdr("name") = "contains"
                opdr("value") = " LIKE "
                operatordt.Rows.Add(opdr)
                opdr = operatordt.NewRow
                opdr("name") = "="
                opdr("value") = " = "
                operatordt.Rows.Add(opdr)
                opdv = New DataView(operatordt)

            Case "system.datetime"
                opdr = operatordt.NewRow
                opdr("name") = "="
                opdr("value") = " = "
                operatordt.Rows.Add(opdr)
                opdr = operatordt.NewRow
                opdr("name") = "after"
                opdr("value") = " >= "
                operatordt.Rows.Add(opdr)
                opdr = operatordt.NewRow
                opdr("name") = "before"
                opdr("value") = " <= "
                operatordt.Rows.Add(opdr)
                opdv = New DataView(operatordt)

            Case "system.boolean"
                opdr = operatordt.NewRow
                opdr("name") = "="
                opdr("value") = " = "
                operatordt.Rows.Add(opdr)
                opdr = operatordt.NewRow
                opdr("name") = "<>"
                opdr("value") = " <> "
                operatordt.Rows.Add(opdr)
                opdv = New DataView(operatordt)

            Case "system.boolean"
                opdr = operatordt.NewRow
                opdr("name") = "="
                opdr("value") = " = "
                operatordt.Rows.Add(opdr)
                opdr = operatordt.NewRow
                opdr("name") = "<>"
                opdr("value") = " <> "
                operatordt.Rows.Add(opdr)
                opdv = New DataView(operatordt)

            Case "system.int32", "system.int16", "system.int64", "system.integer", "system.int64", "system.single", "system.double", "system.decimal"
                opdr = operatordt.NewRow
                opdr("name") = "="
                opdr("value") = " = "
                operatordt.Rows.Add(opdr)
                opdr = operatordt.NewRow
                opdr("name") = ">="
                opdr("value") = " >= "
                operatordt.Rows.Add(opdr)
                opdr = operatordt.NewRow
                opdr("name") = "<="
                opdr("value") = " <= "
                operatordt.Rows.Add(opdr)
                opdv = New DataView(operatordt)

            Case Else
                opdr = operatordt.NewRow
                opdr("name") = "unknown?"
                opdr("value") = "unknown?"
                operatordt.Rows.Add(opdr)
                opdv = New DataView(operatordt)
        End Select


        DDL2.DataSource = opdv
        DDL2.DataTextField = "name"
        DDL2.DataValueField = "value"
        DDL2.DataBind()


        Return String.Empty
    End Function



End Class



