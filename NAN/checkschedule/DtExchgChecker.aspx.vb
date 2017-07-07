Imports System
Imports System.Configuration
Imports System.Data
Imports System.Data.SqlClient
Imports Microsoft.VisualBasic
Imports System.Runtime.InteropServices
Imports System.Diagnostics
Imports System.IO
Imports System.Data.OleDb



Partial Class checkschedule_DtExchgChecker
    Inherits basepage1

    Public Const excelDatabaseFile = "checker"
    Private Const Separator1 = "#"

    Protected Sub Page_Init1(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Init
        accesslevel = 2 ' any one can access this page
    End Sub


    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Response.Clear()
        Dim query1 As String, connstr As String, actionRequested As String
        Dim startD As Date = CDate(Today)
        Dim daysFromStart As Integer = 30
        Dim pixelsPerDay As Long = 120
        Dim sndbck As System.Text.StringBuilder = New System.Text.StringBuilder()

        'Considering special case in Nansha, its sorting is by dat_finish_date instead of dat_start_date like Shanghai and Chongqing
        Dim DAT_START_DATE As String = "dat_start_date", DAT_FINISH_DATE As String = "dat_finish_date"
        If valueOf("strGanttStyle").ToUpper.StartsWith("SORTBYFINISHTIME") Then
            DAT_START_DATE = "dat_finish_date"
            DAT_FINISH_DATE = "dat_start_date"
        End If

        actionRequested = Request.Params("action")




        Dim temp As New NameValueCollection

        If CacheFrom("checker") Is Nothing Then 'need refresh data for planning period parameter

            Dim connParam As SqlConnection = New SqlConnection(ConfigurationManager.ConnectionStrings(dbConnForParam).ConnectionString)
            connParam.Open()


            'get the lastest data from table Esch_Na_tbl_plnprmtr to preset some planning parameter

            Dim command1 As New SqlCommand("Select paramname, paramvalue From Esch_Na_tbl_plnprmtr where category = 'checker'", connParam)
            Dim reader1 As SqlDataReader
            reader1 = command1.ExecuteReader()
            Do While reader1.Read()
                temp.Add(reader1("paramname").ToString, reader1("paramvalue").ToString)
            Loop

            reader1.Close()

            command1.Dispose()

            connParam.Close()
            connParam.Dispose()


            CacheInsert("checker", temp)

        End If


        temp = CType(CacheFrom("checker"), NameValueCollection)

        If temp("autoStart").ToLower = "no" Then
            startD = CDate(temp("starttime"))
        End If
        daysFromStart = CInt(temp("daystobeshown"))





        'Get production lines' information ===
        If actionRequested = "productionlines" Then

            Dim arrayLine1 As Integer() = arrayOfLines()

            Dim i As Integer = arrayLine1.Count
            'ReDim Preserve arrayLine1(i)
            'arrayLine1(i) = CInt(valueOf("intDummyLine"))
            'i += 1
            Dim j1 As Integer = 0


            sndbck.Append("<table border='1' ><tr>")
            For j1 = 0 To (i - 1)
                sndbck.Append("<td onclick='PX.dsplyOrderByLine(this);' class='linemenu' id='" & arrayLine1(j1) & "'>" & "L" & arrayLine1(j1) & "</td>")
            Next
            sndbck.Append("</tr></table>")


        End If

        'provide content for scalar bar [   ]
        If actionRequested = "scalar" Then
            For j1 = 1 To daysFromStart
                sndbck.Append("<div class='sc' style='text-align:center;width:" & (pixelsPerDay - 1) & "px;left:" & (pixelsPerDay * (j1 - 1)) & "px;" & IIf(j1 Mod 2 = 1, "background-color:#00FFFF;", "") & "'>" & DateAdd("d", j1 - 1, startD).ToString("ddd  MMM d") & "</div>")
            Next
        End If

        'provide background for background mark lines ------------------
        If actionRequested = "bckgrndImg" Then
            For j1 = 1 To daysFromStart
                sndbck.Append("<div class='bg' style='width:" & (pixelsPerDay - 1) & "px;left:" & (pixelsPerDay * (j1 - 1)) & "px;'></div>")
            Next
        End If





        'generate excel file as the database to show gantt chart
        If actionRequested = "generateExcelFile" Then
            connstr = ConfigurationManager.ConnectionStrings(dbConnectionName).ConnectionString
            Dim conn As SqlConnection = New SqlConnection(connstr)


            query1 = "Select table1.txt_order_key,int_line_no,dat_start_date,dat_finish_date,txt_lot_no,int_span,planned_production_qty,flt_unallocate_qty,flt_order_qty,txt_color,txt_item_no,dat_etd,txt_VIP,txt_process_technics,txt_FDA,txt_remark,txt_payment_status,table2.txt_pull_screw From  " & _
             " ((Select txt_order_key,int_line_no,dat_start_date,dat_finish_date,txt_lot_no,int_span,planned_production_qty,flt_unallocate_qty,flt_order_qty,txt_color,txt_item_no,dat_etd,txt_VIP,txt_process_technics,txt_FDA,txt_remark,txt_payment_status From Esch_Na_tbl_orders  Where CAST(int_line_no AS VARCHAR(5)) <> '" & valueOf("intDummyLine") & "' And  ((dat_start_date  between " & dateSeparator & startD & dateSeparator & " And " & dateSeparator & DateAdd("d", daysFromStart, startD) & dateSeparator & ") Or (dat_finish_date  between " & dateSeparator & startD & dateSeparator & " And " & dateSeparator & DateAdd("d", daysFromStart, startD) & dateSeparator & "))) as table1 Left Join (select txt_order_key,txt_pull_screw From Esch_Na_tbl_prductn_prmtr) as table2 On (table1.txt_order_key = table2.txt_order_key)) Order by int_line_no, " & DAT_START_DATE & "  ASC, " & DAT_FINISH_DATE & " ASC"

            Dim dtUpdateFrom As SqlDataAdapter = New SqlDataAdapter(query1, conn)
            Dim dtFromTable As New DataTable()
            dtUpdateFrom.Fill(dtFromTable)

            Dim excelFileName As String = generateExcel(Nothing, dtFromTable, ConfigurationManager.AppSettings("excelFolder"))

            dtFromTable.Dispose()
            dtUpdateFrom.Dispose()


        End If



        If actionRequested = "orderlines" Then

            Dim excelconnectionstr As String
            Dim continueOrNot = True

            Dim filePath As String = ConfigurationManager.AppSettings("excelFolder")
            If Not filePath.EndsWith("\") Then
                filePath &= "\"
            End If

            If File.Exists(filePath & excelDatabaseFile & ".xlsx") Then
                excelconnectionstr = String.Format("provider=Microsoft.ACE.OleDb.12.0; Data Source='{0}';" & "Extended Properties='Excel 12.0 Xml;HDR=YES;IMEX=1'", filePath & excelDatabaseFile & ".xlsx")
            Else
                If File.Exists(filePath & excelDatabaseFile & ".xls") Then
                    excelconnectionstr = String.Format("provider=Microsoft.Jet.OleDb.4.0; Data Source='{0}';" & "Extended Properties='Excel 8.0;HDR=YES;IMEX=1'", filePath & excelDatabaseFile & ".xls")
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

                Dim reader As OleDbDataReader
                reader = command1.ExecuteReader()

                Dim interMinutes As Long = 0
                Dim percentageOfCompletion As String = String.Empty
                Dim additionGap As Integer = 0 'add extra space when there is percentage of completion
                Dim orderSatus As String = String.Empty, leftPosForRSDspan As String
                Dim screwType As String = String.Empty, FDAremark As String = String.Empty, VIPremark As String = String.Empty
                Dim currentOrderkey As String = String.Empty
                Dim blockBorderStyle As String = String.Empty
                Dim pullScrew As String = String.Empty
                Dim currency1 As String = String.Empty
                Dim currentLotNo As String = String.Empty, previousLotNo As String = String.Empty
                Dim listOfLotNo As New List(Of String)
                Dim SPANandETD As String = String.Empty
                Dim RDDqtyNewOffset As Integer = 0
                Dim marginTop As String = String.Empty
                Dim offsetToLeftDueToScrewForCompletionPercent As String = "3px"
                Dim numOfAllowableChanges As String = String.Empty

                Do While reader.Read()

                    currentLotNo = reader("txt_lot_no").ToString

                    SPANandETD = reader("int_span") & " (" & reader("dat_etd") & ")"

                    'orderSatus = "<span style='left:-" & (9 + offsetForChrome) & "em;position:relative;color:red;'>new</span>"
                    leftPosForRSDspan = "left:-" & (SPANandETD.Length + RDDqtyNewOffset) * 7 & "px;"

                    currentOrderkey = reader("txt_order_key")

                    blockBorderStyle = "border:1px solid black;"

                    interMinutes = CLng(DateDiff(DateInterval.Minute, reader("dat_start_date"), reader("dat_finish_date")) * pixelsPerDay / (24 * 60))

                    If DBNull.Value.Equals(reader("txt_process_technics")) Then
                        screwType = String.Empty
                    Else
                        screwType = "&nbsp;&nbsp;(screw:" & reader("txt_process_technics") & ")&nbsp;&nbsp;"
                    End If

                    If Not DBNull.Value.Equals(reader("txt_pull_screw")) Then
                        additionGap += 2 * 6
                        If String.Equals(reader("txt_pull_screw"), "yellow", StringComparison.OrdinalIgnoreCase) Then
                            pullScrew = "<span class = 'arrow-right' style='left:1px;border-left:10px solid #FFD700;position:absolute;'></span>"
                        Else
                            pullScrew = "<span class = 'arrow-right' style='left:1px;border-left:10px solid red;position:absolute;'></span>"
                        End If

                        'marginTop = "margin-top:-0.25em;"
                        offsetToLeftDueToScrewForCompletionPercent = "left:15px"
                    Else
                        pullScrew = String.Empty
                        'marginTop = String.Empty
                        offsetToLeftDueToScrewForCompletionPercent = "left:3px"
                    End If

                    sndbck.Append("<div  class='g-" & reader("txt_color").ToString() & "' style='height:16px;position:relative;" & blockBorderStyle & "margin:5px 0px 5px 0px;left:" & CLng(DateDiff("n", startD, reader("dat_start_date")) * pixelsPerDay / (24 * 60)) & "px;width:" & (interMinutes - 2) & "px;' id='" & reader("txt_order_key") & "'><span style='" & leftPosForRSDspan & "position:relative;'>" & SPANandETD & currency1 & "&nbsp;" & orderSatus & "</span>" & "<span class='sw' style='left:" & interMinutes - 2 & "px;" & marginTop & "'>" & pullScrew & "<span style='" & offsetToLeftDueToScrewForCompletionPercent & ";'>" & percentageOfCompletion & "</span></span>" &
                                  "<span style='position:absolute;left:" & (interMinutes + additionGap) & "px'>" & "&nbsp;&nbsp;&nbsp;&nbsp;" & currentLotNo & "&nbsp;&nbsp;&nbsp;" & reader("txt_item_no") & "&nbsp;" & reader("planned_production_qty") & "KG " & screwType & reader("txt_FDA") & "&nbsp;<span style='color:red;'>" & reader("txt_VIP") & "</span>&nbsp;&nbsp;(" & currentOrderkey & ") " & reader("txt_payment_status") & "&nbsp;<span style='color:red;'>" & reader("txt_remark") & "</span></span></div>")

                Loop

                reader.Close()

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
    ''' get data from dataTable and a excel is created to store the data, and generated excel is expected for downloading
    ''' </summary>
    ''' <param name="StatusLabel">to show the warning message or others</param>
    ''' <returns>return a excel file name</returns>
    Public Overloads Function generateExcel(ByRef StatusLabel As Label, ByRef dtTable As DataTable, ByVal filePath As String) As String



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

            xlBook.SaveCopyAs(filePath & "\" & excelfilename & xclFileExtsn)


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




        Return excelfilename & xclFileExtsn

    End Function


End Class
