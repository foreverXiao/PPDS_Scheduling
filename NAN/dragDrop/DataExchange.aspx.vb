Imports System
Imports System.Configuration
Imports System.Data
Imports System.Data.OleDb
Imports Microsoft.VisualBasic



Partial Class dragDrop_DataExchange
    Inherits FrequentPlanActions


    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Response.Clear()

        Dim query1 As String, connstr As String, actionRequested As String
        Dim startD As Date = CDate(Today.AddDays(-1))
        Dim daysFromStart As Integer = 90
        Dim pixelsPerDay As Long = 90


        Dim sndbck As System.Text.StringBuilder = New System.Text.StringBuilder()
        

        actionRequested = Request.Params("action")

        Dim temp As New NameValueCollection
        
        If CacheFrom("schedule") Is Nothing Then 'need refresh data for planning period parameter

            Dim connParam As OleDbConnection = New OleDbConnection(ConfigurationManager.ConnectionStrings(dbConnForParam).ProviderName & ConfigurationManager.ConnectionStrings(dbConnForParam).ConnectionString)
            connParam.Open()


            'get the lastest data from table Esch_Na_tbl_plnprmtr to preset some planning parameter

            Dim command1 As New OleDbCommand("Select paramname, paramvalue From Esch_Na_tbl_plnprmtr where category = 'schedule'", connParam)
            Dim reader1 As OleDbDataReader
            reader1 = command1.ExecuteReader()
            Do While reader1.Read()
                temp.Add(reader1("paramname").ToString, reader1("paramvalue").ToString)
            Loop

            reader1.Close()

            command1.Dispose()

            connParam.Close()
            connParam.Dispose()


            CacheInsert("schedule", temp)

        End If


        temp = CType(CacheFrom("schedule"), NameValueCollection)

        If temp("autoStart").ToLower = "no" Then
            startD = CDate(temp("starttime"))
        End If
        daysFromStart = CInt(temp("daystobeshown"))


        Dim userName As String = lockKeyTable(priority.ganttChart) 'no other user operating on the key table
        Dim autOp As Boolean = lineListOwnedByUser.Contains("'" & Request.Params("lineno") & "'") AndAlso String.IsNullOrEmpty(userName)


        connstr = ConfigurationManager.ConnectionStrings(dbConnectionName).ConnectionString
        Dim conn As OleDbConnection = New OleDbConnection(connstr)
        conn.Open()

        Dim dtTable As New DataTable
        Dim command As New OleDbCommand(Nothing, conn)
        Dim reader As OleDbDataReader

        'Considering special case in Nansha, its sorting is by dat_finish_date instead of dat_start_date like Shanghai and Chongqing
        Dim DAT_START_DATE As String = "dat_start_date", DAT_FINISH_DATE As String = "dat_finish_date", ShowRMBcurrency As Boolean = False
        If valueOf("strGanttStyle").ToUpper.StartsWith("SORTBYFINISHTIME") Then
            DAT_START_DATE = "dat_finish_date"
            DAT_FINISH_DATE = "dat_start_date"
            ShowRMBcurrency = True
        End If


        If autOp Then 'the user has the right to do planning on this production line
            'adjust order one by one ======
            If actionRequested = "updateorder" Then
                query1 = "Select * From Esch_Na_tbl_orders Where txt_order_key  ='" & Request.Params("orderkey") & "'"
                Dim dtAdapter1 As OleDbDataAdapter = New OleDbDataAdapter(query1, conn)
                Dim cmdbAccessCmdBuilder As New OleDbCommandBuilder(dtAdapter1)
                dtAdapter1.UpdateCommand = cmdbAccessCmdBuilder.GetUpdateCommand()

                dtAdapter1.Fill(dtTable)

                Dim keys(0) As DataColumn
                keys(0) = dtTable.Columns("txt_order_key")
                dtTable.PrimaryKey = keys

                Dim cRow As DataRow = dtTable.Rows(0)
				'FANAR+ GO-LIVE, on Nov.9.2016  need calculate start time based on finished time, backward calculation
				Dim timedifference1 As Long
				timedifference1 = DateDiff(DateInterval.Minute, cRow("dat_start_date"), cRow("dat_finish_date"))
                cRow("dat_start_date") = DateAdd(DateInterval.Minute, CLng(Request.Params("timepixels")) * 24 * 60 / pixelsPerDay, startD)
				'FANAR+ GO-LIVE, on Nov.9.2016  need calculate start time based on finished time, backward calculation
				cRow("dat_finish_date") = DateAdd(DateInterval.Minute, timedifference1, cRow("dat_start_date"))


                Dim hasException As Boolean = False
                finishTime_exPlantDate_Span1(conn, dtTable.Select(Nothing), hasException)
                If hasException Then
                    'exceptionR.Visible = True
                    CacheInsert("Oexception", 1)
                End If

                additionDaysOnExplantDate(conn, dtTable, DataViewRowState.ModifiedCurrent)


                dtAdapter1.Update(dtTable)
                cmdbAccessCmdBuilder.Dispose()
                dtAdapter1.Dispose()

                'to record the last group of order keys you changed in order to mark them with different dot line 
                CacheInsert("slctKeys", Request.Params("orderkey") & "")


                actionRequested = "orderlines" ' continue next step to display all order lines

            End If


            'change production line for an order ======
            If actionRequested = "changeline" Then
                If autOp Then 'the user has the right to do planning on this production line
                    query1 = "Select * From Esch_Na_tbl_orders Where txt_order_key  ='" & Request.Params("orderkey") & "'"
                    Dim dtAdapter1 As OleDbDataAdapter = New OleDbDataAdapter(query1, conn)
                    Dim cmdbAccessCmdBuilder As New OleDbCommandBuilder(dtAdapter1)
                    dtAdapter1.UpdateCommand = cmdbAccessCmdBuilder.GetUpdateCommand()

                    dtAdapter1.Fill(dtTable)

                    Dim keys(0) As DataColumn
                    keys(0) = dtTable.Columns("txt_order_key")
                    dtTable.PrimaryKey = keys

                    Dim cRow As DataRow = dtTable.Rows(0)
                    cRow("int_line_no") = Request.Params("newlineno")


                    Dim hasException As Boolean = False
                    finishTime_exPlantDate_Span1(conn, dtTable.Select(Nothing), hasException)
                    If hasException Then
                        'exceptionR.Visible = True
                        CacheInsert("Oexception", 1)
                    End If

                    additionDaysOnExplantDate(conn, dtTable, DataViewRowState.ModifiedCurrent)


                    dtAdapter1.Update(dtTable)
                    cmdbAccessCmdBuilder.Dispose()
                    dtAdapter1.Dispose()

                    'to record the last group of order keys you changed in order to mark them with different dot line 
                    CacheInsert("slctKeys", Request.Params("orderkey") & "")


                End If

                actionRequested = "orderlines" ' continue next step to display all order lines

            End If



            'change start time for one order ======
            If actionRequested = "newStartTime" Then
                query1 = "Select * From Esch_Na_tbl_orders Where txt_order_key  ='" & Request.Params("orderkey") & "'"
                Dim dtAdapter1 As OleDbDataAdapter = New OleDbDataAdapter(query1, conn)
                Dim cmdbAccessCmdBuilder As New OleDbCommandBuilder(dtAdapter1)
                dtAdapter1.UpdateCommand = cmdbAccessCmdBuilder.GetUpdateCommand()

                dtAdapter1.Fill(dtTable)

                Dim keys(0) As DataColumn
                keys(0) = dtTable.Columns("txt_order_key")
                dtTable.PrimaryKey = keys

                Dim cRow As DataRow = dtTable.Rows(0)
				
				'FANAR+ GO-LIVE, on Nov.9.2016  need calculate start time based on finished time, backward calculation
				Dim timedifference1 As Long
				timedifference1 = DateDiff(DateInterval.Minute, cRow("dat_start_date"), cRow("dat_finish_date"))
                				
				
                If IsDate(Request.Params("newTime")) Then
                    cRow("dat_start_date") = CDate(Request.Params("newTime"))
                End If

				'FANAR+ GO-LIVE, on Nov.9.2016  need calculate start time based on finished time, backward calculation
				cRow("dat_finish_date") = DateAdd(DateInterval.Minute, timedifference1, cRow("dat_start_date"))
				
                Dim hasException As Boolean = False
                finishTime_exPlantDate_Span1(conn, dtTable.Select(Nothing), hasException)
                If hasException Then
                    'exceptionR.Visible = True
                    CacheInsert("Oexception", 1)
                End If

                additionDaysOnExplantDate(conn, dtTable, DataViewRowState.ModifiedCurrent)


                dtAdapter1.Update(dtTable)
                cmdbAccessCmdBuilder.Dispose()
                dtAdapter1.Dispose()

                'to record the last group of order keys you changed in order to mark them with different dot line 
                CacheInsert("slctKeys", Request.Params("orderkey") & "")


                actionRequested = "orderlines" ' continue next step to display all order lines

            End If



            'adjust some orders in one batch, link orders and push forward======
            If actionRequested = "updateorderinbatch" Then
                Dim minDatetime As DateTime, maxDatetime As DateTime
                query1 = "Select dat_start_date,dat_finish_date,txt_order_key From Esch_Na_tbl_orders  Where txt_order_key in ('" & Request.Params("frstslctnID") & "','" & Request.Params("scndslctnID") & "') Order by dat_start_date ASC,dat_finish_date ASC,txt_order_key ASC"
                command.CommandText = query1
                reader = command.ExecuteReader()
                If reader.Read() Then minDatetime = reader(DAT_START_DATE)
                If reader.Read() Then maxDatetime = reader(DAT_START_DATE)
                reader.Close()

                query1 = "Select * From Esch_Na_tbl_orders Where (" & DAT_START_DATE & " between " & dateSeparator & minDatetime & dateSeparator & " And " & dateSeparator & maxDatetime & dateSeparator & ") And ( int_line_no =  '" & Request.Params("lineno") & "' ) Order by " & DAT_START_DATE & " ASC," & DAT_FINISH_DATE & " ASC,txt_order_key ASC"
                Dim dtAdapter1 As OleDbDataAdapter = New OleDbDataAdapter(query1, conn)
                Dim cmdbAccessCmdBuilder As New OleDbCommandBuilder(dtAdapter1)
                dtAdapter1.UpdateCommand = cmdbAccessCmdBuilder.GetUpdateCommand()

                dtAdapter1.Fill(dtTable)

                Dim keys(0) As DataColumn
                keys(0) = dtTable.Columns("txt_order_key")
                dtTable.PrimaryKey = keys

                Dim cRow0 As DataRow = dtTable.Rows(0), cRow1 As DataRow
                Dim i As Long = dtTable.Rows.Count, j As Long
                Dim oldGapBetweenFinishAndExplant As Integer = 0
                Dim timedifference1 As Long
                For j = 1 To i - 1  ' i should be greater or equal to 2 because two elements were selected in the previous step
                    cRow1 = dtTable.Rows(j)
                    timedifference1 = DateDiff(DateInterval.Minute, cRow1("dat_start_date"), cRow1("dat_finish_date"))

                    If DateDiff(DateInterval.Minute, cRow0("dat_finish_date"), cRow1("dat_start_date")) < 0 Then
                        cRow1("dat_start_date") = cRow0("dat_finish_date")
                        cRow1("dat_finish_date") = DateAdd(DateInterval.Minute, timedifference1, cRow1("dat_start_date"))
                    End If

                    cRow0 = cRow1
                Next

                Dim hasException As Boolean = False
                finishTime_exPlantDate_Span1(conn, dtTable.Select(Nothing), hasException)
                If hasException Then
                    'exceptionR.Visible = True
                    CacheInsert("Oexception", 1)
                End If

                additionDaysOnExplantDate(conn, dtTable, DataViewRowState.ModifiedCurrent)


                dtAdapter1.Update(dtTable)
                cmdbAccessCmdBuilder.Dispose()
                dtAdapter1.Dispose()

                'to record the last group of order keys you changed in order to mark them with different dot line 
                CacheInsert("slctKeys", Request.Params("frstslctnID") & "," & Request.Params("scndslctnID"))

                actionRequested = "orderlines" ' continue next step to display all order lines

            End If


            'move some orders in one grouplink orders and push back (advance)======
            If actionRequested = "updateorderinbatchA" Then
                Dim minDatetime As DateTime, maxDatetime As DateTime
                query1 = "Select dat_start_date,dat_finish_date,txt_order_key From Esch_Na_tbl_orders  Where txt_order_key in ('" & Request.Params("frstslctnID") & "','" & Request.Params("scndslctnID") & "') Order by " & DAT_START_DATE & " ASC," & DAT_FINISH_DATE & " ASC,txt_order_key ASC"
                command.CommandText = query1
                reader = command.ExecuteReader()
                If reader.Read() Then minDatetime = reader(DAT_START_DATE)
                If reader.Read() Then maxDatetime = reader(DAT_START_DATE)
                reader.Close()

                query1 = "Select * From Esch_Na_tbl_orders Where (" & DAT_START_DATE & " between " & dateSeparator & minDatetime & dateSeparator & " And " & dateSeparator & maxDatetime & dateSeparator & ") And ( int_line_no =  '" & Request.Params("lineno") & "' ) Order by " & DAT_START_DATE & " DESC," & DAT_FINISH_DATE & " DESC,txt_order_key ASC"
                Dim dtAdapter1 As OleDbDataAdapter = New OleDbDataAdapter(query1, conn)
                Dim cmdbAccessCmdBuilder As New OleDbCommandBuilder(dtAdapter1)
                dtAdapter1.UpdateCommand = cmdbAccessCmdBuilder.GetUpdateCommand()

                dtAdapter1.Fill(dtTable)

                Dim keys(0) As DataColumn
                keys(0) = dtTable.Columns("txt_order_key")
                dtTable.PrimaryKey = keys

                Dim cRow0 As DataRow = dtTable.Rows(0), cRow1 As DataRow
                Dim i As Long = dtTable.Rows.Count, j As Long
                Dim oldGapBetweenStartAndExplant As Integer = 0
                Dim timedifference1 As Long
                For j = 1 To i - 1  ' i should be greater or equal to 2 because two elements were selected in the previous step
                    cRow1 = dtTable.Rows(j)
                    timedifference1 = DateDiff(DateInterval.Minute, cRow1("dat_start_date"), cRow1("dat_finish_date"))

                    If DateDiff(DateInterval.Minute, cRow0("dat_start_date"), cRow1("dat_finish_date")) > 0 Then
                        cRow1("dat_finish_date") = cRow0("dat_start_date")
                        cRow1("dat_start_date") = DateAdd(DateInterval.Minute, -timedifference1, cRow1("dat_finish_date"))
                    End If

                    cRow0 = cRow1
                Next

                Dim hasException As Boolean = False
                finishTime_exPlantDate_Span1(conn, dtTable.Select(Nothing), hasException)
                If hasException Then
                    'exceptionR.Visible = True
                    CacheInsert("Oexception", 1)
                End If

                additionDaysOnExplantDate(conn, dtTable, DataViewRowState.ModifiedCurrent)


                dtAdapter1.Update(dtTable)
                cmdbAccessCmdBuilder.Dispose()
                dtAdapter1.Dispose()

                'to record the last group of order keys you changed in order to mark them with different dot line 
                CacheInsert("slctKeys", Request.Params("frstslctnID") & "," & Request.Params("scndslctnID"))


                actionRequested = "orderlines" ' continue next step to display all order lines

            End If


            'move some orders in one group ======
            If actionRequested = "moveingroup" Then
                Dim minDatetime As DateTime, maxDatetime As DateTime, minDatetimeSt As DateTime
                query1 = "Select dat_start_date,dat_finish_date,txt_order_key  From Esch_Na_tbl_orders  Where txt_order_key in ('" & Request.Params("frstslctnID") & "','" & Request.Params("scndslctnID") & "') Order by dat_start_date ASC,dat_finish_date ASC,txt_order_key ASC"
                command.CommandText = query1
                reader = command.ExecuteReader()
                If reader.Read() Then
                    minDatetime = reader("dat_finish_date")
                    minDatetimeSt = reader("dat_start_date")
                End If

                If reader.Read() Then maxDatetime = reader("dat_start_date")
                reader.Close()

                query1 = "Select * From Esch_Na_tbl_orders Where (dat_start_date between " & dateSeparator & minDatetimeSt & dateSeparator & " And " & dateSeparator & maxDatetime & dateSeparator & ")  And  ( int_line_no =  '" & Request.Params("lineno") & "')  Order by dat_start_date ASC,dat_finish_date ASC,txt_order_key ASC"
                Dim dtAdapter1 As OleDbDataAdapter = New OleDbDataAdapter(query1, conn)
                Dim cmdbAccessCmdBuilder As New OleDbCommandBuilder(dtAdapter1)
                dtAdapter1.UpdateCommand = cmdbAccessCmdBuilder.GetUpdateCommand()


                dtAdapter1.Fill(dtTable)

                Dim keys(0) As DataColumn
                keys(0) = dtTable.Columns("txt_order_key")
                dtTable.PrimaryKey = keys

                Dim cRow0 As DataRow, continue1 As Boolean = False
                Dim i As Long = dtTable.Rows.Count, j As Long
                Dim timedifference1 As Long = DateDiff(DateInterval.Minute, minDatetimeSt, DateAdd(DateInterval.Minute, CLng(Request.Params("areaoffsetleft")) * 24 * 60 / pixelsPerDay, startD))

				
                For j = 0 To i - 1  ' i should be greater or equal to 2 because two elements were selected in the previous step
                    cRow0 = dtTable.Rows(j)
                    If cRow0("txt_order_key") = Request.Params("frstslctnID") Then continue1 = True
                    If continue1 Then
					
                        'cRow0("dat_start_date") = DateAdd(DateInterval.Minute, timedifference1, cRow0("dat_start_date"))
						'FANAR+ GO-LIVE, on Nov.9.2016  need calculate start time based on finished time, backward calculation
						cRow0("dat_finish_date") = DateAdd(DateInterval.Minute, timedifference1, cRow0("dat_finish_date"))
                    End If
                    If cRow0("txt_order_key") = Request.Params("scndslctnID") Then
                        Exit For
                    End If
                Next


                Dim hasException As Boolean = False
                finishTime_exPlantDate_Span1(conn, dtTable.Select(Nothing), hasException)
                If hasException Then
                    'exceptionR.Visible = True
                    CacheInsert("Oexception", 1)
                End If

                additionDaysOnExplantDate(conn, dtTable, DataViewRowState.ModifiedCurrent)


                dtAdapter1.Update(dtTable)
                cmdbAccessCmdBuilder.Dispose()
                dtAdapter1.Dispose()

                'to record the last group of order keys you changed in order to mark them with different dot line 
                CacheInsert("slctKeys", Request.Params("frstslctnID") & "," & Request.Params("scndslctnID"))


                actionRequested = "orderlines" ' continue next step to display all order lines

            End If




            'add or delete screw pull mark======
            If actionRequested = "pullscrew" Then
                query1 = "select * From Esch_Na_tbl_prductn_prmtr where  txt_order_key = '" & Request.Params("orderkey") & "'"
                Dim dtAdapter1 As OleDbDataAdapter = New OleDbDataAdapter(query1, conn)
                Dim cmdbAccessCmdBuilder As New OleDbCommandBuilder(dtAdapter1)
                dtAdapter1.DeleteCommand = cmdbAccessCmdBuilder.GetDeleteCommand()
                dtAdapter1.InsertCommand = cmdbAccessCmdBuilder.GetInsertCommand()

                dtAdapter1.Fill(dtTable)

                Dim keys(0) As DataColumn
                keys(0) = dtTable.Columns("txt_order_key")
                dtTable.PrimaryKey = keys


                If dtTable.Rows.Count = 0 Then
                    Dim newRow As DataRow = dtTable.NewRow()
                    newRow("txt_order_key") = Request.Params("orderkey")
                    newRow("txt_pull_screw") = "red"
                    dtTable.Rows.Add(newRow)
                Else
                    dtTable.Rows(0).Delete()
                End If

                dtAdapter1.Update(dtTable)
                cmdbAccessCmdBuilder.Dispose()
                dtAdapter1.Dispose()

                'to record the last group of order keys you changed in order to mark them with different dot line 
                CacheInsert("slctKeys", Request.Params("orderkey") & "")


                actionRequested = "orderlines" ' continue next step to display all order lines

            End If



            'add or delete screw cleaning mark======
            If actionRequested = "cleanscrew" Then

                query1 = "select * From Esch_Na_tbl_prductn_prmtr where  txt_order_key = '" & Request.Params("orderkey") & "'"
                Dim dtAdapter1 As OleDbDataAdapter = New OleDbDataAdapter(query1, conn)
                Dim cmdbAccessCmdBuilder As New OleDbCommandBuilder(dtAdapter1)
                dtAdapter1.DeleteCommand = cmdbAccessCmdBuilder.GetDeleteCommand()
                dtAdapter1.InsertCommand = cmdbAccessCmdBuilder.GetInsertCommand()

                dtAdapter1.Fill(dtTable)

                Dim keys(0) As DataColumn
                keys(0) = dtTable.Columns("txt_order_key")
                dtTable.PrimaryKey = keys


                If dtTable.Rows.Count = 0 Then
                    Dim newRow As DataRow = dtTable.NewRow()
                    newRow("txt_order_key") = Request.Params("orderkey")
                    newRow("txt_pull_screw") = "yellow"
                    dtTable.Rows.Add(newRow)
                Else
                    dtTable.Rows(0).Delete()
                End If

                dtAdapter1.Update(dtTable)
                cmdbAccessCmdBuilder.Dispose()
                dtAdapter1.Dispose()

                'to record the last group of order keys you changed in order to mark them with different dot line 
                CacheInsert("slctKeys", Request.Params("orderkey") & "")

                actionRequested = "orderlines" ' continue next step to display all order lines

            End If

        End If

        'If want to get all order lines' information ==========
        If actionRequested = "orderlines" Then
            query1 = "Select table1.txt_order_key,txt_currency,int_status_key,int_line_no,dat_start_date,dat_finish_date,txt_lot_no,int_span,planned_production_qty,flt_unallocate_qty,flt_order_qty,flt_actual_completed,txt_color,txt_item_no,dat_etd,txt_VIP,txt_process_technics,txt_FDA,txt_remark,txt_payment_status,table2.txt_pull_screw From  " & _
            " ((Select txt_order_key,txt_currency,int_status_key,int_line_no,dat_start_date,dat_finish_date,txt_lot_no,int_span,planned_production_qty,flt_unallocate_qty,flt_order_qty,flt_actual_completed,txt_color,txt_item_no,dat_etd,txt_VIP,txt_process_technics,txt_FDA,txt_remark,txt_payment_status From Esch_Na_tbl_orders  Where ( int_line_no  = '" & Request.Params("lineno") & "' ) And (" & DAT_FINISH_DATE & "  between " & dateSeparator & startD & dateSeparator & " And " & dateSeparator & DateAdd("d", daysFromStart, startD) & dateSeparator & ")) as table1 Left Join (select txt_order_key,txt_pull_screw From Esch_Na_tbl_prductn_prmtr) as table2 On (table1.txt_order_key = table2.txt_order_key)) " & _
            " Order by " & DAT_START_DATE & " ASC," & DAT_FINISH_DATE & " ASC,table1.txt_order_key ASC "
            command.CommandText = query1
            reader = command.ExecuteReader()
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

            Dim selectKeys As String = String.Empty
            If Not CacheFrom("slctKeys") Is Nothing Then
                selectKeys = CType(CacheFrom("slctKeys"), String)
            End If

            Do While reader.Read()

                currentLotNo = reader("txt_lot_no").ToString
                If (Not String.IsNullOrEmpty(currentLotNo)) AndAlso (Not String.Equals(currentLotNo, previousLotNo, StringComparison.OrdinalIgnoreCase)) Then
                    If listOfLotNo.Contains(currentLotNo) Then
                        currentLotNo = "<span style='color:red;'>" & currentLotNo & "</span>"
                    Else
                        listOfLotNo.Add(currentLotNo)
                    End If
                End If

                previousLotNo = currentLotNo


                interMinutes = CLng(DateDiff(DateInterval.Minute, reader("dat_start_date"), reader("dat_finish_date")) * pixelsPerDay / (24 * 60))

                If DBNull.Value.Equals(reader("txt_process_technics")) Then
                    screwType = String.Empty
                Else
                    screwType = "&nbsp;&nbsp;(screw:" & reader("txt_process_technics") & ")&nbsp;&nbsp;"
                End If



                percentageOfCompletion = reader("flt_actual_completed").ToString()
                additionGap = 0
                If Not String.IsNullOrEmpty(percentageOfCompletion) Then
                    If CInt(percentageOfCompletion) > 0 Then
                        additionGap = 4 * 6
                        percentageOfCompletion &= "%"
                    Else
                        percentageOfCompletion = String.Empty
                    End If
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

                orderSatus = reader("int_status_key").ToString

                numOfAllowableChanges = Right(orderSatus, 1)
                If Not IsNumeric(numOfAllowableChanges) Then
                    numOfAllowableChanges = String.Empty
                End If


                'If orderSatus.IndexOf("2") <> -1 Then
                '    numOfAllowableChanges = "2"
                'Else
                '    If orderSatus.IndexOf("1") <> -1 Then
                '        numOfAllowableChanges = "1"
                '    Else
                '        If orderSatus.IndexOf("0") <> -1 Then
                '            numOfAllowableChanges = "0"
                '        Else
                '            numOfAllowableChanges = String.Empty
                '        End If
                '    End If
                'End If


                RDDqtyNewOffset = 5
                If orderSatus.IndexOf("NEW", StringComparison.OrdinalIgnoreCase) <> -1 Then 'new orders
                    orderSatus = "<span id='x'>new" & numOfAllowableChanges & "</span>"
                Else
                    If orderSatus.IndexOf("-R", StringComparison.OrdinalIgnoreCase) <> -1 Then 'RDD changed
                        orderSatus = "<span id='y'>rdd" & numOfAllowableChanges & "</span>"
                    Else
                        If orderSatus.IndexOf("-Q", StringComparison.OrdinalIgnoreCase) <> -1 Then 'Quantity changed
                            orderSatus = "<span id='y'>qty" & numOfAllowableChanges & "</span>"
                        Else
                            RDDqtyNewOffset = 0
                            orderSatus = String.Empty
                        End If
                    End If
                End If


                'Shanghai does not want this feature
                If ShowRMBcurrency AndAlso (Not DBNull.Value.Equals(reader("txt_currency"))) AndAlso reader("txt_currency").ToString.Equals("RMB", StringComparison.OrdinalIgnoreCase) Then
                    currency1 = "<span style='color:#6F2927;'>R</span>"
                    RDDqtyNewOffset += 1
                Else
                    currency1 = String.Empty
                End If

                SPANandETD = reader("int_span") & " (" & reader("dat_etd") & ")"

                'orderSatus = "<span style='left:-" & (9 + offsetForChrome) & "em;position:relative;color:red;'>new</span>"
                leftPosForRSDspan = "left:-" & (SPANandETD.Length + RDDqtyNewOffset) * 7 & "px;"

                currentOrderkey = reader("txt_order_key")
                If selectKeys.IndexOf(currentOrderkey) > -1 Then
                    blockBorderStyle = "border-width:3px 1px;border-style:solid;border-color:red purple blue green;"
                Else
                    blockBorderStyle = "border:1px solid black;"
                End If



                sndbck.Append("<div  class='g-" & reader("txt_color").ToString() & "' style='height:16px;position:relative;" & blockBorderStyle & "margin:5px 0;left:" & CLng(DateDiff("n", startD, reader("dat_start_date")) * pixelsPerDay / (24 * 60)) & "px;width:" & (interMinutes - 2) & "px' id='" & reader("txt_order_key") & "'><span style='" & leftPosForRSDspan & "position:relative'>" & SPANandETD & currency1 & "&nbsp;" & orderSatus & "</span>" & "<span class='sw' style='left:" & interMinutes - 2 & "px;" & marginTop & "'>" & pullScrew & "<span style='" & offsetToLeftDueToScrewForCompletionPercent & "'>" & percentageOfCompletion & "</span></span>" &
                             "<span style='position:absolute;left:" & (interMinutes + additionGap) & "px'>" & "&nbsp;&nbsp;&nbsp;&nbsp;" & currentLotNo & "&nbsp;&nbsp;&nbsp;" & reader("txt_item_no") & "&nbsp;" & reader("planned_production_qty") & "KG " & screwType & reader("txt_FDA") & reader("txt_VIP") & "&nbsp;&nbsp;(&nbsp;<span id='x'>" & currentOrderkey & "</span>) " & reader("txt_payment_status") & "&nbsp;" & reader("txt_remark") & "</span></div>")

            Loop

            reader.Close()

            If Not String.IsNullOrEmpty(userName) Then 'show a message there is other user operating on the key table
                sndbck.Insert(0, "<div style='color:red;font-size:150%;'>" & userName & " is using the order detail table!</div><br />")
            End If

            If Not autOp Then 'if you do not have the authority to do the operaton.
                sndbck.Insert(0, "<div style='color:red;font-size:150%;'>You are not allowed to do the operation on this page!</div><br />")
            End If

        End If




        'Get production lines' information ===
        If actionRequested = "productionlines" Then

            Dim arrayLine1 As Integer() = arrayOfLines()

            Dim i As Integer = arrayLine1.Count
            ReDim Preserve arrayLine1(i)
            arrayLine1(i) = CInt(valueOf("intDummyLine"))
            i += 1
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
                'sndbck.Append("<div class='sc' style='text-align:center;width:" & (pixelsPerDay - 1) & "px;" & IIf(j1 Mod 2 = 1, "background-color:#00FFFF;", "") & "'>" & DateAdd("d", j1 - 1, startD).ToString("ddd  MMM d") & "</div>")
            Next
        End If

        'provide background for background mark lines ------------------
        If actionRequested = "bckgrndImg" Then
            For j1 = 1 To daysFromStart
                sndbck.Append("<div class='bg' style='width:" & (pixelsPerDay - 1) & "px;left:" & (pixelsPerDay * (j1 - 1)) & "px;'></div>")
                'sndbck.Append("<div class='bg' style='width:" & (pixelsPerDay - 1) & "px;'></div>")
            Next
        End If

        'provide content for context menu
        If actionRequested = "contextMenu" Then

            command.CommandText = "SELECT dat_start_date FROM Esch_Na_tbl_orders WHERE txt_order_key = '" & Request.Params("orderkey") & "'"
            reader = command.ExecuteReader()
            Dim startTimeForThisOrder As DateTime = DateTime.Now
            If reader.Read() Then startTimeForThisOrder = CDate(reader("dat_start_date"))
            If Not reader.IsClosed Then reader.Close()


            Dim arrayLine1 As Integer() = arrayOfLines()

            Dim i As Integer = arrayLine1.Count
            ReDim Preserve arrayLine1(i)
            arrayLine1(i) = CInt(valueOf("intDummyLine"))
            Dim j1 As Integer = 0


            sndbck.Append("<ul class='cmenu' id1='" & Request.Params("orderkey") & "'>")
            For j1 = 0 To i
                If arrayLine1(j1) = CInt((Request.Params("lineno"))) Then
                    sndbck.Append("<li  onmousemove='PX.dsplyCntxtMnuLst(this);' onmouseout='PX.subMouseOverContex(false);' class='cmenulist' id1='" & arrayLine1(j1) & "'><div id='dateTextContainer'><input name='dateText' id='dateText' type='text' onkeydown='if(event.which || event.keyCode){if ((event.which == 13) || (event.keyCode == 13)) {PX.newStartTime(this);PX.CloseContext();return false;}} else {return true};' value = '" & startTimeForThisOrder.ToString & "' /><button><img src='calendar.gif'></button></div></li>")
                Else
                    sndbck.Append("<li onclick='PX.changeLine(this);PX.CloseContext();' onmousemove='PX.dsplyCntxtMnuLst(this);' onmouseout='PX.subMouseOverContex(false);' class='cmenulist' id1='" & arrayLine1(j1) & "'>" & "L" & arrayLine1(j1) & "</li>")
                End If
            Next
            sndbck.Append("</ul>")


        End If


        conn.Close()

        Response.Write(sndbck.ToString())

    End Sub



End Class
