Imports System.IO
Imports System.Text
Imports System.Data
Imports System.Data.OleDb
Imports System.Globalization
Imports Microsoft.VisualBasic
Imports System.Linq
Imports System.Threading



Partial Class interface_batchCreationAndUpload
    Inherits basepage1

    Public Delegate Sub asychrSub() 'to run an asynchronous function, keep a record on what item has been arranged on which production line. Based on historical records, we can know which production line has been arranged to produce the FG item more than other production lines


    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        SDS1.ConnectionString = ConfigurationManager.ConnectionStrings(dbConnectionName).ProviderName & ConfigurationManager.ConnectionStrings(dbConnectionName).ConnectionString
        Message.Text = String.Empty
        StatusLabel.Text = String.Empty


        'create some checkbox to show production lines
        Dim checked As Boolean = cbAllLines.Checked


        Dim arrayLine1 As Integer() = arrayOfLines()
        For i As Integer = 0 To arrayLine1.Count - 1
            Dim a As CheckBox = New CheckBox
            a.Checked = checked
            a.Text = arrayLine1(i)
            a.ID = arrayLine1(i)
            a.EnableTheming = True
            a.TextAlign = TextAlign.Left

            linesCollection.Controls.Add(a)
        Next





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

            earlierTime.Text = DateTime.Today.ToShortDateString
            laterTime.Text = DateTime.Today.AddDays(2).ToShortDateString

            'date time format based on original culture
            Thread.CurrentThread.CurrentCulture = originalCulture

            ddlHour1.SelectedIndex = DateTime.Now.Hour
            ddlHour2.SelectedIndex = ddlHour1.SelectedIndex
            ddlMinute1.SelectedIndex = DateTime.Now.Minute
            ddlMinute2.SelectedIndex = ddlMinute1.SelectedIndex



        End If

        cbAllOrders.Attributes.Add("onclick", CreateBatch.ClientID & ".disabled=true;" & gnrtFlndFTP.ClientID & ".disabled=true;")
        CreateBatch.Attributes.Add("onclick", gnrtFlndFTP.ClientID & ".disabled=true;")

        displayButtonStatus()

    End Sub

    'display button's status, enabled or not
    Protected Sub displayButtonStatus()
        If CacheFrom("batchCreationBC") Is Nothing Then
            CreateBatch.Enabled = False
        Else
            CreateBatch.Enabled = True
        End If

        If CacheFrom("batchCreationGFF") Is Nothing Then
            gnrtFlndFTP.Enabled = False
        Else
            gnrtFlndFTP.Enabled = True
        End If

    End Sub


    Protected Sub CreateBatch_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles CreateBatch.Click


        Dim msgRtrn As New StringBuilder()

        'get data from table Esch_Na_tbl_batch_no_group_and_batch_rules
        Dim connParam As OleDbConnection = New OleDbConnection(ConfigurationManager.ConnectionStrings(dbConnForParam).ProviderName & ConfigurationManager.ConnectionStrings(dbConnForParam).ConnectionString)
        Dim dtAdapterParam As OleDbDataAdapter = New OleDbDataAdapter("SELECT * FROM Esch_Na_tbl_batch_no_group_and_batch_rules", connParam)
        Dim cmdbAccessCmdBuilder1 As New OleDbCommandBuilder(dtAdapterParam)
        dtAdapterParam.UpdateCommand = cmdbAccessCmdBuilder1.GetUpdateCommand()
        Dim dtTableParam As DataTable = New DataTable
        dtAdapterParam.Fill(dtTableParam)
        ''update field txt_current_no based on the field value of txt_last_no
        'For Each a As DataRow In dtTableParam.Rows
        '    a.Item("txt_current_no") = a.Item("txt_last_no")
        'Next



        'get data from table Esch_Na_tbl_BatchNO
        Dim conn As OleDbConnection = New OleDbConnection(SDS1.ConnectionString)
        Dim dtAdapter As OleDbDataAdapter = New OleDbDataAdapter("SELECT * FROM Esch_Na_tbl_BatchNO ", conn)
        Dim cmdbAccessCmdBuilder As New OleDbCommandBuilder(dtAdapter)
        dtAdapter.UpdateCommand = cmdbAccessCmdBuilder.GetUpdateCommand()
        Dim dtTable As DataTable = New DataTable()
        dtAdapter.Fill(dtTable)

        Dim errorOccurred As Boolean = False
        Dim txt_lot_no As String = String.Empty

        Dim batchNoRows() As DataRow = dtTable.Select("txt_line_group IS NULL")
        If batchNoRows.Count > 0 Then
            errorOccurred = True
            msgRtrn.Append("<div style='color:red;'>Please check currency or POD or shipto, etc. for order: " & batchNoRows(0).Item("txt_order_key") & "</div>")
        End If

        If Not errorOccurred Then

            'iterating repeateritem to find new value on the web and use its value to update background dataTable 
            For Each rp As RepeaterItem In Rt1.Items
                errorOccurred = False
                Dim dtRows() As DataRow = dtTable.Select("txt_order_key = '" & DirectCast(rp.FindControl("O_K"), Label).Text & "'")
                If dtRows.Count > 0 Then
                    dtRows(0).Item("bnl_Y_or_N") = DirectCast(rp.FindControl("Y_N"), CheckBox).Checked ' decide whether to create new batch No. for this row or not
                    txt_lot_no = DirectCast(rp.FindControl("L_N"), TextBox).Text


                    If String.IsNullOrEmpty(txt_lot_no) OrElse (IsNumeric(txt_lot_no) AndAlso (txt_lot_no.Length = 8) AndAlso (CLng(txt_lot_no) < 100000000)) Then
                        dtRows(0).Item("txt_lot_no") = txt_lot_no
                    Else
                        errorOccurred = True
                        msgRtrn.AppendLine("<div style='color:red;'> txt_lot_no must be 8-digit number. Wrong one like : " & txt_lot_no & "</div>")
                        Exit For
                    End If

                End If
            Next

        End If


        Dim sortStr As String = SDS1.SelectCommand.Remove(0, SDS1.SelectCommand.IndexOf("ORDER BY") + 8)
        sortStr = "int_line_no ASC,txt_item_no,txt_currency ASC,txt_package_code,dat_start_date ASC,dat_finish_date,txt_order_key ASC"

        Dim accumulateQty As Integer = 0, maxQty As Integer = CInt(valueOf("intMaxQuantityForOneLot")), toleranceForCombination As Integer = CInt(valueOf("intCombinetolerance"))

        Dim batchGroups = (From groupName In dtTable.AsEnumerable() Select groupName.Field(Of String)("txt_line_group")).Distinct()

        Dim affectedOrderLines As Integer = 0
        Dim newlyCreatedLotNo As Integer = 0
        'iterating all the orders for each batch group

        If Not errorOccurred Then
            For Each batchGroup As String In batchGroups

                Dim RowsToBeCreatedBatchNo() As DataRow = dtTable.Select("txt_lot_no = '' And bnl_Y_or_N = True And txt_line_group = '" & batchGroup & "'", sortStr) 'lot is null and flag is checked to create new lot no.
                Dim batchRow() As DataRow = dtTableParam.Select("int_Batch_NO_group = " & batchGroup)

                If batchRow.Count = 0 Then
                    errorOccurred = True
                    'msgPopUP("Batch No. can not be found corresponding to txl_line_group (int_Batch_NO_group):" & batchGroup & " and txt_currency:" & RowsToBeCreatedBatchNo(i).Item("txt_currency"), Message)
                    msgRtrn.AppendLine("<div style='color:red;'> Batch No. group can not be found for " & batchGroup & ".</div>")
                    Exit For
                End If


                If RowsToBeCreatedBatchNo.Count > 0 Then

                    If batchRow.Count > 0 Then
                        batchRow(0).Item("txt_current_no") = CLng(batchRow(0).Item("txt_current_no")) + 1
                        batchRow(0).Item("txt_current_no") = batchRow(0).Item("txt_current_no").Insert(0, New String("0"c, 8 - batchRow(0).Item("txt_current_no").Length))
                        newlyCreatedLotNo += 1
                        If CLng(batchRow(0).Item("txt_current_no")) < CLng(batchRow(0).Item("txt_minimum_no")) OrElse CLng(batchRow(0).Item("txt_current_no")) > CLng(batchRow(0).Item("txt_maximum_no")) Then
                            errorOccurred = True 'if the newly generated batch no exceeds the limits
                            msgRtrn.AppendLine("<div style='color:red;'> Batch No.: " & batchRow(0).Item("txt_current_no") & " is already outside the limits set by minimum and maximum of batch group : " & batchGroup & ".</div>")
                        End If

                        RowsToBeCreatedBatchNo(0).Item("txt_lot_no") = batchRow(0).Item("txt_current_no")
                        affectedOrderLines += 1
                        RowsToBeCreatedBatchNo(0).Item("bnlNewBatchNO") = True 'marked with new status for this row for further process on generating EDI file later on
                        accumulateQty = RowsToBeCreatedBatchNo(0).Item("planned_production_qty")
                    End If
                End If

                If Not errorOccurred Then
                    For i As Integer = 1 To (RowsToBeCreatedBatchNo.Count - 1)

                        accumulateQty += RowsToBeCreatedBatchNo(i).Item("planned_production_qty")

                        'same production line and same item and same currency and same package code and total quantity less than a specific quantity say 20000kg and adjacent orders start time and finish time are close
                        If RowsToBeCreatedBatchNo(i).Item("int_line_no") = RowsToBeCreatedBatchNo(i - 1).Item("int_line_no") AndAlso
                            RowsToBeCreatedBatchNo(i).Item("txt_item_no") = RowsToBeCreatedBatchNo(i - 1).Item("txt_item_no") AndAlso
                            RowsToBeCreatedBatchNo(i).Item("txt_package_code") = RowsToBeCreatedBatchNo(i - 1).Item("txt_package_code") AndAlso
                            accumulateQty <= maxQty AndAlso
                            (Math.Abs(DateDiff(DateInterval.Minute, RowsToBeCreatedBatchNo(i).Item("dat_start_date"), RowsToBeCreatedBatchNo(i - 1).Item("dat_finish_date"))) <= toleranceForCombination OrElse
                             Math.Abs(DateDiff(DateInterval.Minute, RowsToBeCreatedBatchNo(i).Item("dat_start_date"), RowsToBeCreatedBatchNo(i - 1).Item("dat_start_date"))) <= toleranceForCombination) Then

                            'RowsToBeCreatedBatchNo(i).Item("txt_currency") = RowsToBeCreatedBatchNo(i - 1).Item("txt_currency") AndAlso
                            'do nothing
                        Else

                            batchRow(0).Item("txt_current_no") = CLng(batchRow(0).Item("txt_current_no")) + 1
                            batchRow(0).Item("txt_current_no") = batchRow(0).Item("txt_current_no").Insert(0, New String("0"c, 8 - batchRow(0).Item("txt_current_no").Length))
                            newlyCreatedLotNo += 1
                            If CLng(batchRow(0).Item("txt_current_no")) < CLng(batchRow(0).Item("txt_minimum_no")) OrElse CLng(batchRow(0).Item("txt_current_no")) > CLng(batchRow(0).Item("txt_maximum_no")) Then
                                errorOccurred = True
                                'msgPopUP("Batch No.: " & batchRow(0).Item("txt_current_no") & " is already outside the limits set by minimum and maximum of batch group " & batchGroup, Message)
                                msgRtrn.AppendLine("<div style='color:red;'> Batch No.: " & batchRow(0).Item("txt_current_no") & " is already outside the limits set by minimum and maximum of batch group : " & batchGroup & ".</div>")
                                Exit For
                            End If

                            accumulateQty = RowsToBeCreatedBatchNo(i).Item("planned_production_qty")

                        End If

                        RowsToBeCreatedBatchNo(i).Item("txt_lot_no") = batchRow(0).Item("txt_current_no")
                        affectedOrderLines += 1
                        RowsToBeCreatedBatchNo(i).Item("bnlNewBatchNO") = True 'marked with new status for this row for further process on generating EDI file later on
                    Next

                End If
            Next

        End If

        'if error occurred, cancel the revision on the datarow
        If errorOccurred Then
            dtTable.RejectChanges()
            dtTableParam.RejectChanges()
        End If


        Try
            dtAdapter.Update(dtTable)
        Catch ex As Exception
            errorOccurred = True
            msgRtrn.AppendLine("<div style='color:red;'>" & ex.Message & ".</div>")
        End Try



        dtTable.Dispose()
        dtAdapter.Dispose()
        'conn.Close()
        conn.Dispose()


        Try
            dtAdapterParam.Update(dtTableParam)
        Catch ex As Exception
            errorOccurred = True
            msgRtrn.AppendLine("<div style='color:red;'>" & ex.Message & ".</div>")
        End Try


        dtTableParam.Dispose()
        dtAdapterParam.Dispose()
        'connParam.Close()

        connParam.Dispose()

        If Not errorOccurred Then
            msgRtrn.AppendLine("<div style='color:blue;'>The number of newly created lot no  is " & newlyCreatedLotNo & "</div>")
            msgRtrn.AppendLine("<div style='color:blue;'>The number of orders which have been assigned new lot no  is " & affectedOrderLines & "</div>")
        End If

        msgPopUP(msgRtrn.ToString, Message, False, False)

        Rt1.DataBind()

        CacheInsert("batchCreationGFF", "enabled")

        displayButtonStatus()

    End Sub


    ''' <summary>
    ''' based on selected production lines and given period, select all the orders to be ready for batch No. creation
    ''' </summary>
    Protected Sub prdctnOrdrs_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles prdctnOrdrs.Click

        Dim msgRtrn As New StringBuilder()

        If Not IsDate(earlierTime.Text) Then
            earlierTime.ForeColor = Drawing.Color.Red
            msgPopUP("Illegal date time!", Message, True, False)
            Return
        End If


        If Not IsDate(laterTime.Text) Then
            laterTime.ForeColor = Drawing.Color.Red
            msgPopUP("Illegal date time!", Message, True, False)
            Return
        End If

        'delete all the records in the table
        Dim connstr As String = ConfigurationManager.ConnectionStrings(dbConnectionName).ProviderName & ConfigurationManager.ConnectionStrings(dbConnectionName).ConnectionString
        Dim conn As OleDbConnection = New OleDbConnection(connstr)


        Dim continueToUpdate As Boolean = True

        Dim dtUpdateTo As OleDbDataAdapter = New OleDbDataAdapter("SELECT * FROM Esch_Na_tbl_BatchNO", conn)
        Dim cmdbAccessCmdBuilder As New OleDbCommandBuilder(dtUpdateTo)
        dtUpdateTo.InsertCommand = cmdbAccessCmdBuilder.GetInsertCommand()
        dtUpdateTo.DeleteCommand = cmdbAccessCmdBuilder.GetDeleteCommand()


        Dim dtTableUpdateTo As DataTable = New DataTable()
        dtUpdateTo.Fill(dtTableUpdateTo)
        For Each a As DataRow In dtTableUpdateTo.Rows
            a.Delete()  'firstly delete all the records in the table
        Next


        Dim sqlWhereClause As StringBuilder = New StringBuilder(" WHERE (txt_lot_no IS NULL OR txt_lot_no = '' ) AND  CAST(int_line_no as VARCHAR(5)) In (")
        Dim noLineWasSelected As Boolean = True 'if none of production lines was selected
        For Each a As CheckBox In linesCollection.Controls
            If a.Checked Then
                sqlWhereClause.Append("'" & a.Text & "',")
                noLineWasSelected = False
            End If
        Next

        If noLineWasSelected Then
            sqlWhereClause.Append("'-1') ") 'assign a false production line  number to the field to delete all the records in table Esch_Na_tbl_BatchNO
        Else
            sqlWhereClause.Remove(sqlWhereClause.Length - 1, 1).Append(") ")
            sqlWhereClause.Append("  And (" & ddlColumn.SelectedValue & " Between " & dateSeparator & CDate(earlierTime.Text).AddHours(ddlHour1.SelectedIndex).AddMinutes(ddlMinute1.SelectedIndex) & dateSeparator & " and " & dateSeparator & CDate(laterTime.Text).AddHours(ddlHour2.SelectedIndex).AddMinutes(ddlMinute2.SelectedIndex) & dateSeparator & ") ")
        End If

        Dim dtFrom0 As OleDbDataAdapter = New OleDbDataAdapter("SELECT * FROM Esch_Na_tbl_orders " & sqlWhereClause.ToString() & " ORDER BY int_line_no,txt_item_no,dat_start_date ", conn)
        Dim dtTableFrom0 As DataTable = New DataTable()
        dtFrom0.Fill(dtTableFrom0)
        'how many records to be inserted into Esch_Na_tbl_BatchNO
        Dim recordsCount As Integer = dtTableFrom0.Rows.Count


        'connect to another database to get information about Batch No. group 
        connstr = ConfigurationManager.ConnectionStrings(dbConnForParam).ProviderName & ConfigurationManager.ConnectionStrings(dbConnForParam).ConnectionString
        Dim connParam As OleDbConnection = New OleDbConnection(connstr)

        'Add BLP logic here =======================================================================
        'add logic for BLP orders which are considered for export but destination is within the same country, then we can use non-fumigated pallet for these BLP export order 
        'use customer information and POD to identify which order could be applied on this rule ========================================
        Dim dtBLP As OleDbDataAdapter = New OleDbDataAdapter("SELECT * FROM Esch_Na_tbl_BLP_new ", connParam)
        Dim dtTableBLP As DataTable = New DataTable()
        dtBLP.Fill(dtTableBLP)

        Dim BLP_Group_list = From c In dtTableBLP.AsEnumerable() Select c.Field(Of String)("groupName") Distinct
        Dim BLPcondition As StringBuilder = New StringBuilder()
        Dim extraChar0 As String

        'for each category in BLP table
        For Each groupName As String In BLP_Group_list

            Dim b() As DataRow = dtTableBLP.Select("groupName ='" & groupName & "'")
            BLPcondition.Clear()
            If b.Count > 0 Then
                For Each i As DataRow In b  'summarize up every condition to specific MOQ rule
                    extraChar0 = String.Empty

                    If i.Item("headerName").ToString().IndexOf("txt") >= 0 Then
                        extraChar0 = "'"
                    End If

                    If i.Item("headerName").ToString().IndexOf("dat") >= 0 Then
                        extraChar0 = dateSeparator
                    End If

                    BLPcondition.Append(i.Item("headerName") & " " & i.Item("relationOperator") & " " & extraChar0 & i.Item("conditionValue") & extraChar0 & " And ")
                Next
                BLPcondition.Remove(BLPcondition.Length - 4 - 1, 5)  'eliminate the last operator ' And '

                Dim c() As DataRow = dtTableFrom0.Select(BLPcondition.ToString(), Nothing)
                'Dim c() As DataRow = dtTableFrom0.Select("txt_ship_method  = 'BS' And txt_local_so  Not Like 'X%NWE%'", Nothing)
                '"txt_ship_method  = 'BS' And txt_local_so  Not Like 'X%NWE%'"
                For Each d As DataRow In c
                    d.Item("txt_currency") = b(0).Item("changeToCurrency")
                Next
            End If
        Next

        'release BLP related dataset dataTable
        dtTableBLP.Dispose()
        dtBLP.Dispose()


        'in order to get line description
        Dim dtFrom2 As OleDbDataAdapter = New OleDbDataAdapter("SELECT * FROM Esch_Na_tbl_LinesAndOwners ", connParam)
        Dim dtTableFrom2 As DataTable = New DataTable()
        dtFrom2.Fill(dtTableFrom2)


        Dim dtAdapterParam As OleDbDataAdapter = New OleDbDataAdapter("SELECT * FROM Esch_Na_tbl_batch_no_group_and_batch_rules", connParam)
        Dim cmdbAccessCmdBuilder1 As New OleDbCommandBuilder(dtAdapterParam)
        dtAdapterParam.UpdateCommand = cmdbAccessCmdBuilder1.GetUpdateCommand()
        Dim dtTableParam As DataTable = New DataTable
        dtAdapterParam.Fill(dtTableParam)
        'update field txt_current_no based on the field value of txt_last_no
        For Each a As DataRow In dtTableParam.Rows
            a.Item("txt_current_no") = a.Item("txt_last_no")
        Next

        Try
            dtAdapterParam.Update(dtTableParam)
        Catch ex As Exception
            continueToUpdate = False
            msgRtrn.AppendLine("<div style='color:red;'>" & ex.Message & ".</div>")
        End Try


        dtTableParam.Dispose()
        dtAdapterParam.Dispose()



        'if we decide to use cheaper non-fumigated pallet for domestic sales only, fumigated pallets are used for export orders only
        Dim defaultCurrency As String = valueOf("strDefaultCurrency")  'default currency used if there is empty value in the field of currency
        Dim defaultPackage As String = valueOf("strDefaultPackage")

        For Each dtrow0 As DataRow In dtTableFrom0.Rows 'get order line data from table Esch_Na_tbl_orders
            Dim newRow As DataRow = dtTableUpdateTo.NewRow()

            If DBNull.Value.Equals(dtrow0.Item("txt_package_code")) OrElse String.IsNullOrEmpty(dtrow0.Item("txt_package_code")) Then
                newRow.Item("txt_package_code") = defaultPackage
            Else
                newRow.Item("txt_package_code") = dtrow0.Item("txt_package_code")
            End If


            If DBNull.Value.Equals(dtrow0.Item("txt_end_user")) OrElse String.IsNullOrEmpty(dtrow0.Item("txt_end_user")) Then
                newRow.Item("txt_end_user") = "EMPTY"
            Else
                newRow.Item("txt_end_user") = dtrow0.Item("txt_end_user")
            End If


            If DBNull.Value.Equals(dtrow0.Item("txt_destination")) OrElse String.IsNullOrEmpty(dtrow0.Item("txt_destination")) Then
                newRow.Item("txt_destination") = "EMPTY"
            Else
                newRow.Item("txt_destination") = dtrow0.Item("txt_destination")
            End If

            If DBNull.Value.Equals(dtrow0.Item("txt_ship_cust_no")) OrElse String.IsNullOrEmpty(dtrow0.Item("txt_ship_cust_no")) Then
                newRow.Item("txt_ship_cust_no") = "EMPTY"
            Else
                newRow.Item("txt_ship_cust_no") = dtrow0.Item("txt_ship_cust_no")
            End If


            If DBNull.Value.Equals(dtrow0.Item("txt_currency")) OrElse String.IsNullOrEmpty(dtrow0.Item("txt_currency")) Then
                newRow.Item("txt_currency") = defaultCurrency
            Else
                newRow.Item("txt_currency") = dtrow0.Item("txt_currency")
            End If


            newRow.Item("int_line_no") = dtrow0.Item("int_line_no")
            newRow.Item("bnl_Y_or_N") = False
            newRow.Item("bnlNewBatchNO") = False
            newRow.Item("txt_item_no") = dtrow0.Item("txt_item_no")
            newRow.Item("txt_lot_no") = dtrow0.Item("txt_lot_no")
            newRow.Item("planned_production_qty") = dtrow0.Item("planned_production_qty")
            newRow.Item("dat_start_date") = dtrow0.Item("dat_start_date")
            newRow.Item("dat_finish_date") = dtrow0.Item("dat_finish_date")
            newRow.Item("txt_order_key") = dtrow0.Item("txt_order_key")



            'insert  txt_line_description and txt_formula_version into table Esch_Na_tbl_BatchNO according to the information in table  Esch_Na_tbl_LinesAndOwners
            Dim linesOwnersGroupRows() As DataRow = dtTableFrom2.Select("int_line_no = " & newRow.Item("int_line_no"))
            If linesOwnersGroupRows.Count > 0 Then
                newRow.Item("txt_line_description") = linesOwnersGroupRows(0).Item("txt_line_description")
                newRow.Item("txt_formula_version") = linesOwnersGroupRows(0).Item("txt_formula_version")
            Else
                continueToUpdate = False
                'msgPopUP("No planner is assigned to production line :" & newRow.Item("int_line_no"), StatusLabel)
                msgRtrn.AppendLine("<div style='color:red;'> " & "No planner is assigned to production line :" & newRow.Item("int_line_no") & "</div>")
                Exit For
            End If



            dtTableUpdateTo.Rows.Add(newRow)
        Next






        'update batch no group based on the conditions set in table Esch_Na_tbl_production_lines_and_batch_no_group
        Dim dtFrom1 As OleDbDataAdapter = New OleDbDataAdapter("SELECT * FROM Esch_Na_tbl_production_lines_and_batch_no_group", connParam)
        Dim dtTableFrom1 As DataTable = New DataTable()
        dtFrom1.Fill(dtTableFrom1)
        Dim listOfBtachNoConditions As New List(Of String)
        Dim listOfBtachGroup As New List(Of String)
        Dim groupNames = From c In dtTableFrom1.AsEnumerable() Select c.Field(Of Integer)("int_Batch_NO_group") Distinct
        Dim batchNoConditions As StringBuilder = New StringBuilder()
        Dim extraChar As String = String.Empty
        Dim conditionValues As String = String.Empty
        For Each group1 In groupNames
            batchNoConditions.Clear()
            Dim b() As DataRow = dtTableFrom1.Select("int_Batch_NO_group = '" & group1 & "'")

            If b.Count > 0 Then
                For Each i As DataRow In b  'summarize up every condition per every batch group no.

                    extraChar = String.Empty
                    conditionValues = i.Item("conditionValue")
                    If i.Item("columnName").ToString().IndexOf("txt") >= 0 OrElse i.Item("columnName").ToString().ToLower = "int_line_no" Then
                        If i.Item("relationOperator").ToString().ToLower().IndexOf("in") >= 0 Then
                            extraChar = ""
                            conditionValues = "('" & conditionValues.Replace(",", "','") & "')"
                        Else
                            extraChar = "'"
                        End If

                    End If

                    If i.Item("columnName").ToString().IndexOf("dat") >= 0 Then
                        extraChar = dateSeparator
                    End If

                    batchNoConditions.Append(i.Item("columnName") & " " & i.Item("relationOperator") & " " & extraChar & conditionValues & extraChar & " And ")
                Next
                batchNoConditions.Remove(batchNoConditions.Length - 4 - 1, 5)  'eliminate the last operator ' And '
                listOfBtachNoConditions.Add(batchNoConditions.ToString)
                listOfBtachGroup.Add(b(0).Item("int_Batch_NO_group"))

            End If

        Next


        Try
            For i As Integer = 0 To listOfBtachNoConditions.Count - 1
                Dim updateToRows() As DataRow = dtTableUpdateTo.Select(listOfBtachNoConditions(i), Nothing, DataViewRowState.Added)

                For Each dtRow As DataRow In updateToRows
                    dtRow.Item("txt_line_group") = listOfBtachGroup(i)
                Next
            Next



        Catch ex As Exception
            msgRtrn.AppendLine("<div style = 'color:red;'> update batch group name to Esch_Na_tbl_BatchNO: " & ex.Message & "</div>")
        End Try




        'write data back to table Esch_Na_tbl_BatchNO
        If continueToUpdate Then
            dtUpdateTo.Update(dtTableUpdateTo)
            'msgPopUP(recordsCount & IIf(recordsCount > 1, " lines are", " line is") & " generated.", Message, False, False)
            msgRtrn.AppendLine("<div style='color:blue;font-size:larger;'> " & recordsCount & IIf(recordsCount > 1, " lines are", " line is") & " generated.</div>")
            Message.ForeColor = Drawing.Color.Blue
        End If





        dtTableFrom0.Dispose()
        dtTableFrom1.Dispose()
        dtTableUpdateTo.Dispose()
        cmdbAccessCmdBuilder.Dispose()
        dtFrom0.Dispose()
        dtFrom1.Dispose()
        dtUpdateTo.Dispose()
        'conn.Close()

        connParam.Dispose()

        conn.Dispose()

        msgPopUP(msgRtrn.ToString, Message, False, False)

        Rt1.DataBind()
        cbAllOrders.Checked = False

        CacheInsert("batchCreationBC", "enabled")
        CacheRemove("batchCreationGFF")

        displayButtonStatus()

    End Sub

    ''' <summary>
    ''' generate EDI file after batch NO. creation and FTP the file to OPM ftp server
    ''' the format of the concent of EDI file is  orgnanization code@Batch No.@Line description - formula version@item no@quantity@start time(mm-dd-yyyy hh:mm:ss)@finish time@order no. aggregation@order no line aggregation@package code
    ''' </summary>
    Protected Sub gnrtFlndFTP_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles gnrtFlndFTP.Click


        If CacheFrom("batchCreationGFF") Is Nothing Then
            displayButtonStatus()
            Return
        End If

        Dim msgRtrn As New StringBuilder()
        Dim errorOccurred As Boolean = False

        Dim userName As String = lockKeyTable(priority.CombineOrBatchCreation)
        If Not String.IsNullOrEmpty(userName) Then
            errorOccurred = True
            msgRtrn.Append("<div style='color:red;font-size:150%;'>" & userName & " is using the order detail table" & "</div>")
        End If


        Dim howManyBatchesSentToOPM As Integer = 0 ' how many batches have been successfully generated and sent to OPM ftp server
        'Dim errMessage As StringBuilder = New StringBuilder()

        'get data from table Esch_Na_tbl_batch_no_group_and_batch_rules
        Dim connParam As OleDbConnection = New OleDbConnection(ConfigurationManager.ConnectionStrings(dbConnForParam).ProviderName & ConfigurationManager.ConnectionStrings(dbConnForParam).ConnectionString)
        Dim dtAdapterParam As OleDbDataAdapter = New OleDbDataAdapter("SELECT * FROM Esch_Na_tbl_batch_no_group_and_batch_rules", connParam)
        Dim cmdbAccessCmdBuilder1 As New OleDbCommandBuilder(dtAdapterParam)
        dtAdapterParam.UpdateCommand = cmdbAccessCmdBuilder1.GetUpdateCommand()
        Dim dtTableParam As DataTable = New DataTable
        dtAdapterParam.Fill(dtTableParam)

        'get data from table Esch_Na_tbl_MTI_List 
        Dim dtAdapterMTIlist As OleDbDataAdapter = New OleDbDataAdapter("SELECT * FROM Esch_Na_tbl_MTI_List WHERE dat_expiration_date >= " & dateSeparator & Today & dateSeparator & " AND dat_effective_date <= " & dateSeparator & Today & dateSeparator, connParam)
        Dim dtTableMTIlist As DataTable = New DataTable
        dtAdapterMTIlist.Fill(dtTableMTIlist)



        'get data from table Esch_Na_tbl_BatchNO
        Dim conn As OleDbConnection = New OleDbConnection(SDS1.ConnectionString)
        Dim dtAdapter As OleDbDataAdapter = New OleDbDataAdapter("SELECT * FROM Esch_Na_tbl_BatchNO", conn)
        Dim cmdbAccessCmdBuilder As New OleDbCommandBuilder(dtAdapter)
        dtAdapter.UpdateCommand = cmdbAccessCmdBuilder.GetUpdateCommand()
        Dim dtTable As DataTable = New DataTable()
        dtAdapter.Fill(dtTable)

        If Not errorOccurred Then
            Dim txt_lot_no As String = String.Empty
            'iterating repeateritem to find new value on the web and use its value to update dataTable behind
            For Each rp As RepeaterItem In Rt1.Items
                errorOccurred = False
                Dim dtRows() As DataRow = dtTable.Select("txt_order_key = '" & DirectCast(rp.FindControl("O_K"), Label).Text & "'")
                If dtRows.Count > 0 Then
                    dtRows(0).Item("bnlNewBatchNO") = DirectCast(rp.FindControl("N_B"), CheckBox).Checked ' decide whether to create new batch No. for this row or not
                    txt_lot_no = DirectCast(rp.FindControl("L_N"), TextBox).Text


                    If String.IsNullOrEmpty(txt_lot_no) OrElse (IsNumeric(txt_lot_no) AndAlso (txt_lot_no.Length = 8) AndAlso (CLng(txt_lot_no) < 100000000)) Then
                        dtRows(0).Item("txt_lot_no") = txt_lot_no
                    Else
                        errorOccurred = True
                        'errMessage.Append("txt_lot_no should be 8-digit number. Wrong one like : " & txt_lot_no)
                        msgRtrn.AppendLine("<div style='color:red;'> txt_lot_no must be 8-digit number. Wrong one like : " & txt_lot_no & "</div>")
                        Exit For
                    End If


                End If

            Next

        End If

        'Dim sortStr = SDS1.SelectCommand.Remove(0, SDS1.SelectCommand.IndexOf("ORDER BY") + 8)
        Dim RowsWithNewBatchNo() As DataRow = dtTable.Select("txt_lot_no <> '' And bnlNewBatchNO = True ", " txt_lot_no ASC,dat_start_date ASC ")
        'Dim accumulateQty As Integer = 0, maxQty As Integer = CInt(valueOf("intMaxQuantityForOneLot")), toleranceForCombination As Integer = CInt(valueOf("intCombinetolerance"))


        If RowsWithNewBatchNo.Count = 0 Then
            errorOccurred = True
            'errMessage.Append("None of new batch No. was created.")
            msgRtrn.AppendLine("<div style='color:red;'>" & "None of new batch number was created." & "</div>")
        End If


        Dim orgnization As String = valueOf("strOrganization")
        Dim pathOfEDI As String = ConfigurationManager.AppSettings("EDIfilesFolder") & "\"
        If Not Directory.Exists(pathOfEDI) Then
            Try
                Directory.CreateDirectory(pathOfEDI)
            Catch ex As Exception
                'msgPopUP("You might not have right to create folder: " & pathOfEDI & " <br />" & ex.Message, Message, True, False)
                msgRtrn.AppendLine("<div style='color:red;'>" & "You might not have right to create folder: " & pathOfEDI & "  " & ex.Message & "</div>")
                Return
            End Try
        End If


        'If pathOfEDI.EndsWith("\\") Then pathOfEDI = pathOfEDI.Replace("\\", "\")
        If Not pathOfEDI.EndsWith("\") Then pathOfEDI &= "\"

        Dim fileNameOfEDI As String = valueOf("strBatchCreatorPrefix") & "." & orgnization & ".dat." & String.Format("{0:yyyyMMddHHmmss}", System.DateTime.Now)
        Dim pathAndFileNameOfEDI As String = pathOfEDI & fileNameOfEDI

        If Not errorOccurred Then
            Using outFile As New StreamWriter(pathAndFileNameOfEDI)

                Dim extraDays As Integer = CInt(valueOf("intDaysAddedForMDD")) 'this is used to provide a buffer for MDD

                Dim counter As Integer = 0 'serve as a counter to record how many order lines would be allowed, say 10, because too many lines would lead to overpass the limit of the content in the database table

                Dim batchNo As String = RowsWithNewBatchNo(0).Item("txt_lot_no")
                Dim prdctnLineAndFormulaVer As String = RowsWithNewBatchNo(0).Item("txt_line_description") & "-" & RowsWithNewBatchNo(0).Item("txt_formula_version")
                Dim itemNo As String = RowsWithNewBatchNo(0).Item("txt_item_no")
                Dim start As String = String.Format("{0:MM-dd-yyyy HH:mm:ss}", CDate(RowsWithNewBatchNo(0).Item("dat_start_date")))
                Dim finish As String = String.Format("{0:MM-dd-yyyy HH:mm:ss}", CDate(RowsWithNewBatchNo(0).Item("dat_finish_date")).AddDays(extraDays))
                Dim qtyTotal As Integer = RowsWithNewBatchNo(0).Item("planned_production_qty")
                Dim aggregatedOrderNo As String = RowsWithNewBatchNo(0).Item("txt_order_key").ToString.Split("-".ToCharArray())(0)
                Dim aggregatedOrderLineNo As String = RowsWithNewBatchNo(0).Item("txt_order_key").ToString.Split("-".ToCharArray())(1)
                Dim package As String = RowsWithNewBatchNo(0).Item("txt_package_code")


                For i As Integer = 1 To (RowsWithNewBatchNo.Count - 1)


                    'same production line and same item and same currency and same package code and total quantity less than a specific quantity say 20000kg and adjacent orders start time and finish time are close
                    If RowsWithNewBatchNo(i).Item("txt_lot_no") = RowsWithNewBatchNo(i - 1).Item("txt_lot_no") Then
                        qtyTotal += RowsWithNewBatchNo(i).Item("planned_production_qty")
                        String.Format("{0:MM-dd-yyyy HH:mm:ss}", CDate(RowsWithNewBatchNo(i).Item("dat_finish_date")).AddDays(extraDays))
                        aggregatedOrderNo &= "/" & RowsWithNewBatchNo(i).Item("txt_order_key").ToString.Split("-".ToCharArray())(0)
                        aggregatedOrderLineNo &= "/" & RowsWithNewBatchNo(i).Item("txt_order_key").ToString.Split("-".ToCharArray())(1)

                        counter += 1

                        If counter > 10 Then
                            errorOccurred = True
                            'errMessage.Append("Too many orders (over 10) are using one batch no.")
                            msgRtrn.AppendLine("<div style='color:red;'>" & "Too many orders (over 10) are using one batch no." & "</div>")
                        End If
                    Else

                        Dim mtiItem0() As DataRow = dtTableMTIlist.Select("txt_MTI_item = '" & itemNo & "'")
                        If mtiItem0.Count > 0 Then
                            aggregatedOrderNo = "MTI"
                            aggregatedOrderLineNo = "1"
                        End If


                        outFile.WriteLine(orgnization & "@" & batchNo & "@" & prdctnLineAndFormulaVer & "@" & itemNo & "@" & qtyTotal & "@" & start & "@" & finish & "@" & aggregatedOrderNo & "@" & aggregatedOrderLineNo & "@" & package & "@")
                        howManyBatchesSentToOPM += 1

                        counter = 0

                        batchNo = RowsWithNewBatchNo(i).Item("txt_lot_no")
                        prdctnLineAndFormulaVer = RowsWithNewBatchNo(i).Item("txt_line_description") & "-" & RowsWithNewBatchNo(i).Item("txt_formula_version")
                        itemNo = RowsWithNewBatchNo(i).Item("txt_item_no")
                        start = String.Format("{0:MM-dd-yyyy HH:mm:ss}", CDate(RowsWithNewBatchNo(i).Item("dat_start_date")))
                        finish = String.Format("{0:MM-dd-yyyy HH:mm:ss}", CDate(RowsWithNewBatchNo(i).Item("dat_finish_date")).AddDays(extraDays))
                        qtyTotal = RowsWithNewBatchNo(i).Item("planned_production_qty")
                        aggregatedOrderNo = RowsWithNewBatchNo(i).Item("txt_order_key").ToString.Split("-".ToCharArray())(0)
                        aggregatedOrderLineNo = RowsWithNewBatchNo(i).Item("txt_order_key").ToString.Split("-".ToCharArray())(1)
                        package = RowsWithNewBatchNo(i).Item("txt_package_code")

                    End If
                Next

                Dim mtiItem() As DataRow = dtTableMTIlist.Select("txt_MTI_item = '" & itemNo & "'")
                If mtiItem.Count > 0 Then
                    aggregatedOrderNo = "MTI"
                    aggregatedOrderLineNo = "1"
                End If

                outFile.WriteLine(orgnization & "@" & batchNo & "@" & prdctnLineAndFormulaVer & "@" & itemNo & "@" & qtyTotal & "@" & start & "@" & finish & "@" & aggregatedOrderNo & "@" & aggregatedOrderLineNo & "@" & package & "@")
                howManyBatchesSentToOPM += 1

            End Using

        End If


        'if error occurred, cancel the revision on the datarow
        If errorOccurred Then

            If File.Exists(pathAndFileNameOfEDI) Then File.Delete(pathAndFileNameOfEDI)

            dtTable.RejectChanges()
            dtTableParam.RejectChanges()
        End If





        'if no error happens, then ftp EDI file to OPM ftp server
        If Not errorOccurred Then
            Dim continues As Boolean = True
            'Dim localpath As String = Server.MapPath("~/")
            Dim localpath As String = ConfigurationManager.AppSettings("EDIfilesFolder")
            If Not localpath.EndsWith("\") Then localpath &= "\"
            localpath &= "interfaceData\"
            If Not Directory.Exists(localpath) Then Directory.CreateDirectory(localpath)


            'get folder and path for batch creation
            Dim localpathsubfolder As String = localpath & "BatchCreation\"
            Dim workingDirectoryAtServer As String = valueOf("strBatchCreatorPath")

            'get user name and password and ftp server IP to log on to server to send files by ftp method
            Dim organization As String = valueOf("strOrganization")
            Dim serverUri As String = valueOf("strFTPserverIP")
            If Not serverUri.StartsWith("ftp://") Then serverUri = "ftp://" & serverUri
            If Not serverUri.EndsWith("/") Then serverUri &= "/"

            If Not workingDirectoryAtServer.EndsWith("/") Then workingDirectoryAtServer &= "/"

            Dim ftpoperation As FTPcls = New FTPcls(valueOf("strFTP_ID"), valueOf("strFTP_PW"), serverUri & workingDirectoryAtServer)

            Dim returnMessageAfterFTPoperation As String = ftpoperation.UploadFileToServer(pathAndFileNameOfEDI, fileNameOfEDI).ToLower()

            If returnMessageAfterFTPoperation = "true" Then

                Dim dtAdapterOrders As OleDbDataAdapter = New OleDbDataAdapter("SELECT txt_order_key,txt_lot_no,int_formula_version FROM Esch_Na_tbl_orders", conn)
                Dim cmdCmdBuilder As New OleDbCommandBuilder(dtAdapterOrders)
                dtAdapterOrders.UpdateCommand = cmdCmdBuilder.GetUpdateCommand()
                Dim dtTableOrders As DataTable = New DataTable()
                dtAdapterOrders.Fill(dtTableOrders)
                Dim keys(0) As DataColumn
                keys(0) = dtTableOrders.Columns("txt_order_key")
                dtTableOrders.PrimaryKey = keys

                'write back all the newly created batch no. to the orders table
                For Each b As DataRow In dtTable.Select("bnlNewBatchNO = True")
                    Dim orderRows() As DataRow = dtTableOrders.Select("txt_order_key = '" & b.Item("txt_order_key") & "'")
                    If orderRows.Count > 0 Then
                        orderRows(0).Item("txt_lot_no") = b.Item("txt_lot_no")
                        orderRows(0).Item("int_formula_version") = b.Item("txt_formula_version")
                    End If
                Next


                dtAdapterOrders.Update(dtTableOrders)
                dtTableOrders.Dispose()
                dtAdapterOrders.Dispose()


                'update table Esch_Na_tbl_BatchNO
                dtAdapter.Update(dtTable)




                'update table Esch_Na_tbl_batch_no_group_and_batch_rules to get the latest sequence of Lot No. (Batch No.)
                'update field txt_current_no based on the field value of txt_last_no
                For Each a As DataRow In dtTableParam.Rows
                    a.Item("txt_last_no") = a.Item("txt_current_no")
                Next

                dtAdapterParam.Update(dtTableParam)




                'backup the EDI file by moving the file to one sub folder
                If File.Exists(pathAndFileNameOfEDI) Then
                    File.Copy(pathAndFileNameOfEDI, localpathsubfolder & fileNameOfEDI, True)
                    File.Delete(pathAndFileNameOfEDI)
                End If


            Else
                'errMessage.Append(returnMessageAfterFTPoperation)
                msgRtrn.AppendLine("<div style='color:red;'>" & returnMessageAfterFTPoperation & "</div>")
                errorOccurred = True
            End If


        End If




        dtTable.Dispose()
        dtAdapter.Dispose()


        dtTableParam.Dispose()
        dtAdapterParam.Dispose()

        conn.Dispose()

        connParam.Dispose()

        'delete the EDI file anyway
        If File.Exists(pathAndFileNameOfEDI) Then
            File.Delete(pathAndFileNameOfEDI)
        End If

        msgRtrn.AppendLine("<div style='color:blue;'>" & "The number of  new batches sent to OPM is " & howManyBatchesSentToOPM & "</div>")

        If String.IsNullOrEmpty(userName) Then
            unlockKeyTable(priority.CombineOrBatchCreation)
        End If

        If Not errorOccurred Then
            Rt1.DataBind()

            CacheRemove("batchCreationGFF")
            CacheRemove("batchCreationBC")
        End If

        displayButtonStatus()

        msgPopUP(msgRtrn.ToString, Message, False, False)

        'use below function to update table 
        Dim asySubroutine As asychrSub = New asychrSub(AddressOf updateLineItemSequenceTable)
        asySubroutine.BeginInvoke(Nothing, Nothing)


    End Sub

    Protected Sub cbAllOrders_CheckedChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles cbAllOrders.CheckedChanged
        Dim conn As OleDbConnection = New OleDbConnection(SDS1.ConnectionString)
        Dim dtAdapter As OleDbDataAdapter = New OleDbDataAdapter("SELECT bnl_Y_or_N,txt_lot_no,txt_order_key FROM Esch_Na_tbl_BatchNO", conn)
        Dim cmdbAccessCmdBuilder As New OleDbCommandBuilder(dtAdapter)
        dtAdapter.UpdateCommand = cmdbAccessCmdBuilder.GetUpdateCommand()
        Dim dtTable As DataTable = New DataTable()
        dtAdapter.Fill(dtTable)

        Dim errorOccurred As Boolean = False
        Dim txt_lot_no As String = String.Empty
        Dim checked As Boolean = cbAllOrders.Checked

        'iterating repeateritem to find new value on the web and use its value to update dataTable behind
        'For Each rp As RepeaterItem In Rt1.Items
        '    errorOccurred = False
        '    Dim dtRows() As DataRow = dtTable.Select("txt_order_key = '" & DirectCast(rp.FindControl("O_K"), Label).Text & "'")
        '    If dtRows.Count > 0 Then
        '        dtRows(0).Item("bnl_Y_or_N") = checked
        '        txt_lot_no = DirectCast(rp.FindControl("L_N"), TextBox).Text


        '        If String.IsNullOrEmpty(txt_lot_no) OrElse (IsNumeric(txt_lot_no) AndAlso (CLng(txt_lot_no) > 10000000) AndAlso (CLng(txt_lot_no) < 100000000)) Then
        '            dtRows(0).Item("txt_lot_no") = txt_lot_no
        '        Else
        '            errorOccurred = True
        '            msgPopUP("txt_lot_no should be 8-digit number. Wrong one like : " & txt_lot_no, Message)
        '            Exit For
        '        End If


        '    End If

        'Next

        For Each rowLine As DataRow In dtTable.Rows
            rowLine.Item("bnl_Y_or_N") = checked
        Next


        'if error occurred, cancel the revision on the datarow
        If errorOccurred Then
            dtTable.RejectChanges()
        End If

        dtAdapter.Update(dtTable)
        dtTable.Dispose()
        dtAdapter.Dispose()
        conn.Close()


        Rt1.DataBind()

    End Sub

    Protected Sub cbAllLines_CheckedChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles cbAllLines.CheckedChanged
        Dim checked As Boolean = cbAllLines.Checked
        For Each a As CheckBox In linesCollection.Controls
            a.Checked = checked
        Next
    End Sub


    ''' <summary>
    ''' keep a record on what item has been arranged on which production line. Based on historical records, we can know which production line has been arranged to produce the FG item more than other production lines
    ''' </summary>
    ''' <remarks></remarks>
    Protected Sub updateLineItemSequenceTable()
        'get data from table Esch_Na_tbl_batch_no_group_and_batch_rules
        Dim conn As OleDbConnection = New OleDbConnection(ConfigurationManager.ConnectionStrings(dbConnectionName).ProviderName & ConfigurationManager.ConnectionStrings(dbConnectionName).ConnectionString)
        Dim dtAdapterFrom As OleDbDataAdapter = New OleDbDataAdapter("SELECT COUNT(bnl_Y_or_N) AS countNum,txt_item_no,int_line_no FROM Esch_Na_tbl_BatchNO WHERE bnl_Y_or_N > 0 GROUP BY txt_item_no,int_line_no ", conn)
        Dim dtTableFrom As DataTable = New DataTable
        dtAdapterFrom.Fill(dtTableFrom)

        Dim dtAdapterTo1 As OleDbDataAdapter = New OleDbDataAdapter("SELECT * FROM Esch_Na_tbl_Item_And_line_no_historical_data ", conn)
        Dim cmdbAccessCmdBuilder1 As New OleDbCommandBuilder(dtAdapterTo1)
        dtAdapterTo1.UpdateCommand = cmdbAccessCmdBuilder1.GetUpdateCommand()
        dtAdapterTo1.InsertCommand = cmdbAccessCmdBuilder1.GetInsertCommand()
        Dim dtTableTo1 As DataTable = New DataTable
        dtAdapterTo1.Fill(dtTableTo1)
        Dim keys(1) As DataColumn
        keys(0) = dtTableTo1.Columns("txtItem")
        keys(1) = dtTableTo1.Columns("txtLineNo")
        dtTableTo1.PrimaryKey = keys


        Dim dtAdapterTo2 As OleDbDataAdapter = New OleDbDataAdapter("SELECT * FROM Esch_Na_tbl_Item_And_line_no_Priority_Sequence ", conn)
        Dim cmdbAccessCmdBuilder2 As New OleDbCommandBuilder(dtAdapterTo2)
        dtAdapterTo2.UpdateCommand = cmdbAccessCmdBuilder2.GetUpdateCommand()
        dtAdapterTo2.InsertCommand = cmdbAccessCmdBuilder2.GetInsertCommand()
        Dim dtTableTo2 As DataTable = New DataTable
        dtAdapterTo2.Fill(dtTableTo2)
        Dim keys2(0) As DataColumn
        keys2(0) = dtTableTo2.Columns("txtItem")
        dtTableTo2.PrimaryKey = keys2

        Dim sequenceOfLine As StringBuilder = New StringBuilder()

        'iterate through dataTable Esch_Na_tbl_BatchNO to update table Esch_Na_tbl_Item_And_line_no_historical_data and  Esch_Na_tbl_Item_And_line_no_Priority_Sequence
        For Each f As DataRow In dtTableFrom.Rows
            Dim to1() As DataRow = dtTableTo1.Select("txtItem = '" & f.Item("txt_item_no") & "' And txtLineNo = '" & f.Item("int_line_no") & "'")
            If to1.Count > 0 Then
                to1(0).Item("lngCountOfArrangement") += f.Item("countNum")
            Else
                Dim newR As DataRow = dtTableTo1.NewRow
                newR.Item("txtItem") = f.Item("txt_item_no")
                newR.Item("txtLineNo") = f.Item("int_line_no")
                newR.Item("lngCountOfArrangement") = f.Item("countNum")
                dtTableTo1.Rows.Add(newR)
            End If

            Dim to11() As DataRow = dtTableTo1.Select("txtItem = '" & f.Item("txt_item_no") & "'", " lngCountOfArrangement Desc ")
            sequenceOfLine.Clear()
            For Each t As DataRow In to11
                sequenceOfLine.Append(t.Item("txtLineNo") & ",")
            Next

            If sequenceOfLine.Length > 0 Then sequenceOfLine.Remove(sequenceOfLine.Length - 1, 1)

            Dim to2() As DataRow = dtTableTo2.Select("txtItem = '" & f.Item("txt_item_no") & "'")

            If to2.Count > 0 Then
                to2(0).Item("txtLineNoSequence") = sequenceOfLine.ToString()
            Else
                Dim newR As DataRow = dtTableTo2.NewRow
                newR.Item("txtItem") = f.Item("txt_item_no")
                newR.Item("txtLineNoSequence") = sequenceOfLine.ToString()
                dtTableTo2.Rows.Add(newR)
            End If


        Next



        dtAdapterTo1.Update(dtTableTo1)
        dtAdapterTo2.Update(dtTableTo2)


        cmdbAccessCmdBuilder1.Dispose()
        cmdbAccessCmdBuilder2.Dispose()
        dtTableTo1.Dispose()
        dtTableTo2.Dispose()
        dtTableFrom.Dispose()

        dtAdapterTo1.Dispose()
        dtAdapterTo2.Dispose()
        dtAdapterFrom.Dispose()

        'conn.Close()
        conn.Dispose()

    End Sub

    Protected Sub listEDI_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles listEDI.Click
        Response.Redirect("batchCreationFilesEDI.aspx")
    End Sub
End Class

