Imports System.IO
Imports System.Text
Imports System.Data
Imports System.Data.SqlClient
Imports System.Globalization
Imports Microsoft.VisualBasic
Imports System.Linq
Imports System.Threading



Partial Class dragDrop_normalOP_combineProduction
    Inherits FrequentPlanActions



    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        SDS1.ConnectionString = ConfigurationManager.ConnectionStrings(dbConnectionName).ConnectionString
        Message.Text = String.Empty
        StatusLabel.Text = String.Empty

        'create some checkbox to show production lines
        Dim checked As Boolean = cbAllLines.Checked


        Dim arrayLine1 As Integer() = arrayOfLines()
        Dim count1 = arrayLine1.Count
        ReDim Preserve arrayLine1(count1 + 1)
        arrayLine1(count1) = valueOf("intDummyLine")
        For i As Integer = 0 To count1
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

            earlierTime.Text = DateTime.Today.AddDays(1).ToShortDateString
            laterTime.Text = DateTime.Today.AddDays(11).ToShortDateString


            'date time format based on original culture
            Thread.CurrentThread.CurrentCulture = originalCulture

            ddlHour1.SelectedIndex = 23
            ddlHour2.SelectedIndex = ddlHour1.SelectedIndex
            ddlMinute1.SelectedIndex = 30
            ddlMinute2.SelectedIndex = ddlMinute1.SelectedIndex
            If String.IsNullOrEmpty(valueOf("intCombineByRSD")) Then
                gapRSD.Text = "7"
            Else
                gapRSD.Text = valueOf("intCombineByRSD")
            End If


        End If


        displayButtonStatus()


    End Sub


    Protected Sub displayButtonStatus()

        If CacheFrom("combineWB") Is Nothing Then
            bTdb.Enabled = False
        Else
            bTdb.Enabled = True
        End If
    End Sub


    ''' <summary>
    ''' based on selected production lines and given period, select all the orders to be ready for combination, all orders with similar items are shown
    ''' all the orders meet above criteria, and orders which are in smalll production lines but their ETD are in the given period
    ''' </summary>
    Protected Sub prdctnOrdrs_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles prdctnOrdrs.Click

        If Not IsDate(earlierTime.Text) Then
            earlierTime.ForeColor = Drawing.Color.Red
            msgPopUP("Illegal date time!", Message)
            Return
        End If


        If Not IsDate(laterTime.Text) Then
            laterTime.ForeColor = Drawing.Color.Red
            msgPopUP("Illegal date time!", Message)
            Return
        End If


        Dim firstDate As DateTime = CDate(earlierTime.Text).AddHours(ddlHour1.SelectedIndex).AddMinutes(ddlMinute1.SelectedIndex)
        Dim secondDate As DateTime = CDate(laterTime.Text).AddHours(ddlHour2.SelectedIndex).AddMinutes(ddlMinute2.SelectedIndex)

        If firstDate.CompareTo(secondDate) > 0 Then
            secondDate = CDate(earlierTime.Text).AddHours(ddlHour1.SelectedIndex).AddMinutes(ddlMinute1.SelectedIndex)
            firstDate = CDate(laterTime.Text).AddHours(ddlHour2.SelectedIndex).AddMinutes(ddlMinute2.SelectedIndex)
        End If


        'delete all the records in the table
        Dim connstr As String = ConfigurationManager.ConnectionStrings(dbConnectionName).ConnectionString
        Dim conn As SqlConnection = New SqlConnection(connstr)



        Dim continueToUpdate As Boolean = True

        Dim dtUpdateTo As SqlDataAdapter = New SqlDataAdapter("SELECT * FROM Esch_Sh_tbl_similar_item_combination", conn)
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

        'where clause includes three parts; ==================
        'one is all the selected lines and orders' start time is in between the two dates
        Dim sqlWhereClause As StringBuilder = New StringBuilder(" WHERE  ")
        sqlWhereClause.Append("( CAST(int_line_no AS VARCHAR(5)) In (")
        Dim noLineWasSelected As Boolean = True 'if none of production lines was selected
        For Each a As CheckBox In linesCollection.Controls
            If a.Checked Then
                sqlWhereClause.Append("'" & a.Text & "',")
                noLineWasSelected = False
            End If
        Next

        If noLineWasSelected Then
            sqlWhereClause.Append("'-1')) ") 'assign a false production line  number to the field to delete all the records in table Esch_Sh_tbl_BatchNO
        Else
            sqlWhereClause.Remove(sqlWhereClause.Length - 1, 1).Append(") ")
            sqlWhereClause.Append("  And (" & ddlColumn.SelectedValue & " Between " & dateSeparator & firstDate & dateSeparator & " and " & dateSeparator & secondDate & dateSeparator & ") And (planned_production_qty > 0)) ")
        End If

        'second is all the small lines picked up and orders' ETD is in between the two dates
        If Not String.IsNullOrEmpty(valueOf("strCombineSpclProdLines")) Then
            sqlWhereClause.Append(" Or ( CAST(int_line_no AS VARCHAR(5)) In ( '" & valueOf("strCombineSpclProdLines").Replace(",", "','") & "' )  And  ( dat_etd  Between " & dateSeparator & firstDate & dateSeparator & " and " & dateSeparator & secondDate & dateSeparator & ")  And (dat_start_date >= " & dateSeparator & firstDate & dateSeparator & ") ) ")
        End If


        'third is all the orders in dummy lines but their quantity is greater than zero
        If fltout.Checked Then
            sqlWhereClause.Append(" Or ( CAST(int_line_no AS VARCHAR(5)) =  '" & valueOf("intDummyLine") & "' And (planned_production_qty > 0)  And (dat_start_date >= " & dateSeparator & firstDate & dateSeparator & ") ) ")
        End If
        'where clause includes three parts; ==================



        Dim dtFrom0 As SqlDataAdapter = New SqlDataAdapter("SELECT txt_item_no,planned_production_qty,int_line_no,txt_lot_no,dat_start_date,dat_finish_date,txt_currency,dat_etd,dat_rdd,txt_remark,txt_grade,txt_color,txt_order_key,int_status_key FROM Esch_Sh_tbl_orders " & sqlWhereClause.ToString() & " ORDER BY txt_item_no ASC, dat_start_date ASC", conn)
        Dim dtTableFrom0 As DataTable = New DataTable()
        dtFrom0.Fill(dtTableFrom0)
        'how many records to be inserted into Esch_Sh_tbl_similar_item_combination
        Dim recordsCount As Integer = 0


        Dim connParam As SqlConnection = New SqlConnection(ConfigurationManager.ConnectionStrings(dbConnForParam).ConnectionString)
        Dim dtFrom1 As SqlDataAdapter = New SqlDataAdapter("SELECT * FROM Esch_Sh_tbl_similar_grade_for_combination", connParam)
        Dim dtTableFrom1 As DataTable = New DataTable()
        dtFrom1.Fill(dtTableFrom1)



        For Each dtrow0 As DataRow In dtTableFrom0.Select(Nothing, "[txt_item_no] ASC, [dat_start_date] ASC, [int_line_no] ASC") 'get order line data from table Esch_Sh_tbl_orders
            Dim newRow As DataRow = dtTableUpdateTo.NewRow()

            Dim itemNo As String = dtrow0.Item("txt_item_no")
            Dim gradeNo As String = dtrow0.Item("txt_grade")

            newRow.Item("txt_item_no") = itemNo
            newRow.Item("planned_production_qty") = dtrow0.Item("planned_production_qty")
            newRow.Item("int_line_no") = dtrow0.Item("int_line_no")
            newRow.Item("txt_lot_no") = dtrow0.Item("txt_lot_no")
            newRow.Item("dat_start_date") = dtrow0.Item("dat_start_date")
            newRow.Item("dat_finish_date") = dtrow0.Item("dat_finish_date")
            newRow.Item("int_status_key") = dtrow0.Item("int_status_key")
            newRow.Item("txt_remark") = dtrow0.Item("txt_remark")
            'newRow.Item("txt_payment_term") = dtrow0.Item("txt_payment_term")
            newRow.Item("txt_currency") = dtrow0.Item("txt_currency")
            'newRow.Item("int_span") = dtrow0.Item("int_span")
            newRow.Item("dat_etd") = dtrow0.Item("dat_etd")
            newRow.Item("dat_rdd") = dtrow0.Item("dat_rdd")
            'newRow.Item("txt_remark") = dtrow0.Item("txt_remark")
            newRow.Item("txt_grade") = gradeNo
            'look up for typical grade 
            Dim typicalGrade() As DataRow = dtTableFrom1.Select("grade = '" & gradeNo & "'")
            If typicalGrade.Count > 0 Then
                newRow.Item("txt_item_no_similar") = itemNo.Replace(gradeNo & "-", typicalGrade(0).Item("typicalGrade").ToString & "-")
            Else
                newRow.Item("txt_item_no_similar") = itemNo
            End If
            newRow.Item("txt_order_key") = dtrow0.Item("txt_order_key")


            dtTableUpdateTo.Rows.Add(newRow)

        Next



        Dim newrow1() As DataRow = dtTableUpdateTo.Select(Nothing, "txt_item_no_similar ASC,txt_item_no ASC, dat_start_date ASC , int_line_no ASC ")
        'currentOne = newrow1.Item("txt_item_no_similar")

        'for one item no, if we do not find at least two lines in the table, we delete the order line because there is no chance to combine with other order
        If newrow1.Count > 2 Then
            Dim currentOne As String = newrow1(1).Item("txt_item_no_similar")
            If currentOne = newrow1(0).Item("txt_item_no_similar") Then
            Else
                newrow1(0).Item("txt_item_no_similar") = "D" ' marked as to be deleted
            End If


            For i As Integer = 2 To newrow1.Count - 1
                currentOne = newrow1(i - 1).Item("txt_item_no_similar")
                If (currentOne = newrow1(i).Item("txt_item_no_similar")) OrElse (currentOne = newrow1(i - 2).Item("txt_item_no_similar")) Then
                Else
                    newrow1(i - 1).Item("txt_combinationFlag") = "D" ' marked as to be deleted
                End If
            Next


            If currentOne = newrow1(newrow1.Count - 1).Item("txt_item_no_similar") Then
            Else
                newrow1(newrow1.Count - 1).Item("txt_combinationFlag") = "D" ' marked as to be deleted
            End If




        Else
            msgPopUP("The number of orders is very small.", Message)
        End If


        For Each a In dtTableUpdateTo.Select("txt_combinationFlag = 'D'")
            a.Delete()
        Next


        'delete those order lines if no one in the list has order status like 'NEW','*-R' or '*-Q'
        Dim newlyAddedRows() As DataRow = dtTableUpdateTo.Select(Nothing, Nothing, DataViewRowState.Added)
        Dim newTable As DataTable = newlyAddedRows.CopyToDataTable()

        If fltout.Checked Then 'whether we apply this planning rule or not
            Dim query1 = (From order1 In newTable.AsEnumerable() Where (order1.Field(Of String)("int_status_key") Like "*REV-R*" _
                          Or order1.Field(Of String)("int_status_key") Like "*REV-Q*" _
                          Or order1.Field(Of String)("int_status_key") Like "*NEW*")
                                Select order1.Field(Of String)("txt_item_no_similar")).Distinct()



            For Each toBeDeleted As DataRow In dtTableUpdateTo.Select(Nothing)
                If Not query1.Contains(toBeDeleted.Item("txt_item_no_similar")) Then
                    toBeDeleted.Delete()
                End If
            Next

        End If


        recordsCount = dtTableUpdateTo.Select(Nothing).Count
        'write data back to table Esch_Sh_tbl_BatchNO
        If continueToUpdate Then
            dtUpdateTo.Update(dtTableUpdateTo)
            msgPopUP("<div style='color:blue;font-size:larger;'>" & recordsCount & IIf(recordsCount > 1, " lines are", " line is") & " generated.(" & Now.ToString & ")</div>", Message, False, False)

        End If


        newTable.Dispose()

        dtTableFrom0.Dispose()
        dtTableFrom1.Dispose()
        dtTableUpdateTo.Dispose()
        cmdbAccessCmdBuilder.Dispose()
        dtFrom0.Dispose()
        dtFrom1.Dispose()
        dtUpdateTo.Dispose()

        conn.Dispose()
        connParam.Close()

        Rt1.DataBind()

        CacheInsert("combineWB", "enabled")
        displayButtonStatus()

    End Sub


    ''' <summary>
    ''' update data back to database
    ''' </summary>
    Protected Sub bTdb_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles bTdb.Click

        If CacheFrom("combineWB") Is Nothing Then
            displayButtonStatus()
            Return
        End If


        Dim errorMsg As StringBuilder = New StringBuilder()
        Dim userName As String = lockKeyTable(priority.CombineOrBatchCreation)
        If Not String.IsNullOrEmpty(userName) Then
            errorMsg.Append("<div style='color:red;font-size:150%;'>" & userName & " is using the order detail table" & "</div>")
        End If

        Dim conn As SqlConnection = New SqlConnection(SDS1.ConnectionString)

        Dim dtFrom0 As SqlDataAdapter = New SqlDataAdapter("SELECT * FROM Esch_Sh_tbl_orders ", conn)
        'Dim dtFrom0 As SqlDataAdapter = New SqlDataAdapter("SELECT txt_remark,planned_production_qty,int_line_no,txt_lot_no,dat_start_date,dat_finish_date,txt_order_key,flt_working_hours,int_change_over_time,dat_new_explant,txt_item_no,txt_grade,int_span,dat_etd,txt_end_user FROM Esch_Sh_tbl_orders ", conn)
        Dim cmdbAccessCmdBuilder0 As New SqlCommandBuilder(dtFrom0)
        dtFrom0.UpdateCommand = cmdbAccessCmdBuilder0.GetUpdateCommand()
        Dim dtTableFrom0 As DataTable = New DataTable()
        dtFrom0.Fill(dtTableFrom0)

        Dim keys0(0) As DataColumn
        keys0(0) = dtTableFrom0.Columns("txt_order_key")
        dtTableFrom0.PrimaryKey = keys0



        Dim dtAdapter As SqlDataAdapter = New SqlDataAdapter("SELECT planned_production_qty,int_line_no,txt_lot_no,dat_start_date,txt_remark,txt_order_key FROM Esch_Sh_tbl_similar_item_combination", conn)
        Dim cmdbAccessCmdBuilder As New SqlCommandBuilder(dtAdapter)
        dtAdapter.UpdateCommand = cmdbAccessCmdBuilder.GetUpdateCommand()
        Dim dtTable As DataTable = New DataTable()
        dtAdapter.Fill(dtTable)

        Dim keys(0) As DataColumn
        keys(0) = dtTable.Columns("txt_order_key")
        dtTable.PrimaryKey = keys



        Dim txt_lot_no As String = String.Empty
        Dim qty As String = String.Empty
        Dim line As String = String.Empty
        Dim start As String = String.Empty
        'Dim finish As String = String.Empty
        Dim remark As String = String.Empty
        Dim orderKey As String = String.Empty

        'iterating repeateritem to find new value on the web and use its value to update dataTable behind
        For Each rp As RepeaterItem In Rt1.Items

            qty = DirectCast(rp.FindControl("qty"), TextBox).Text
            txt_lot_no = DirectCast(rp.FindControl("Lot"), TextBox).Text
            start = DirectCast(rp.FindControl("strt"), TextBox).Text
            'finish = DirectCast(rp.FindControl("fnsh"), TextBox).Text
            line = DirectCast(rp.FindControl("line"), TextBox).Text
            remark = DirectCast(rp.FindControl("rmk"), TextBox).Text

            orderKey = DirectCast(rp.FindControl("O_K"), Label).Text

            Dim dtRows() As DataRow = dtTable.Select("txt_order_key = '" & orderKey & "'")

            If IsNumeric(qty) Then
                dtRows(0).Item("planned_production_qty") = qty
            Else
                errorMsg.Append("planned_production_qty is illegal number. by referrence to txt_order_key : " & orderKey)
                Exit For
            End If

            If Len(txt_lot_no) = 8 OrElse Len(txt_lot_no) = 0 Then
                dtRows(0).Item("txt_lot_no") = txt_lot_no
            Else
                errorMsg.Append("txt_lot_no is illegal. by referrence to txt_order_key : " & orderKey)
                Exit For
            End If

            If IsNumeric(line) Then
                dtRows(0).Item("int_line_no") = line
            Else
                errorMsg.Append("Line number is illegal. by referrence to txt_order_key : " & orderKey)
                Exit For
            End If


            If IsDate(start) Then
                dtRows(0).Item("dat_start_date") = CDate(start)
                'dtRows(0).Item("dat_finish_date") = CDate(finish)
            Else
                errorMsg.Append("Start date is illegal. by referrence to txt_order_key : " & orderKey)
                Exit For
            End If

            If Len(remark) < 150 Then
                dtRows(0).Item("txt_remark") = remark
            Else
                errorMsg.Append("Too many characters in remark. by referrence to txt_order_key : " & orderKey)
                Exit For
            End If


        Next


        ''update back to table Esch_Sh_tbl_orders
        'For Each a As DataRow In dtTableFrom0.Rows
        '    Dim rows() As DataRow = dtTable.Select("txt_order_key = '" & a.Item("txt_order_key") & "'")
        '    If rows.Count > 0 Then
        '        a.Item("dat_start_date") = rows(0).Item("dat_start_date")
        '        a.Item("dat_finish_date") = rows(0).Item("dat_finish_date")
        '        a.Item("planned_production_qty") = rows(0).Item("planned_production_qty")
        '        a.Item("txt_lot_no") = rows(0).Item("txt_lot_no")
        '        a.Item("int_line_no") = rows(0).Item("int_line_no")
        '    End If
        'Next

        'update back to table Esch_Sh_tbl_orders
        For Each sm As DataRow In dtTable.Select(Nothing, Nothing, DataViewRowState.ModifiedCurrent)
            Dim rows() As DataRow = dtTableFrom0.Select("txt_order_key = '" & sm.Item("txt_order_key") & "'")
            If rows.Count > 0 Then
                rows(0).Item("dat_start_date") = sm.Item("dat_start_date")
                'rows(0).Item("dat_finish_date") = sm.Item("dat_finish_date")
                rows(0).Item("planned_production_qty") = sm.Item("planned_production_qty")
                rows(0).Item("txt_lot_no") = sm.Item("txt_lot_no")
                rows(0).Item("int_line_no") = sm.Item("int_line_no")
                rows(0).Item("txt_remark") = sm.Item("txt_remark")
            End If
        Next



        Dim changedRows() As DataRow = dtTableFrom0.Select("int_line_no <> " & CInt(valueOf("intDummyLine")), Nothing, System.Data.DataViewRowState.ModifiedCurrent)
        Dim hasException As Boolean = False
        finishTime_exPlantDate_Span1(conn, changedRows, hasException)

        If hasException Then
            CacheInsert("Oexception", 1)
        End If

        additionDaysOnExplantDate(conn, dtTableFrom0, DataViewRowState.ModifiedCurrent)


        'if no error happens, 
        If errorMsg.Length > 0 Then
            msgPopUP(errorMsg.ToString(), Message)
        Else
            dtAdapter.Update(dtTable)
            dtFrom0.Update(dtTableFrom0)
            msgPopUP("Data updated back to DB.(" & Now.ToString & ")", Message, False, False)
            Message.ForeColor = Drawing.Color.Blue
        End If


        dtTable.Dispose()
        dtAdapter.Dispose()

        dtTableFrom0.Dispose()
        dtFrom0.Dispose()


        conn.Dispose()

        If errorMsg.Length = 0 Then
            CacheRemove("combineWB")
        End If

        displayButtonStatus()

        If String.IsNullOrEmpty(userName) Then
            unlockKeyTable(priority.CombineOrBatchCreation)
        End If

    End Sub



    Protected Sub cbAllLines_CheckedChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles cbAllLines.CheckedChanged
        Dim checked As Boolean = cbAllLines.Checked
        For Each a As CheckBox In linesCollection.Controls
            a.Checked = checked
        Next
    End Sub


    Protected Sub prdctnOrdrs2_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles prdctnOrdrs2.Click
        If Not IsDate(earlierTime.Text) Then
            earlierTime.ForeColor = Drawing.Color.Red
            msgPopUP("Illegal date time!", Message)
            Return
        End If


        If Not IsDate(laterTime.Text) Then
            laterTime.ForeColor = Drawing.Color.Red
            msgPopUP("Illegal date time!", Message)
            Return
        End If

        If Not IsNumeric(gapRSD.Text) OrElse CInt(gapRSD.Text) < 0 Then
            gapRSD.ForeColor = Drawing.Color.Red
            msgPopUP("Illegal number(Can not be less than 0)!", Message)
            Return
        End If

        Dim firstDate As DateTime = CDate(earlierTime.Text).AddHours(ddlHour1.SelectedIndex).AddMinutes(ddlMinute1.SelectedIndex)
        Dim secondDate As DateTime = CDate(laterTime.Text).AddHours(ddlHour2.SelectedIndex).AddMinutes(ddlMinute2.SelectedIndex)

        If firstDate.CompareTo(secondDate) > 0 Then
            secondDate = CDate(earlierTime.Text).AddHours(ddlHour1.SelectedIndex).AddMinutes(ddlMinute1.SelectedIndex)
            firstDate = CDate(laterTime.Text).AddHours(ddlHour2.SelectedIndex).AddMinutes(ddlMinute2.SelectedIndex)
        End If


        'delete all the records in the table
        Dim connstr As String = ConfigurationManager.ConnectionStrings(dbConnectionName).ConnectionString
        Dim conn As SqlConnection = New SqlConnection(connstr)



        Dim continueToUpdate As Boolean = True

        Dim dtUpdateTo As SqlDataAdapter = New SqlDataAdapter("SELECT * FROM Esch_Sh_tbl_similar_item_combination", conn)
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

        'where clause includes one part; ==================
        'one is all the selected lines and orders' start time is in between the two dates and planned_production_qty > 0
        Dim sqlWhereClause As StringBuilder = New StringBuilder(" WHERE  ")
        sqlWhereClause.Append("( CAST(int_line_no AS VARCHAR(5)) In (")
        Dim noLineWasSelected As Boolean = True 'if none of production lines was selected
        For Each a As CheckBox In linesCollection.Controls
            If a.Checked Then
                sqlWhereClause.Append("'" & a.Text & "',")
                noLineWasSelected = False
            End If
        Next

        If noLineWasSelected Then
            sqlWhereClause.Append("'-1') )") 'assign a false production line  number to the field to delete all the records in table Esch_Sh_tbl_BatchNO
        Else
            sqlWhereClause.Remove(sqlWhereClause.Length - 1, 1).Append(") ")
            sqlWhereClause.Append("  And (" & ddlColumn.SelectedValue & " Between " & dateSeparator & firstDate & dateSeparator & " and " & dateSeparator & secondDate & dateSeparator & ") And (planned_production_qty > 0)) ")
        End If


        'where clause includes one part; ==================



        Dim dtFrom0 As SqlDataAdapter = New SqlDataAdapter("SELECT txt_item_no,planned_production_qty,int_line_no,txt_lot_no,dat_start_date,dat_finish_date,txt_currency,dat_etd,dat_rdd,txt_remark,txt_grade,txt_color,txt_order_key,int_status_key FROM Esch_Sh_tbl_orders " & sqlWhereClause.ToString(), conn)
        Dim dtTableFrom0 As DataTable = New DataTable()
        dtFrom0.Fill(dtTableFrom0)
        'how many records to be inserted into Esch_Sh_tbl_similar_item_combination
        Dim recordsCount As Integer = 0


        Dim connParam As SqlConnection = New SqlConnection(ConfigurationManager.ConnectionStrings(dbConnForParam).ConnectionString)
        Dim dtFrom1 As SqlDataAdapter = New SqlDataAdapter("SELECT * FROM Esch_Sh_tbl_similar_grade_for_combination", connParam)
        Dim dtTableFrom1 As DataTable = New DataTable()
        dtFrom1.Fill(dtTableFrom1)



        For Each dtrow0 As DataRow In dtTableFrom0.Select(Nothing, "[txt_item_no] ASC, [dat_etd] ASC") 'get order line data from table Esch_Sh_tbl_orders
            Dim newRow As DataRow = dtTableUpdateTo.NewRow()

            Dim itemNo As String = dtrow0.Item("txt_item_no")
            Dim gradeNo As String = dtrow0.Item("txt_grade")

            newRow.Item("txt_item_no") = itemNo
            newRow.Item("planned_production_qty") = dtrow0.Item("planned_production_qty")
            newRow.Item("int_line_no") = dtrow0.Item("int_line_no")
            newRow.Item("txt_lot_no") = dtrow0.Item("txt_lot_no")
            newRow.Item("dat_start_date") = dtrow0.Item("dat_start_date")
            newRow.Item("dat_finish_date") = dtrow0.Item("dat_finish_date")
            newRow.Item("int_status_key") = dtrow0.Item("int_status_key")
            newRow.Item("txt_remark") = dtrow0.Item("txt_remark")
            'newRow.Item("txt_payment_term") = dtrow0.Item("txt_payment_term")
            newRow.Item("txt_currency") = dtrow0.Item("txt_currency")
            'newRow.Item("int_span") = dtrow0.Item("int_span")
            newRow.Item("dat_etd") = dtrow0.Item("dat_etd")
            newRow.Item("dat_rdd") = dtrow0.Item("dat_rdd")
            'newRow.Item("txt_remark") = dtrow0.Item("txt_remark")
            newRow.Item("txt_grade") = gradeNo
            'look up for typical grade 
            Dim typicalGrade() As DataRow = dtTableFrom1.Select("grade = '" & gradeNo & "'")
            If typicalGrade.Count > 0 Then
                newRow.Item("txt_item_no_similar") = itemNo.Replace(gradeNo & "-", typicalGrade(0).Item("typicalGrade").ToString & "-")
            Else
                newRow.Item("txt_item_no_similar") = itemNo
            End If
            newRow.Item("txt_order_key") = dtrow0.Item("txt_order_key")


            dtTableUpdateTo.Rows.Add(newRow)

        Next



        Dim newrow1() As DataRow = dtTableUpdateTo.Select(Nothing, "txt_item_no_similar ASC,txt_item_no ASC,dat_etd ASC, dat_start_date ASC , int_line_no ASC ")
        'currentOne = newrow1.Item("txt_item_no_similar")

        'for one item no, if we do not find at least two lines in the table, we delete the order line because there is no chance to combine with other order
        If newrow1.Count > 2 Then
            Dim currentOne As String = newrow1(1).Item("txt_item_no_similar")
            If currentOne = newrow1(0).Item("txt_item_no_similar") Then
            Else
                newrow1(0).Item("txt_combinationFlag") = "D" ' marked as to be deleted
            End If


            For i As Integer = 2 To newrow1.Count - 1
                currentOne = newrow1(i - 1).Item("txt_item_no_similar")
                If (currentOne = newrow1(i).Item("txt_item_no_similar")) OrElse (currentOne = newrow1(i - 2).Item("txt_item_no_similar")) Then
                Else
                    newrow1(i - 1).Item("txt_combinationFlag") = "D" ' marked as to be deleted
                End If
            Next


            If currentOne = newrow1(newrow1.Count - 1).Item("txt_item_no_similar") Then
            Else
                newrow1(newrow1.Count - 1).Item("txt_combinationFlag") = "D" ' marked as to be deleted
            End If




        Else
            msgPopUP("The number of orders is very small.", Message)
        End If


        For Each a In dtTableUpdateTo.Select("txt_combinationFlag = 'D'")
            a.Delete()
        Next


        'Only hold those orders with the gap of RSD is less than 7 and their production line is not same or  start time is not close
        Dim newlyAddedRows() As DataRow = dtTableUpdateTo.Select(Nothing, " txt_item_no_similar ASC,txt_item_no ASC, dat_etd ASC, dat_start_date ASC ,int_line_no ASC ", DataViewRowState.Added)
        Dim count2 As Integer = newlyAddedRows.Count
        If count2 > 2 Then

            For i = 0 To (count2 - 2)
                If newlyAddedRows(i).Item("txt_item_no_similar") = newlyAddedRows(i + 1).Item("txt_item_no_similar") OrElse newlyAddedRows(i).Item("txt_item_no") = newlyAddedRows(i + 1).Item("txt_item_no") Then
                    If (Math.Abs(DateDiff(DateInterval.Day, newlyAddedRows(i + 1).Item("dat_etd"), newlyAddedRows(i).Item("dat_etd"))) > CInt(gapRSD.Text)) OrElse
                         (newlyAddedRows(i).Item("int_line_no") = newlyAddedRows(i + 1).Item("int_line_no") AndAlso
                        (Math.Abs(DateDiff(DateInterval.Day, newlyAddedRows(i + 1).Item("dat_start_date"), newlyAddedRows(i).Item("dat_finish_date"))) <= 2 OrElse Math.Abs(DateDiff(DateInterval.Day, newlyAddedRows(i + 1).Item("dat_start_date"), newlyAddedRows(i).Item("dat_start_date"))) <= 2)) Then
                        If Not (newlyAddedRows(i).Item("txt_combinationFlag").ToString = "K") Then
                            newlyAddedRows(i).Item("txt_combinationFlag") = "D"
                        End If

                        If Not (newlyAddedRows(i + 1).Item("txt_combinationFlag").ToString = "K") Then
                            newlyAddedRows(i + 1).Item("txt_combinationFlag") = "D"
                        End If
                    Else
                        'newlyAddedRows(i).Item("txt_item_no_similar") = newlyAddedRows(i + 1).Item("txt_item_no_similar")
                        newlyAddedRows(i).Item("txt_combinationFlag") = "K"
                        newlyAddedRows(i + 1).Item("txt_combinationFlag") = "K"
                    End If

                End If
            Next


        End If


        For Each a In dtTableUpdateTo.Select("txt_combinationFlag = 'D'")
            a.Delete()
        Next

        If True Then
            Dim newrow2() As DataRow = dtTableUpdateTo.Select(Nothing, "txt_item_no_similar ASC,txt_item_no ASC,dat_etd ASC, dat_start_date ASC , int_line_no ASC ")
            'currentOne = newrow2.Item("txt_item_no_similar")

            'for one item no, if we do not find at least two lines in the table, we delete the order line because there is no chance to combine with other order
            If newrow2.Count > 2 Then
                Dim currentOne As String = newrow2(1).Item("txt_item_no_similar")
                If currentOne = newrow2(0).Item("txt_item_no_similar") Then
                Else
                    newrow2(0).Item("txt_combinationFlag") = "D" ' marked as to be deleted
                End If


                For i As Integer = 2 To newrow2.Count - 1
                    currentOne = newrow2(i - 1).Item("txt_item_no_similar")
                    If (currentOne = newrow2(i).Item("txt_item_no_similar")) OrElse (currentOne = newrow2(i - 2).Item("txt_item_no_similar")) Then
                    Else
                        newrow2(i - 1).Item("txt_combinationFlag") = "D" ' marked as to be deleted
                    End If
                Next


                If currentOne = newrow2(newrow2.Count - 1).Item("txt_item_no_similar") Then
                Else
                    newrow2(newrow2.Count - 1).Item("txt_combinationFlag") = "D" ' marked as to be deleted
                End If




            Else
                msgPopUP("The number of orders is very small.", Message)
            End If


            For Each a In dtTableUpdateTo.Select("txt_combinationFlag = 'D'")
                a.Delete()
            Next

        End If

        recordsCount = dtTableUpdateTo.Select(Nothing).Count
        'write data back to table Esch_Sh_tbl_BatchNO
        If continueToUpdate Then
            dtUpdateTo.Update(dtTableUpdateTo)
            msgPopUP("<div style='color:blue;font-size:larger;'>" & recordsCount & IIf(recordsCount > 1, " lines are", " line is") & " generated.(" & Now.ToString & ")</div>", Message, False, False)

        End If


        dtTableFrom0.Dispose()
        dtTableFrom1.Dispose()
        dtTableUpdateTo.Dispose()
        cmdbAccessCmdBuilder.Dispose()
        dtFrom0.Dispose()
        dtFrom1.Dispose()
        dtUpdateTo.Dispose()

        conn.Dispose()
        connParam.Close()

        SDS1.SelectCommand = "SELECT * FROM Esch_Sh_tbl_similar_item_combination ORDER BY txt_item_no_similar ASC,txt_item_no ASC,dat_etd ASC, dat_start_date ASC , int_line_no ASC "

        Rt1.DataBind()

        CacheInsert("combineWB", "enabled")
        displayButtonStatus()

    End Sub

End Class

