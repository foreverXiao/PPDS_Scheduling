Imports Microsoft.VisualBasic
Imports System.Data.SqlClient
Imports System.Data
Imports System.IO


Partial Public Class orderRelatedPlanning
    Inherits FrequentPlanActions


    ''' <summary>
    ''' 
    ''' </summary>
    ''' <param name="continued"></param>
    ''' <returns></returns>
    ''' <remarks>This is to check whether production line is set up correctly in table BatchNoGroupAndBatchRules.aspx,BatchNoGroupAndBatchRules.aspx and currentLeadtime.aspx</remarks>
    Public Function preCheckBeforeOrdersImportation(ByRef continued As Boolean) As String
        Dim rrtMessage As New StringBuilder()

        continued = True
        Dim arrayLine1 As Integer() = arrayOfLines()


        If True Then
            Dim compString As New StringBuilder("@")

            Dim connstr As String = ConfigurationManager.ConnectionStrings(dbConnectionName).ConnectionString
            Dim conn As SqlConnection = New SqlConnection(connstr)
            conn.Open()

            Dim command As New SqlCommand("Select distinct Production_line From Esch_Na_tbl_Lead_Time", conn)
            Dim reader As SqlDataReader
            reader = command.ExecuteReader()
            Do While reader.Read()
                compString.Append(reader("Production_line") & "@")
            Loop

            Dim cmpSTR As String = compString.ToString

            For i1 = 0 To arrayLine1.Count - 1
                If cmpSTR.IndexOf("@" & arrayLine1(i1) & "@") < 0 Then
                    continued = False
                    rrtMessage.Append("Production line ( <span style='color:red;'>" & arrayLine1(i1) & "</span> ) has not been set up in <a  href='../SCMrelated/currentLeadtime.aspx'> table Esch_Na_tbl_Lead_Time.<a/>")
                    Exit For
                End If
            Next
            reader.Close()


            command.CommandText = "Select distinct int_line_no From Esch_Na_tbl_orders"
            reader = command.ExecuteReader()
            compString.Clear()
            compString.Append("@")
            For i1 = 0 To arrayLine1.Count - 1
                compString.Append(arrayLine1(i1) & "@")
            Next
            compString.Append(valueOf("intDummyLine") & "@")

            cmpSTR = compString.ToString
            Do While reader.Read()
                If cmpSTR.IndexOf("@" & reader("int_line_no") & "@") < 0 Then
                    continued = False
                    rrtMessage.Append("Production line (" & reader("int_line_no") & ") has not been set up in <a  href='../plansetting/PrdctnLinesAndBatchNoGroup.aspx'> table Esch_Na_tbl_output_by_line_only.</a>")
                    Exit Do
                End If
            Loop


            For i1 = 0 To arrayLine1.Count - 1

            Next
            reader.Close()



            command.Dispose()
            conn.Close()
            conn.Dispose()
        End If

        If continued Then

            Dim connParam As SqlConnection = New SqlConnection(ConfigurationManager.ConnectionStrings(dbConnForParam).ConnectionString)
            connParam.Open()
            Dim commandP As New SqlCommand("Select distinct int_line_no From Esch_Na_tbl_LinesAndOwners", connParam)
            Dim readerP As SqlDataReader
            readerP = commandP.ExecuteReader()

            Dim compStringP As New StringBuilder("@")
            Do While readerP.Read()
                compStringP.Append(readerP("int_line_no") & "@")
            Loop

            Dim cmpSTRP As String = compStringP.ToString
            For i1 = 0 To arrayLine1.Count - 1
                If cmpSTRP.IndexOf("@" & arrayLine1(i1) & "@") < 0 Then
                    continued = False
                    rrtMessage.Append("No production line (" & arrayLine1(i1) & ") has not been set up in <a  href='../plansetting/LinesAndOwners.aspx'> table Esch_Na_tbl_LinesAndOwners</a>")
                    Exit For
                End If
            Next
            readerP.Close()

            commandP.Dispose()

            connParam.Close()
            connParam.Dispose()

        End If



        Return rrtMessage.ToString


    End Function





    'further pass newly imported order data to main order table, insert new order and update old order and cancel orders. 
    ' Esch_Na_Esch_Na_tbl_orders_new_revision_from_OPM ===> Esch_Na_tbl_orders

    '''<summary>
    '''to update order status (int_status_key),whether new order, or revision order with RDD changed,order quantity changed,POD changed or shipment method changed
    ''' int_status_key 0 means new or revision,-1 means cancelled,20 means invoiced, 10 means picked
    '''</summary>
    Public Sub NewRevisionOrderDataToMainOrderTable(ByRef conn As SqlConnection, ByRef msg As String)

        Dim sqlSelect As System.Text.StringBuilder = New System.Text.StringBuilder()
        Dim command As New SqlCommand("SELECT txtFieldName FROM Esch_Na_tbl_interface_mapping WHERE intImportOrderMapping > 0 ORDER BY intImportOrderMapping", conn)
        Dim reader As SqlDataReader = command.ExecuteReader()
        'Dim mapping1() As String = {String.Empty}, icount As Integer = 0
        While reader.Read()
            'ReDim Preserve mapping1(icount)
            'mapping1(icount) = reader("txtFieldName")
            'icount += 1
            sqlSelect.Append(reader("txtFieldName") & " ,")
        End While
        reader.Close()


        sqlSelect.Append(" txt_order_key , txt_grade , txt_color ")
        Dim dtUpdateFrom As SqlDataAdapter = New SqlDataAdapter("SELECT " & sqlSelect.ToString() & " FROM Esch_Na_Esch_Na_tbl_orders_new_revision_from_OPM", conn)

        sqlSelect.Append(" , txt_order_status,planned_production_qty,dat_finish_date,txt_order_type ")
        Dim dtUpdateTo As SqlDataAdapter = New SqlDataAdapter("SELECT " & sqlSelect.ToString() & " FROM Esch_Na_tbl_orders ", conn)
        Dim cmdbAccessCmdBuilder As New SqlCommandBuilder(dtUpdateTo)
        dtUpdateTo.UpdateCommand = cmdbAccessCmdBuilder.GetUpdateCommand()
        dtUpdateTo.InsertCommand = cmdbAccessCmdBuilder.GetInsertCommand()
        dtUpdateTo.DeleteCommand = cmdbAccessCmdBuilder.GetDeleteCommand()

        Dim dsAccess As DataSet = New DataSet


        dtUpdateTo.Fill(dsAccess, "UpdateTo")
        Dim keys(0) As DataColumn
        keys(0) = dsAccess.Tables("UpdateTo").Columns("txt_order_key")
        dsAccess.Tables("UpdateTo").PrimaryKey = keys




        dtUpdateFrom.FillLoadOption = LoadOption.Upsert
        dtUpdateFrom.Fill(dsAccess, "UpdateTo")


        Dim status1 As System.Text.StringBuilder = New System.Text.StringBuilder()

        'make changes for those cancelled or shipped orders (order status from OPM: -1 means cancelled,20 means invoiced )
        Dim cancelledRws As DataRow() = dsAccess.Tables("UpdateTo").Select("int_status_key = '-1'", Nothing, DataViewRowState.ModifiedCurrent)
        For Each a As DataRow In cancelledRws
            a.Item("int_status_key") = "cancelled"
            a.Item("txt_order_status") = "cancelled"
        Next

        Dim invoicedRws As DataRow() = dsAccess.Tables("UpdateTo").Select("int_status_key = '20'", Nothing, DataViewRowState.ModifiedCurrent)
        For Each a As DataRow In invoicedRws
            a.Item("int_status_key") = "invoiced"
            a.Item("txt_order_status") = "invoiced"
        Next

        'make changes for those revised orders (order status from OPM: 0 means changed , 10 means picked )
        Dim revRws As DataRow() = dsAccess.Tables("UpdateTo").Select("(int_status_key = '0') Or (int_status_key = '10')", Nothing, DataViewRowState.ModifiedCurrent)
        Dim frequentlyUsedStatus As String = String.Empty
        For Each a As DataRow In revRws
            status1.Clear()
            status1.Append("REV")

            If a.Item("dat_rdd", DataRowVersion.Current) <> a.Item("dat_rdd", DataRowVersion.Original) Then
                status1.Append("-R")
            End If

            If a.Item("flt_order_qty", DataRowVersion.Current) <> a.Item("flt_order_qty", DataRowVersion.Original) Then
                status1.Append("-Q")
            End If

            'If a.Item("flt_unallocate_qty", DataRowVersion.Current) <> a.Item("flt_unallocate_qty", DataRowVersion.Original) Then
            '    status1.Append("-U")
            'End If



            If status1.Length > 4 Then  ' Like 'REV-%' ' no change on RDD or Quantity
                a.Item("int_status_key") = status1.ToString() & "2"  'you have two chances to change start time
            Else
                frequentlyUsedStatus = a.Item("int_status_key", DataRowVersion.Original).ToString
                If frequentlyUsedStatus.IndexOf("old") = 0 OrElse frequentlyUsedStatus.ToUpper = "REV" Then
                    a.Item("int_status_key") = frequentlyUsedStatus
                Else
                    a.Item("int_status_key") = "old_" & frequentlyUsedStatus
                End If

            End If


        Next


        'do something for newly entered orders which not appear before this importings
        Dim newRws As DataRow() = dsAccess.Tables("UpdateTo").Select(Nothing, Nothing, DataViewRowState.Added)
        For Each a As DataRow In newRws
            If a.Item("int_status_key") = "0" Then
                a.Item("int_status_key") = newOrderStatus
                a.Item("txt_order_type") = "MTO"
            Else
                a.Delete() 'maybe newly added but cancelled or voided immediately 
            End If
        Next



        'do something for those old orders which have no changes since last working day, of course we don't handle those cancelled (but not invoiced) orders
        If True Then
            Dim oldRws As DataRow() = dsAccess.Tables("UpdateTo").Select("(int_status_key <> 'old')  And (int_status_key <> 'cancelled') And (int_status_key <> 'invoiced') ", Nothing, DataViewRowState.Unchanged)
            For Each a As DataRow In oldRws
                If a.Item("int_status_key").ToString.IndexOf("old") <> 0 Then
                    a.Item("int_status_key") = "old_" & a.Item("int_status_key")
                End If
            Next
        End If


        'write changed data back to access database
        Try

            dtUpdateTo.Update(dsAccess, "UpdateTo")

        Catch ex As Exception
            msg &= "<br /> NewRevisionOrderDataToMainOrderTable: " & ex.Message
        Finally
            dtUpdateTo.Dispose()
            dtUpdateFrom.Dispose()
            dsAccess.Dispose()
        End Try

    End Sub




    '''<summary>
    ''' assign production line for new orders,calculate their working time duration. according to three tables by quantity, by grade , by item
    ''' </summary>
    Public Sub assignLineToNewOrder(ByRef conn As SqlConnection, ByRef msg As String)

        Dim connParam As SqlConnection = New SqlConnection(ConfigurationManager.ConnectionStrings(dbConnForParam).ConnectionString)
        connParam.Open()

        Dim dtUpdateFrom0 As SqlDataAdapter = New SqlDataAdapter("SELECT * FROM Esch_Na_tbl_assign_line_by_order_qty", connParam)
        Dim dtUpdateFrom1 As SqlDataAdapter = New SqlDataAdapter("SELECT * FROM Esch_Na_tbl_assign_line_by_grade", connParam)
        Dim dtUpdateFrom2 As SqlDataAdapter = New SqlDataAdapter("SELECT * FROM Esch_Na_tbl_Item_And_line_no_Priority_Sequence", conn)
        Dim dtUpdateTo As SqlDataAdapter = New SqlDataAdapter("SELECT txt_order_key,int_line_no,dat_start_date,txt_item_no,txt_grade,planned_production_qty FROM Esch_Na_tbl_orders WHERE (int_status_key = '" & newOrderStatus & "')", conn)
        Dim cmdbAccessCmdBuilder As New SqlCommandBuilder(dtUpdateTo)
        dtUpdateTo.UpdateCommand = cmdbAccessCmdBuilder.GetUpdateCommand()
        dtUpdateTo.InsertCommand = cmdbAccessCmdBuilder.GetInsertCommand()
        Dim dsAccess As DataSet = New DataSet


        dtUpdateTo.Fill(dsAccess, "UpdateTo")
        Dim keys(0) As DataColumn
        keys(0) = dsAccess.Tables("UpdateTo").Columns("txt_order_key")
        dsAccess.Tables("UpdateTo").PrimaryKey = keys

        dtUpdateFrom0.Fill(dsAccess, "by_qty")
        dtUpdateFrom1.Fill(dsAccess, "by_grade")
        dtUpdateFrom2.Fill(dsAccess, "by_item")

        Dim dummyLine As String = valueOf("intDummyLine") 'dummy line
        Dim veryLateDateForDummyLine As DateTime = CDate(valueOf("datDummyVeryLateStartDate"))

        Dim byQtyLineArray() As String, byGradeLineArray() As String, byItemLineArray() As String, finalChoices() As String
        Dim match1 As Boolean, match2 As Boolean
        Dim lineMatch2 As StringBuilder = New StringBuilder()
        Dim lineMatchAll As StringBuilder = New StringBuilder()
        Dim random1 As New Random(), rndInt As Integer
        For Each a As DataRow In dsAccess.Tables("UpdateTo").Rows
            byQtyLineArray = Nothing
            byGradeLineArray = Nothing
            byItemLineArray = Nothing
            finalChoices = Nothing
            lineMatch2.Clear()
            lineMatchAll.Clear()


            Dim byQty() As DataRow = dsAccess.Tables("by_qty").Select(" maximum >= " & a.Item("planned_production_qty") & " And minimum <= " & a.Item("planned_production_qty"))
            If byQty.Count > 0 Then
                byQtyLineArray = byQty(0).Item("txt_line_no").ToString().Trim().Split(New Char() {","}, StringSplitOptions.RemoveEmptyEntries)
            Else
                msg &= " <div style = 'color:red;'> assignLineToNewOrder:  no line group by order's quantity for this " & a.Item("planned_production_qty") & " kg (txt_order_key:" & a.Item("txt_order_key") & ")</div>"
                Continue For
            End If
            Dim byGrade() As DataRow = dsAccess.Tables("by_grade").Select("txt_grade = '" & a.Item("txt_grade") & "'")
            If byGrade.Count > 0 Then byGradeLineArray = byGrade(0).Item("txt_line_no_group").ToString().Split(New Char() {","}, StringSplitOptions.RemoveEmptyEntries)
            Dim byItem() As DataRow = dsAccess.Tables("by_item").Select("txtItem = '" & a.Item("txt_item_no") & "'")
            If byItem.Count > 0 Then byItemLineArray = byItem(0).Item("txtLineNoSequence").ToString().Split(New Char() {","}, StringSplitOptions.RemoveEmptyEntries)

            ''decide which line is most suitable for the order based on the information provided by three conditions
            'If byQtyLineArray.Count > 1 Then
            '    For i As Integer = 0 To (byQtyLineArray.Count - 1)
            '        match1 = False
            '        match2 = False


            '        If byGradeLineArray IsNot Nothing Then
            '            For Each j In byGradeLineArray
            '                If byQtyLineArray(i) = j Then  ' also be found in byGrade
            '                    match1 = True
            '                    lineMatch1.Append(byQtyLineArray(i) & ",")
            '                    Exit For
            '                End If
            '            Next
            '        End If


            '        If byItemLineArray IsNot Nothing Then
            '            For Each k In byItemLineArray 'also be found in byItem
            '                If byQtyLineArray(i) = k Then
            '                    match2 = True
            '                    Exit For
            '                End If
            '            Next
            '        End If

            '        If match1 AndAlso match2 Then lineMatchAll.Append(byQtyLineArray(i) & ",")

            '    Next


            '    If Not String.IsNullOrEmpty(lineMatchAll.ToString()) Then
            '        finalChoices = lineMatchAll.ToString().Split(New Char() {","}, StringSplitOptions.RemoveEmptyEntries) ' valid line numbers  be found in byQty and byGrade and byItem
            '    Else
            '        If Not String.IsNullOrEmpty(lineMatch1.ToString()) Then
            '            finalChoices = lineMatch1.ToString().Split(New Char() {","}, StringSplitOptions.RemoveEmptyEntries)  ' valid line numbers only be found in byQty and byGrade
            '        Else
            '            finalChoices = byQtyLineArray ' valid line numbers only be found in byQty
            '        End If
            '    End If

            'Else
            '    finalChoices = byQtyLineArray  'only have one choice
            'End If


            ''randomly choose one line if there are multiple line available at the same time, otherwise you choose the only choice
            'If finalChoices.Count > 1 Then
            '    rndInt = CInt(random1.Next(finalChoices.Count - 1))
            'Else
            '    rndInt = 0
            'End If



            'decide which line is most suitable for the order based on the information provided by three conditions
            If byItemLineArray IsNot Nothing Then
                For i As Integer = 0 To (byItemLineArray.Count - 1)
                    match1 = False
                    match2 = False

                    If byGradeLineArray IsNot Nothing Then
                        For Each j In byGradeLineArray
                            If byItemLineArray(i) = j Then  ' also be found in byGrade
                                match1 = True
                                Exit For
                            End If
                        Next
                    End If


                    If byQtyLineArray IsNot Nothing Then
                        For Each k In byQtyLineArray 'also be found in by_qty
                            If byItemLineArray(i) = k Then
                                match2 = True
                                Exit For
                            End If
                        Next
                    End If

                    If match1 AndAlso match2 Then
                        lineMatchAll.Append(byItemLineArray(i) & ",")
                        Exit For
                    End If

                Next


                If Not String.IsNullOrEmpty(lineMatchAll.ToString()) Then
                    finalChoices = lineMatchAll.ToString().Split(New Char() {","}, StringSplitOptions.RemoveEmptyEntries) ' valid line numbers  be found in byQty and byGrade and byItem
                Else
                    finalChoices = byQtyLineArray ' valid line numbers only be found in byQty
                End If

            Else

                If byGradeLineArray IsNot Nothing Then
                    For Each j In byGradeLineArray
                        If byQtyLineArray IsNot Nothing Then
                            For Each k In byQtyLineArray 'also be found in by_qty
                                If j = k Then
                                    lineMatch2.Append(k & ",")
                                    Exit For
                                End If
                            Next
                        End If
                    Next
                End If

                If Not String.IsNullOrEmpty(lineMatch2.ToString()) Then
                    finalChoices = lineMatch2.ToString().Split(New Char() {","}, StringSplitOptions.RemoveEmptyEntries)
                Else
                    finalChoices = byQtyLineArray ' valid line numbers only be found in byQty
                End If

            End If


            'randomly choose one line if there are multiple line available at the same time, otherwise you choose the only choice
            If finalChoices.Count > 1 Then
                rndInt = CInt(random1.Next(finalChoices.Count - 1))
            Else
                rndInt = 0
            End If



            a.Item("int_line_no") = finalChoices(rndInt)
            'if order is arranged to a dummy production line, then give a very late start time to it to means pending this order
            If String.Equals(finalChoices(rndInt), dummyLine) Then a.Item("dat_start_date") = veryLateDateForDummyLine

        Next



        Try
            dtUpdateTo.Update(dsAccess, "UpdateTo")  'resolve changes back to database
        Catch ex As Exception
            msg &= " <div style = 'color:red;'> assignLineToNewOrder: " & ex.Message & "</div>"
        End Try


        cmdbAccessCmdBuilder.Dispose()

        dsAccess.Dispose()
        dtUpdateTo.Dispose()
        dtUpdateFrom0.Dispose()
        dtUpdateFrom1.Dispose()
        dtUpdateFrom2.Dispose()

        connParam.Close()
        connParam.Dispose()




    End Sub





    ''' <summary>
    ''' Handle those new orders violating MOQ rules, put them at dummy production line and give a faraway production day
    ''' </summary>
    Public Sub MOQappliedToNewOrder(ByRef conn As SqlConnection, ByRef msg As String)

        Dim connParam As SqlConnection = New SqlConnection(ConfigurationManager.ConnectionStrings(dbConnForParam).ConnectionString)
        connParam.Open()

        Dim dtUpdateFrom0 As SqlDataAdapter = New SqlDataAdapter("SELECT groupName,headerName,relationOperator,conditionValue,remarkToBeAdded FROM Esch_Na_tbl_MOQ", connParam)
        Dim dtUpdateFrom1 As SqlDataAdapter = New SqlDataAdapter("SELECT DISTINCT groupName FROM Esch_Na_tbl_MOQ", connParam)
        Dim dsAccess As DataSet = New DataSet

        dtUpdateFrom0.Fill(dsAccess, "MOQ")
        dtUpdateFrom1.Fill(dsAccess, "MOQcategory")

        Dim columnNames = From c In dsAccess.Tables("MOQ").AsEnumerable() Select c.Field(Of String)("headerName") Distinct
        Dim sqlSelect As StringBuilder = New StringBuilder()
        For Each i In columnNames
            sqlSelect.Append("," & i.ToString())
        Next


        Dim dtUpdateTo As SqlDataAdapter = New SqlDataAdapter("SELECT txt_order_key,int_line_no,dat_start_date,txt_remark " & sqlSelect.ToString & " FROM  Esch_Na_tbl_orders  WHERE (int_status_key = '" & newOrderStatus & "') ", conn)
        Dim cmdbAccessCmdBuilder As New SqlCommandBuilder(dtUpdateTo)
        dtUpdateTo.UpdateCommand = cmdbAccessCmdBuilder.GetUpdateCommand()
        dtUpdateTo.InsertCommand = cmdbAccessCmdBuilder.GetInsertCommand()



        dtUpdateTo.Fill(dsAccess, "UpdateTo")
        Dim keys(0) As DataColumn
        keys(0) = dsAccess.Tables("UpdateTo").Columns("txt_order_key")
        dsAccess.Tables("UpdateTo").PrimaryKey = keys





        Dim dummyLateDate As DateTime = CDate(valueOf("datDummyVeryLateStartDate")), dummyLine As Integer = CInt(valueOf("intDummyLine"))
        Dim MOQcondition As StringBuilder = New StringBuilder()
        Dim extraChar As String


        'for each category in MOQ table
        For Each a As DataRow In dsAccess.Tables("MOQcategory").Rows

            Dim b() As DataRow = dsAccess.Tables("MOQ").Select("groupName ='" & a.Item("groupName") & "'")
            MOQcondition.Clear()
            If b.Count > 0 Then
                For Each i As DataRow In b  'summarize up every condition to specific MOQ rule
                    extraChar = String.Empty

                    If i.Item("headerName").ToString().IndexOf("txt") >= 0 Then
                        extraChar = "'"
                    End If

                    If i.Item("headerName").ToString().IndexOf("dat") >= 0 Then
                        extraChar = dateSeparator
                    End If

                    MOQcondition.Append(i.Item("headerName") & " " & i.Item("relationOperator") & " " & extraChar & i.Item("conditionValue") & extraChar & " And ")
                Next
                MOQcondition.Remove(MOQcondition.Length - 4 - 1, 5)  'eliminate the last operator ' And '

                Dim c() As DataRow = dsAccess.Tables("UpdateTo").Select(MOQcondition.ToString())
                For Each d As DataRow In c
                    d.Item("int_line_no") = dummyLine
                    d.Item("dat_start_date") = dummyLateDate
                    'If d.Item("txt_remark").ToString().IndexOf(b(0).Item("remarkToBeAdded")) = -1 Then 'no MOQ remark has ever been added
                    d.Item("txt_remark") &= " " & b(0).Item("remarkToBeAdded")
                    'End If
                Next
            End If

        Next



        Try
            dtUpdateTo.Update(dsAccess, "UpdateTo")  'resolve changes back to database
        Catch ex As Exception
            msg &= "<div style = 'color:red;'> MOQappliedToNewOrder: " & ex.Message & "</div>"
        End Try

        cmdbAccessCmdBuilder.Dispose()
        dsAccess.Dispose()
        dtUpdateTo.Dispose()
        dtUpdateFrom0.Dispose()
        dtUpdateFrom1.Dispose()

        connParam.Close()
        connParam.Dispose()


    End Sub


    ''' <summary>
    ''' do some preparation for futher automatic scheduling, such as update standard leadtime for those VIP orders,
    ''' and update field txt_auxiliary_code for further proceedings
    ''' </summary>
    Public Sub Preparation_For_Automatic_scheduling(ByRef conn As SqlConnection, ByRef msg As String)

        Dim connParam As SqlConnection = New SqlConnection(ConfigurationManager.ConnectionStrings(dbConnForParam).ConnectionString)
        connParam.Open()

        Dim dtUpdateFrom0 As SqlDataAdapter = New SqlDataAdapter("SELECT * FROM Esch_Na_tbl_VIP_lead_time ORDER BY DaysOfCommittedLeadtime DESC,VIPgroup ASC", connParam)  'sort the data in order to make the shortest lead time VIP prioritized
        Dim dtUpdateFrom1 As SqlDataAdapter = New SqlDataAdapter("SELECT [VIPleadtimeTable].VIPgroup FROM (SELECT VIPgroup, avg(DaysOfCommittedLeadtime) AS leadtimeDesc FROM Esch_Na_tbl_VIP_lead_time GROUP BY VIPgroup)  AS VIPleadtimeTable ORDER BY [VIPleadtimeTable].leadtimeDesc DESC", connParam)
        Dim dsAccess As DataSet = New DataSet

        'get the list of column name to be used to identify VIP customer
        dtUpdateFrom0.Fill(dsAccess, "VIP")

        Dim columnNames = From c In dsAccess.Tables("VIP").AsEnumerable() Select c.Field(Of String)("headerName") Distinct
        Dim sqlSelect As StringBuilder = New StringBuilder()
        For Each i In columnNames
            sqlSelect.Append("," & i.ToString())
        Next


        'scope will cover those start time not earlier than yesterday and those new orders, all shipped or cancelled are excluded.
        Dim dtUpdateTo As SqlDataAdapter = New SqlDataAdapter("SELECT txt_order_key,txt_order_type,int_status_key,txt_VIP,txt_auxiliary_code,lng_VIP_lead_time,lng_AdvanceOfRevision " & sqlSelect.ToString() & " FROM  Esch_Na_tbl_orders  WHERE (int_status_key Not In ('invoiced','cancelled')) And ((dat_start_date Is Null) Or (dat_start_date > " & dateSeparator & DateTime.Today.AddDays(-1) & dateSeparator & "))", conn)
        Dim cmdbAccessCmdBuilder As New SqlCommandBuilder(dtUpdateTo)
        dtUpdateTo.UpdateCommand = cmdbAccessCmdBuilder.GetUpdateCommand()
        dtUpdateTo.InsertCommand = cmdbAccessCmdBuilder.GetInsertCommand()



        dtUpdateTo.Fill(dsAccess, "UpdateTo")
        Dim keys(0) As DataColumn
        keys(0) = dsAccess.Tables("UpdateTo").Columns("txt_order_key")
        dsAccess.Tables("UpdateTo").PrimaryKey = keys

        dtUpdateFrom1.Fill(dsAccess, "VIPgroup")


        Dim dummyLateDate As DateTime = CDate(valueOf("datDummyVeryLateStartDate")), dummyLine As Integer = CInt(valueOf("intDummyLine"))
        Dim VIPcondition As StringBuilder = New StringBuilder()
        Dim extraChar As String, VIPgroupName As String = String.Empty
        Dim VipLeadTime As Integer = 90, daysAdvancePerRevision As Integer = 90

        'for each group in VIP table
        For Each a As DataRow In dsAccess.Tables("VIPgroup").Rows
            VIPgroupName = a.Item("VIPgroup")
            Dim b() As DataRow = dsAccess.Tables("VIP").Select("VIPgroup ='" & VIPgroupName & "'")
            VIPcondition.Clear()
            If b.Count > 0 Then
                For Each i As DataRow In b  'summarize up every condition according to specific VIP rule
                    extraChar = String.Empty

                    If i.Item("headerName").ToString().IndexOf("txt") >= 0 Then 'use different VB.net synax when there is different type of data
                        extraChar = "'"
                    End If

                    If i.Item("headerName").ToString().IndexOf("dat") >= 0 Then
                        extraChar = dateSeparator
                    End If

                    VIPcondition.Append(i.Item("headerName") & " " & i.Item("Operator") & " " & extraChar & i.Item("VIPCondition") & extraChar & " Or ")
                Next
                VIPcondition.Remove(VIPcondition.Length - 3 - 1, 4)  'eliminate the last operator ' Or '

                VipLeadTime = CInt(b(0).Item("DaysOfCommittedLeadtime"))
                daysAdvancePerRevision = CInt(b(0).Item("DaysAdvanceBeforeRevision"))

                Dim c() As DataRow = dsAccess.Tables("UpdateTo").Select(VIPcondition.ToString())
                For Each d As DataRow In c
                    d.Item("txt_VIP") = VIPgroupName
                    d.Item("lng_VIP_lead_time") = VipLeadTime
                    d.Item("lng_AdvanceOfRevision") = daysAdvancePerRevision
                Next
            End If

        Next

        'update for all non VIP orders
        Dim nonVIP() As DataRow = dsAccess.Tables("UpdateTo").Select("txt_VIP Is Null")
        For Each d As DataRow In nonVIP
            d.Item("lng_VIP_lead_time") = 180
            d.Item("lng_AdvanceOfRevision") = 180
        Next


        'update auxiliary_code field to do preparation for automatic scheduling in the later steps for MTO orders
        Dim allMTO_Orders() As DataRow = dsAccess.Tables("UpdateTo").Select("txt_order_type = 'MTO'")
        For Each d As DataRow In allMTO_Orders
            If d.Item("int_status_key").ToString.IndexOf(newOrderStatus) = 0 Then  'if this is a NEW order
                d.Item("txt_auxiliary_code") = "NEW"
            Else  'RDD or quantity is changed, txt_auxiliary_code is marked as 'REV-' (not include the case that unallocate quantity is changed)
                If d.Item("int_status_key").ToString.IndexOf("REV-R") = 0 OrElse d.Item("int_status_key").ToString.IndexOf("REV-Q") = 0 Then
                    d.Item("txt_auxiliary_code") = "REV-"
                Else
                    d.Item("txt_auxiliary_code") = "OLD"
                End If
            End If
        Next


        'update auxiliary_code field to do preparation for automatic scheduling in the later steps for MTI orders
        Dim allMTI_Orders() As DataRow = dsAccess.Tables("UpdateTo").Select("txt_order_type = 'MTI'")
        For Each d As DataRow In allMTI_Orders
            d.Item("txt_auxiliary_code") = "OLD"
        Next



        Try
            dtUpdateTo.Update(dsAccess, "UpdateTo")  'resolve changes back to database
        Catch ex As Exception
            msg &= "<div style = 'color:red;'> Preparation_For_Automatic_scheduling: " & ex.Message & "</div>"
        Finally
            dsAccess.Dispose()
            dtUpdateTo.Dispose()
            dtUpdateFrom0.Dispose()
            dtUpdateFrom1.Dispose()

            connParam.Close()
            connParam.Dispose()

        End Try
    End Sub


    ''' <summary>
    ''' return the number representing the reserved capacity for VIP customers every week
    ''' </summary>
    Public Function reservedCapForVIP(ByRef conn As SqlConnection) As Long


        reservedCapForVIP = 0

        If CBool(valueOf("bnlReserveCapForVIP")) Then  'if you decide to reserve capacity for VIP customers

            Dim connParam As SqlConnection = New SqlConnection(ConfigurationManager.ConnectionStrings(dbConnForParam).ConnectionString)
            connParam.Open()

            Dim command As New SqlCommand("SELECT sum(ReservedCap) AS sumReservedCap FROM (SELECT Avg(ReservedCapPerWeek) as ReservedCap  From Esch_Na_tbl_VIP_lead_time GROUP BY VIPgroup)", conn)
            Dim reader As SqlDataReader
            reader = command.ExecuteReader()
            If reader.Read() Then reservedCapForVIP = CLng(reader("sumReservedCap"))
            reader.Close()
            command.Dispose()

            connParam.Close()
            connParam.Dispose()

        End If

    End Function


    ''' <summary>
    ''' Use to calculate leadtime per line per line group based on the setting in the table Esch_Na_tbl_Lead_Time
    ''' </summary>
    Public Sub calculateLeadtime(ByRef conn As SqlConnection, ByRef msg As String)

        'scope will cover those start time not earlier than now and those new orders and txt_auxiliary_code is Null, all shipped or cancelled are excluded.
        Dim dtUpdateTo As SqlDataAdapter = New SqlDataAdapter("SELECT txt_order_key,int_line_no,dat_start_date,txt_auxiliary_code,int_status_key,txt_VIP,lng_VIP_lead_time,lng_AdvanceOfRevision FROM  Esch_Na_tbl_orders  WHERE (int_line_no <> " & valueOf("intDummyLine") & " ) And (int_status_key Not In ('invoiced','cancelled')) And ((dat_start_date Is Null) Or (dat_start_date > " & dateSeparator & DateTime.Now & dateSeparator & "))", conn)
        Dim cmdbAccessCmdBuilder As New SqlCommandBuilder(dtUpdateTo)
        dtUpdateTo.UpdateCommand = cmdbAccessCmdBuilder.GetUpdateCommand()
        dtUpdateTo.InsertCommand = cmdbAccessCmdBuilder.GetInsertCommand()
        Dim dsAccess As DataSet = New DataSet

        Try
            dtUpdateTo.Fill(dsAccess, "UpdateTo")
            Dim keys(0) As DataColumn
            keys(0) = dsAccess.Tables("UpdateTo").Columns("txt_order_key")
            dsAccess.Tables("UpdateTo").PrimaryKey = keys

        Catch ex As Exception
            msg &= "<br /> calculateLeadtime: " & ex.Message
        End Try

        Dim dummyLateDate As DateTime = CDate(valueOf("datDummyVeryLateStartDate")), dummyLine As Integer = CInt(valueOf("intDummyLine"))
        Dim LeadtimeCondition As StringBuilder = New StringBuilder()
        Dim LeadtimeGroupName As String = String.Empty
        Dim dailyOutputByLeadtimeGroup As Long = 0, daysAdvancePerRevision As Integer = 90

        'for each group in leadtime table
        For Each a As DataRow In dsAccess.Tables("LeadtimeGroup").Rows
            LeadtimeGroupName = a.Item("Lead_time_Group")
            Dim b() As DataRow = dsAccess.Tables("Leadtime").Select("Lead_time_Group ='" & LeadtimeGroupName & "'")
            LeadtimeCondition.Clear()
            If b.Count > 0 Then
                For Each i As DataRow In b  'summarize up every condition according to specific lead time group rule
                    LeadtimeCondition.Append("int_line_no = " & i.Item("Production_line") & " Or ")
                Next
                LeadtimeCondition.Remove(LeadtimeCondition.Length - 3 - 1, 4)  'eliminate the last operator ' Or '

                'LeadTime = CInt(b(0).Item("DaysOfCommittedLeadtime"))
                daysAdvancePerRevision = CInt(b(0).Item("DaysAdvanceBeforeRevision"))

                Dim c() As DataRow = dsAccess.Tables("UpdateTo").Select(LeadtimeCondition.ToString())
                For Each d As DataRow In c

                    'd.Item("lng_VIP_lead_time") = VipLeadTime
                    d.Item("lng_AdvanceOfRevision") = daysAdvancePerRevision

                    If d.Item("int_status_key").ToString.IndexOf("REV") = -1 Then  'if this is a revision order
                        d.Item("txt_auxiliary_code") = d.Item("int_status_key")
                    Else
                        d.Item("txt_auxiliary_code") = "REV"
                    End If

                Next
            End If

        Next



        Try
            dtUpdateTo.Update(dsAccess, "UpdateTo")  'resolve changes back to database
        Catch ex As Exception
            msg &= "<div style = 'color:red;'> Calculate Leadtime: " & ex.Message & "</div>"
        End Try

        cmdbAccessCmdBuilder.Dispose()
        dsAccess.Dispose()
        dtUpdateTo.Dispose()





    End Sub

    ''' <summary>
    ''' set the value for the specific lead time variable  according to leadtime variable table
    ''' </summary>
    ''' <param name="name">the parameter's name</param>
    Public Sub setLTvalueAs(ByVal name As String, ByVal value As String)

        initiateCacheFromLeadtimeTable()

        Dim cacheDictnry As Dictionary(Of String, String) = CType(Cache("Leadtime"), Dictionary(Of String, String))

        If cacheDictnry.ContainsKey(name) Then
            cacheDictnry.Item(name) = value
            Cache.Insert("Leadtime", cacheDictnry, Nothing, Cache.NoAbsoluteExpiration, TimeSpan.FromMinutes(120))
        End If

    End Sub


    ''' <summary>
    ''' get the value for the specific lead time variable  according to leadtime variable table
    ''' </summary>
    ''' <param name="name">the parameter's name</param>
    ''' <returns>return a string,but you can convert the string to any type of data as you want</returns>
    Public Function LTvalueOf(ByVal name As String) As String

        initiateCacheFromLeadtimeTable()

        Dim cacheDictnry As Dictionary(Of String, String) = CType(Cache("Leadtime"), Dictionary(Of String, String))
        If cacheDictnry.ContainsKey(name) Then
            Return cacheDictnry(name)
        Else
            Return String.Empty
        End If

    End Function


    ''' <summary>
    ''' to initiate  cache variable: store data of production line, daily output, lead time,start time and coefficient per leadtime group
    ''' </summary>
    Public Sub initiateCacheFromLeadtimeTable()

        If Cache("Leadtime") Is Nothing Then

            Dim db As String = ConfigurationManager.ConnectionStrings(dbConnectionName).ConnectionString
            Dim connstr As String = db

            Dim conn As SqlConnection = New SqlConnection(connstr)
            conn.Open()

            Dim connParam As SqlConnection = New SqlConnection(ConfigurationManager.ConnectionStrings(dbConnForParam).ConnectionString)
            connParam.Open()

            Dim dtUpdateFrom0 As SqlDataAdapter = New SqlDataAdapter("SELECT * FROM Esch_Na_tbl_Lead_Time", conn)

            Dim dtUpdateFrom1 As SqlDataAdapter = New SqlDataAdapter("SELECT max(ReservedCapPerWeek) AS weeklyMax FROM Esch_Na_tbl_VIP_lead_time Group By VIPgroup", connParam)

            Dim dsAccess As DataSet = New DataSet

            Dim cacheDictionary As New Dictionary(Of String, String)



            dtUpdateFrom0.Fill(dsAccess, "Leadtime")

            dtUpdateFrom1.Fill(dsAccess, "Esch_Na_tbl_VIP_lead_time")

            Dim dailyOutputIncludeAllLines = (From productionLine In dsAccess.Tables("Leadtime").AsEnumerable Select productionLine.Field(Of Integer)("lngDaily_output")).Sum()

            Dim reservedCapForVIPeveryWeek = (From vip In dsAccess.Tables("Esch_Na_tbl_VIP_lead_time").AsEnumerable Select vip.Field(Of Integer)("weeklyMax")).Sum()

            'how many percentage is VIP customer orders loading VS. total output on daily basis
            cacheDictionary.Item("vipVStotal") = (CSng(reservedCapForVIPeveryWeek) / CSng(dailyOutputIncludeAllLines) / 7.0).ToString()


            Dim LeadtimeGroupName As String = String.Empty
            Dim linesByGroup As StringBuilder = New StringBuilder(), dailyOutputByGroup As Long = 0
            Dim leadtimeByGroup As Integer = 90

            'get distinct group name
            Dim leadtimeGroup = From ctgry In dsAccess.Tables("Leadtime").AsEnumerable Select ctgry.Field(Of String)("Lead_time_Group") Distinct

            For Each a In leadtimeGroup
                LeadtimeGroupName = a
                Dim b() As DataRow = dsAccess.Tables("Leadtime").Select("Lead_time_Group ='" & LeadtimeGroupName & "'")
                linesByGroup.Clear()
                linesByGroup.Append("(")
                dailyOutputByGroup = 0
                leadtimeByGroup = 90

                If b.Count > 0 Then
                    leadtimeByGroup = b(0).Item("lnglead_time")
                    cacheDictionary.Add("Grp" & LeadtimeGroupName, leadtimeByGroup.ToString())  'store lead time group's leadtime
                    For Each i As DataRow In b  'all the production lines have been put together
                        linesByGroup.Append(i.Item("Production_line") & ",")
                        dailyOutputByGroup += CLng(i.Item("lngDaily_output"))
                    Next
                    linesByGroup.Remove(linesByGroup.Length - 1, 1)  'eliminate the last comma
                    linesByGroup.Append(")")

                    For Each i As DataRow In b
                        ' the list of all productions lines could be searched by reference to one of these lines
                        cacheDictionary.Add(i.Item("Production_line"), linesByGroup.ToString())
                        'the daily output for the group which includes this production line
                        cacheDictionary.Add(i.Item("Production_line") & "OP", dailyOutputByGroup.ToString())
                        'store group name in order to further search for leadtime in a later stage
                        cacheDictionary.Add(i.Item("Production_line") & "Grp", "Grp" & LeadtimeGroupName)
                        'coefficient for each production line
                        cacheDictionary.Add(i.Item("Production_line") & "CO", i.Item("sng_coefficient").ToString())
                    Next

                End If

            Next

            connParam.Close()
            connParam.Dispose()

            dsAccess.Dispose()
            dtUpdateFrom0.Dispose()
            dtUpdateFrom1.Dispose()

            conn.Close()
            conn.Dispose()



            Cache.Insert("Leadtime", cacheDictionary, Nothing, Cache.NoAbsoluteExpiration, TimeSpan.FromMinutes(120))

        End If

    End Sub



    ''' <summary>
    ''' reasign a start time to those  orders
    ''' old VIP order and rev VIP orders --> new VIP orders --> old Non-VIP and rev Non-VIP --> new Non-VIP
    ''' </summary>
    Public Sub get_Orders_Start_Time(ByRef conn As SqlConnection, ByRef msg As String)

        'Try

        'scope will cover those start time not earlier than now and those new orders and txt_auxiliary_code is Null, all shipped or cancelled are excluded.
        Dim startNow As DateTime = DateTime.Now()
        Dim dtUpdateTo As SqlDataAdapter = New SqlDataAdapter("SELECT txt_order_key,flt_unallocate_qty,int_line_no,dat_start_date,txt_auxiliary_code,int_span,int_status_key,txt_VIP,lng_VIP_lead_time,lng_AdvanceOfRevision,dat_order_added,dat_etd FROM  Esch_Na_tbl_orders  WHERE (int_line_no <> " & valueOf("intDummyLine") & " ) And (int_status_key Not In ('invoiced','cancelled')) And ((dat_start_date Is Null) Or (dat_finish_date > " & dateSeparator & startNow.Date & dateSeparator & "))", conn)
        Dim cmdbAccessCmdBuilder As New SqlCommandBuilder(dtUpdateTo)
        dtUpdateTo.UpdateCommand = cmdbAccessCmdBuilder.GetUpdateCommand()
        dtUpdateTo.InsertCommand = cmdbAccessCmdBuilder.GetInsertCommand()
        Dim dsAccess As DataSet = New DataSet


        dtUpdateTo.Fill(dsAccess, "UpdateTo")
        Dim keys(0) As DataColumn
        keys(0) = dsAccess.Tables("UpdateTo").Columns("txt_order_key")
        dsAccess.Tables("UpdateTo").PrimaryKey = keys

        Dim dsTable As DataTable = dsAccess.Tables("UpdateTo")




        Dim lineNo As String 'line no
        Dim linesPerGroup As String ' all the lines per group
        Dim dailyOutputPerGroup As Long 'total daily output for the specific group
        Dim leadtimePerGroup As Single 'the latest leadtime for the specific group
        Dim startTimeToBeAsigned As DateTime


        'frozen period for all the revised orders, if order's start time falls into this period, do not adjust the start time
        Dim theEarliestAdjustableDate As DateTime = DateTime.Today.AddDays(CInt(valueOf("intFrozenDays")))
        'the earliest possible date when consider production frozen window
        Dim theEarliestDate As DateTime = DateTime.Today.AddDays(CInt(valueOf("intProductionFrozenWindow")))
        'how many extra days to add on production finish time to get ex-plant date
        Dim extraDaysAddedForExplantDate As Integer = CInt(valueOf("intDefautDaysAddForEx"))
        'determine the real order entry date in case some new orders which are split from the old order lines and system regard the oldest date as order entry date
        Dim lastWorkingDate As DateTime = DateTime.Today.AddDays(-1)
        If lastWorkingDate.DayOfWeek = DayOfWeek.Sunday Then lastWorkingDate = lastWorkingDate.AddDays(-2)
        Dim orderEntryRevisionDate As DateTime

        Dim reservedCapPercentageForVIP As Single = CSng(LTvalueOf("vipVStotal"))


        '============= old&revised VIP orders ===============================
        Dim filterVIP As String
        If CBool(valueOf("bnlConsiderCSI")) Then 'if we consider the requirements of CSI
            filterVIP = "(txt_VIP Is Not Null) And  (txt_auxiliary_code = 'REV-')"
        Else
            filterVIP = "(txt_VIP Is Not Null) And  (txt_auxiliary_code <> 'NEW')"
        End If


        Dim VIPoldrevRows() As DataRow = dsTable.Select(filterVIP, "lng_VIP_lead_time ASC , dat_start_date ASC") 'Don't adjust orders due to the changes by customers

        For Each oldRev As DataRow In VIPoldrevRows

            If CDate(oldRev.Item("dat_start_date")).CompareTo(theEarliestAdjustableDate) > 0 Then 'if the order's start time is earlier than the preset date,no adjustment on the start time

                lineNo = oldRev.Item("int_line_no").ToString()
                'all the production lines in the same lead time group
                linesPerGroup = LTvalueOf(lineNo)
                'the total daily output for this group
                dailyOutputPerGroup = LTvalueOf(lineNo & "OP") * (1 - reservedCapPercentageForVIP)

                'the lead time for this group (unit: day)
                leadtimePerGroup = CSng(LTvalueOf(LTvalueOf(lineNo & "Grp")))

                leadtimePerGroup = getLeadtimePerGroup(dsTable, leadtimePerGroup, dailyOutputPerGroup, linesPerGroup, startNow)

                Dim span1 As Integer = oldRev.Item("int_span")
                orderEntryRevisionDate = lastWorkingDate

                'if order is not in production frozen window
                If theEarliestDate.CompareTo(CDate(oldRev.Item("dat_start_date"))) < 0 Then
                    If span1 < 0 Then
                        'if this order is early produced, just delay it to make span equal to 0 
                        oldRev.Item("dat_start_date") = CDate(oldRev.Item("dat_start_date")).AddDays(-span1)

                    Else
                        'do comparison between lastWorkingDate and dat_order_added, chose the later one
                        If orderEntryRevisionDate.CompareTo(CDate(oldRev.Item("dat_order_added"))) < 0 Then orderEntryRevisionDate = CDate(oldRev.Item("dat_order_added"))
                        'choose the earlier one between the date that VIP could be advanced to based on order revision date  and the date  that is the earliest based on leadtime calculation
                        startTimeToBeAsigned = IIf(orderEntryRevisionDate.AddDays(oldRev.Item("lng_AdvanceOfRevision")).CompareTo(startNow.AddDays(leadtimePerGroup)) > 0, startNow.AddDays(leadtimePerGroup), orderEntryRevisionDate.AddDays(oldRev.Item("lng_AdvanceOfRevision")))
                        'choose the later one, make sure no negative span
                        If startTimeToBeAsigned.CompareTo(CDate(oldRev.Item("dat_start_date")).AddDays(-span1)) < 0 Then startTimeToBeAsigned = CDate(oldRev.Item("dat_start_date")).AddDays(-span1)
                        'can not be earlier than the earliest date decided by production frozen window
                        If startTimeToBeAsigned.CompareTo(theEarliestDate) < 0 Then startTimeToBeAsigned = theEarliestDate
                        'if newly proposed start time is later than the current one while span is positive, keep current value unchanged
                        If CDate(oldRev.Item("dat_start_date")).CompareTo(startTimeToBeAsigned) > 0 Then oldRev.Item("dat_start_date") = startTimeToBeAsigned
                    End If
                End If
                'update the leadtime for the group
                setLTvalueAs(LTvalueOf(lineNo & "Grp"), leadtimePerGroup.ToString())

            End If
        Next


        '============= new VIP orders ===============================
        'select all the new orders in the dataTable "UpdateTo" including VIP orders 
        Dim VIPnewRows() As DataRow = dsTable.Select("(txt_VIP Is Not Null) And (txt_auxiliary_code ='NEW')", "lng_VIP_lead_time ASC")

        For Each newOrder As DataRow In VIPnewRows
            lineNo = newOrder.Item("int_line_no").ToString()
            'all the production lines in the same lead time group
            linesPerGroup = LTvalueOf(lineNo)
            'the total daily output for this group
            dailyOutputPerGroup = LTvalueOf(lineNo & "OP") * (1 - reservedCapPercentageForVIP)

            'the lead time for this group (unit: day)
            leadtimePerGroup = CSng(LTvalueOf(LTvalueOf(lineNo & "Grp")))

            leadtimePerGroup = getLeadtimePerGroup(dsTable, leadtimePerGroup, dailyOutputPerGroup, linesPerGroup, startNow)

            orderEntryRevisionDate = lastWorkingDate

            'do comparison between lastWorkingDate and dat_order_added, chose the later one
            If orderEntryRevisionDate.CompareTo(CDate(newOrder.Item("dat_order_added"))) < 0 Then orderEntryRevisionDate = CDate(newOrder.Item("dat_order_added"))
            'choose the earlier one between the date that VIP could be advanced to based on order entry date & committed leadtime  and the date  that is the earliest based on required shipped date (start time should be at least 2 day: extraDaysAddedForExplantDate)
            startTimeToBeAsigned = IIf(orderEntryRevisionDate.AddDays(newOrder.Item("lng_VIP_lead_time")).CompareTo(startNow.AddDays(leadtimePerGroup)) > 0, startNow.AddDays(leadtimePerGroup), orderEntryRevisionDate.AddDays(newOrder.Item("lng_VIP_lead_time")))
            'choose the later one, make sure no negative span (assume working minutes are 0: start time = finish time to make it simple)
            If startTimeToBeAsigned.CompareTo(CDate(newOrder.Item("dat_etd")).AddDays(-extraDaysAddedForExplantDate)) < 0 Then startTimeToBeAsigned = CDate(newOrder.Item("dat_etd")).AddDays(-extraDaysAddedForExplantDate)
            'can not be earlier than the earliest date decided by production frozen window
            If startTimeToBeAsigned.CompareTo(theEarliestDate) < 0 Then startTimeToBeAsigned = theEarliestDate



            newOrder.Item("dat_start_date") = startTimeToBeAsigned

            'update the leadtime for the group
            setLTvalueAs(LTvalueOf(lineNo & "Grp"), leadtimePerGroup.ToString())

        Next



        '============= old&revised Non-VIP ===============================
        Dim filterNonVIP As String
        If CBool(valueOf("bnlConsiderCSI")) Then 'if we consider the requirements of CSI
            filterNonVIP = "(txt_VIP Is  Null) And  (txt_auxiliary_code = 'REV-')"
        Else
            filterNonVIP = "(txt_VIP Is  Null) And  (txt_auxiliary_code <> 'NEW')"
        End If
        'select all the old and revised  orders in the dataTable "UpdateTo" 
        'Dim oldrevRows() As DataRow = dsTable.Select("(txt_VIP Is Null) And ((txt_auxiliary_code ='old') OR (txt_auxiliary_code ='REV'))", "dat_start_date ASC")
        Dim oldrevRows() As DataRow = dsTable.Select(filterNonVIP, "dat_start_date ASC")  'Don't adjust orders due to the changes by customers

        For Each oldRev As DataRow In oldrevRows

            If CDate(oldRev.Item("dat_start_date")).CompareTo(theEarliestAdjustableDate) > 0 Then 'if the order's start time is earlier than the preset date,no adjustment on the start time

                lineNo = oldRev.Item("int_line_no").ToString()
                'all the production lines in the same lead time group
                linesPerGroup = LTvalueOf(lineNo)
                'the total daily output for this group
                dailyOutputPerGroup = LTvalueOf(lineNo & "OP") * (1 - reservedCapPercentageForVIP)

                'the lead time for this group (unit: day)
                leadtimePerGroup = CSng(LTvalueOf(LTvalueOf(lineNo & "Grp")))

                leadtimePerGroup = getLeadtimePerGroup(dsTable, leadtimePerGroup, dailyOutputPerGroup, linesPerGroup, startNow)

                Dim span1 As Integer = oldRev.Item("int_span")
                'orderEntryRevisionDate = lastWorkingDate




                'if order is not in production frozen window
                If theEarliestDate.CompareTo(CDate(oldRev.Item("dat_start_date"))) < 0 Then
                    If span1 < 0 Then
                        'if this order is early produced, just delay it to make span equal to 0 
                        oldRev.Item("dat_start_date") = CDate(oldRev.Item("dat_start_date")).AddDays(-span1)

                    Else
                        'decide start time according to calculated leadtime
                        startTimeToBeAsigned = startNow.AddDays(leadtimePerGroup)
                        'choose the later one, make sure no negative span
                        If startTimeToBeAsigned.CompareTo(CDate(oldRev.Item("dat_start_date")).AddDays(-span1)) < 0 Then startTimeToBeAsigned = CDate(oldRev.Item("dat_start_date")).AddDays(-span1)
                        'can not be earlier than the earliest date decided by production frozen window
                        If startTimeToBeAsigned.CompareTo(theEarliestDate) < 0 Then startTimeToBeAsigned = theEarliestDate
                        'if newly proposed start time is later than the current one while span is positive, keep current value unchanged
                        If CDate(oldRev.Item("dat_start_date")).CompareTo(startTimeToBeAsigned) > 0 Then oldRev.Item("dat_start_date") = startTimeToBeAsigned
                    End If
                End If
                'update the leadtime for the group
                setLTvalueAs(LTvalueOf(lineNo & "Grp"), leadtimePerGroup.ToString())

            End If

        Next


        '============= new Non-VIP orders ===============================
        'select all the new orders in the dataTable "UpdateTo" 
        Dim newRows() As DataRow = dsTable.Select("(txt_VIP Is  Null) And (txt_auxiliary_code ='NEW')")

        For Each newOrder As DataRow In newRows
            lineNo = newOrder.Item("int_line_no").ToString()
            'all the production lines in the same lead time group
            linesPerGroup = LTvalueOf(lineNo)
            'the total daily output for this group
            dailyOutputPerGroup = LTvalueOf(lineNo & "OP") * (1 - reservedCapPercentageForVIP)

            'the lead time for this group (unit: day)
            leadtimePerGroup = CSng(LTvalueOf(LTvalueOf(lineNo & "Grp")))

            leadtimePerGroup = getLeadtimePerGroup(dsTable, leadtimePerGroup, dailyOutputPerGroup, linesPerGroup, startNow)

            orderEntryRevisionDate = lastWorkingDate

            'decide start time according to calculated leadtime
            startTimeToBeAsigned = startNow.AddDays(leadtimePerGroup)
            'choose the later one, make sure no negative span (assume working minutes are 0: start time = finish time to make it simple)
            If startTimeToBeAsigned.CompareTo(CDate(newOrder.Item("dat_etd")).AddDays(-extraDaysAddedForExplantDate)) < 0 Then startTimeToBeAsigned = CDate(newOrder.Item("dat_etd")).AddDays(-extraDaysAddedForExplantDate)
            'can not be earlier than the earliest date decided by production frozen window
            If startTimeToBeAsigned.CompareTo(theEarliestDate) < 0 Then startTimeToBeAsigned = theEarliestDate


            newOrder.Item("dat_start_date") = startTimeToBeAsigned

            'update the leadtime for the group
            setLTvalueAs(LTvalueOf(lineNo & "Grp"), leadtimePerGroup.ToString())

        Next




        dtUpdateTo.Update(dsAccess, "UpdateTo")  'resolve changes back to database

        dsAccess.Dispose()
        dtUpdateTo.Dispose()

        'update leadtime table with the latest leadtime 

        'Catch ex As Exception
        'msg &= "<div style = 'color:red;'> get_Orders_Start_Time: " & ex.Message & "</div>"

        'End Try

    End Sub


    ''' <summary>
    ''' give high priority to new sample orders, price = 0,txt_local_so Like 'FN%',txt_local_so Like '%SMP%'
    ''' </summary>
    Public Sub scheduleFreeSampleStartTime(ByRef conn As SqlConnection, ByRef msg As String)

        Try

            Dim dtUpdateTo As SqlDataAdapter = New SqlDataAdapter("SELECT txt_order_key,dat_start_date,dat_etd,txt_remark  FROM  Esch_Na_tbl_orders  WHERE (int_line_no <> " & valueOf("intDummyLine") & " ) And (int_status_key = '" & newOrderStatus & "') And ((txt_local_so Like 'FN%') Or (txt_local_so Like '%SMP%') Or (flt_sales_price = 0))", conn)
            Dim cmdbAccessCmdBuilder As New SqlCommandBuilder(dtUpdateTo)
            dtUpdateTo.UpdateCommand = cmdbAccessCmdBuilder.GetUpdateCommand()
            dtUpdateTo.InsertCommand = cmdbAccessCmdBuilder.GetInsertCommand()
            Dim dsAccess As DataSet = New DataSet

            dtUpdateTo.Fill(dsAccess, "UpdateTo")
            Dim keys(0) As DataColumn
            keys(0) = dsAccess.Tables("UpdateTo").Columns("txt_order_key")
            dsAccess.Tables("UpdateTo").PrimaryKey = keys

            Dim startNow As DateTime = DateTime.Now().AddDays(1) ' the earliest time new sample could be arranged = morrow
            Dim extraDaysAddedForEX As Integer = CInt(valueOf("intDefautDaysAddForExWhenFreeSample"))

            For Each a As DataRow In dsAccess.Tables("UpdateTo").Rows
                If startNow.CompareTo(CDate(a.Item("dat_etd")).AddDays(-extraDaysAddedForEX)) < 0 Then
                    a.Item("dat_start_date") = CDate(a.Item("dat_etd")).AddDays(-extraDaysAddedForEX)
                Else
                    a.Item("dat_start_date") = startNow
                End If

            Next


            dtUpdateTo.Update(dsAccess, "UpdateTo")  'resolve changes back to database

            dsAccess.Dispose()
            dtUpdateTo.Dispose()

        Catch ex As Exception
            msg &= "<div style = 'color:red;'> scheduleFreeSampleStartTime: " & ex.Message & "</div>"

        End Try

    End Sub

    ''' <summary>
    ''' delay start time for those orders without payment (Need exclude free sample orders price = 0,txt_local_so Like 'FN%',txt_local_so Like '%SMP%')
    ''' </summary>
    Public Sub paymentTermsCheck(ByRef conn As SqlConnection, ByRef msg As String)

        Try

            Dim paymentTerms As String = "('" & valueOf("strPaymentTermsList").Replace(" ", "").Replace(",", "','") & "')"

            Dim dtUpdateTo As SqlDataAdapter = New SqlDataAdapter("SELECT txt_order_key,dat_start_date,txt_payment_status  FROM  Esch_Na_tbl_orders  WHERE (int_line_no <> " & valueOf("intDummyLine") & " ) And (int_status_key = '" & newOrderStatus & "') And Not ((txt_local_so Like 'FN%') Or (txt_local_so Like '%SMP%') Or (flt_sales_price = 0)) And txt_payment_term In " & paymentTerms, conn)
            Dim cmdbAccessCmdBuilder As New SqlCommandBuilder(dtUpdateTo)
            dtUpdateTo.UpdateCommand = cmdbAccessCmdBuilder.GetUpdateCommand()
            dtUpdateTo.InsertCommand = cmdbAccessCmdBuilder.GetInsertCommand()
            Dim dsAccess As DataSet = New DataSet

            dtUpdateTo.Fill(dsAccess, "UpdateTo")
            Dim keys(0) As DataColumn
            keys(0) = dsAccess.Tables("UpdateTo").Columns("txt_order_key")
            dsAccess.Tables("UpdateTo").PrimaryKey = keys

            Dim startOK As DateTime = DateTime.Now().AddDays(5) ' the earliest time new orders without open account could be arranged
            Dim extraDaysAddedForEX As Integer = CInt(valueOf("intDefautDaysAddForExWhenFreeSample"))

            Dim paymentStatus As String = "np" & DateTime.Today.ToShortDateString()


            For Each a As DataRow In dsAccess.Tables("UpdateTo").Rows
                If startOK.CompareTo(CDate(a.Item("dat_start_date"))) > 0 Then
                    a.Item("dat_start_date") = startOK
                End If
                a.Item("txt_payment_status") = paymentStatus
            Next


            dtUpdateTo.Update(dsAccess, "UpdateTo")  'resolve changes back to database

            dsAccess.Dispose()
            dtUpdateTo.Dispose()

        Catch ex As Exception
            msg &= "<div style = 'color:red;'> paymentTermsCheck: " & ex.Message & "</div>"

        End Try

    End Sub



    ''' <summary>
    ''' Need resolve the updated lead time back to table Esch_Na_tbl_Lead_Time
    ''' </summary>
    Public Sub resolveLeadtimeBackToTable(ByRef conn As SqlConnection, ByRef msg As String)

        Try

            Dim dtUpdateTo As SqlDataAdapter = New SqlDataAdapter("SELECT *  FROM  Esch_Na_tbl_Lead_Time", conn)
            Dim cmdbAccessCmdBuilder As New SqlCommandBuilder(dtUpdateTo)
            dtUpdateTo.UpdateCommand = cmdbAccessCmdBuilder.GetUpdateCommand()
            dtUpdateTo.InsertCommand = cmdbAccessCmdBuilder.GetInsertCommand()

            Dim dtTable As DataTable = New DataTable()

            dtUpdateTo.Fill(dtTable)
            Dim keys(1) As DataColumn
            keys(0) = dtTable.Columns("Lead_time_Group")
            keys(1) = dtTable.Columns("Production_line")
            dtTable.PrimaryKey = keys

            Dim startPoint As DateTime = DateTime.Today.AddDays(1) 'add one extra day

            For Each a As DataRow In dtTable.Rows
                a.Item("lnglead_time") = CInt(LTvalueOf("Grp" & a.Item("Lead_time_Group")))
                a.Item("datStartTime") = startPoint.AddDays(a.Item("lnglead_time")).Date
            Next


            dtUpdateTo.Update(dtTable)  'resolve changes back to database

            dtTable.Dispose()
            dtUpdateTo.Dispose()

        Catch ex As Exception
            msg &= "<div style = 'color:red;'> paymentTermsCheck: " & ex.Message & "</div>"
        End Try

    End Sub


    ''' <summary>
    ''' get the latest leadtime per group
    ''' </summary>
    ''' <param name="dtTable">the dataTable to be use to calculate lead time based on quantity loading</param>
    ''' <param name="leadtimePerGroup">the leadtime to be updated and used through the whole processing</param>
    ''' <param name="dailyOutputPerGroup">daily output for the same group</param>
    ''' <param name="linesPerGroup">the production lines group which shares the same lead time</param>
    ''' <param name="startNow">record the memoment when to do leadtime calculation and anuto scheduling</param>
    Public Function getLeadtimePerGroup(ByRef dtTable As DataTable, ByVal leadtimePerGroup As Single, ByVal dailyOutputPerGroup As Long, ByVal linesPerGroup As String, ByVal startNow As DateTime) As Single

        Dim allOrders = dtTable.Select("(dat_start_date Is Not Null) And (int_line_no In " & linesPerGroup & ")").AsEnumerable()

        'Dim totalQtyPerGroup = (From qty In allOrders Where (DateDiff(DateInterval.Minute, qty.Field(Of DateTime)("dat_start_date"), startNow.AddDays(leadtimePerGroup)) > 0)
        Dim totalQtyPerGroup = (From qty In allOrders Where (qty.Field(Of DateTime)("dat_start_date").CompareTo(startNow.AddDays(leadtimePerGroup)) < 0)
                                    Select qty.Field(Of Integer)("flt_unallocate_qty")).Sum()

        'Dim gap As Long = Math.Abs(totalQtyPerGroup - dailyOutputPerGroup * leadtimePerGroup)
        Dim tempLeadtimeLow As Single = 0, tempLT As Single = 0
        'give a initial value for the lead time
        If leadtimePerGroup < 32 Then
            leadtimePerGroup *= 4
        Else
            leadtimePerGroup *= 2
        End If

        ' if the gap between totalQtyPerGroup and theoretical output during the calculation period   is less than 3 days, then the lead time is OK
        'when we can insert 3 days' orders along the time horizon
        While Math.Abs(leadtimePerGroup - tempLT) > 3

            If totalQtyPerGroup > dailyOutputPerGroup * (leadtimePerGroup - 1) Then
                tempLT = leadtimePerGroup
                'leadtimePerGroup = leadtimePerGroup - (leadtimePerGroup - tempLeadtimeLow) * 0.5
                leadtimePerGroup = leadtimePerGroup * 1.5 - tempLeadtimeLow * 0.5
                tempLeadtimeLow = tempLT
            Else
                tempLT = leadtimePerGroup
                'leadtimePerGroup = leadtimePerGroup + (leadtimePerGroup - tempLeadtimeLow) * 0.5
                leadtimePerGroup = leadtimePerGroup * 0.5 + tempLeadtimeLow * 0.5
                'tempLeadtimeLow = tempLT
            End If

            totalQtyPerGroup = (From qty In allOrders Where (qty.Field(Of DateTime)("dat_start_date").CompareTo(startNow.AddDays(leadtimePerGroup)) < 0)
                                Select qty.Field(Of Integer)("flt_unallocate_qty")).Sum()


        End While

        Return leadtimePerGroup

    End Function



End Class
