Imports Microsoft.VisualBasic
Imports System.Data.SqlClient
Imports System.Data
Imports System.IO

'another part of class basepage1
Partial Public Class FrequentPlanActions
    Inherits InteracWithExcel
    'further pass newly imported order data to main order table, insert new order and update old order and cancel orders. 
    ' Esch_Na_Esch_Na_tbl_orders_new_revision_from_OPM ===> Esch_Na_tbl_orders
    Public Shared ReadOnly workingDaysOfWeek() As Integer = {0, 1, 1, 1, 1, 1, 0}

    Public Function Planned_production_qty(ByRef conn As SqlConnection) As String

        Dim msgRtrn As New StringBuilder

        'planned production quantity will not change if the start time is within 16 hours
        Dim dtUpdateTo As SqlDataAdapter = New SqlDataAdapter("SELECT txt_order_key,flt_unallocate_qty,planned_production_qty,int_line_no FROM Esch_Na_tbl_orders WHERE ((dat_start_date > " & dateSeparator & DateTime.Now().AddHours(2) & dateSeparator & ") AND ( planned_production_qty <> flt_unallocate_qty) And (int_status_key Not In ('invoiced','cancelled'))) Or (planned_production_qty Is Null)", conn)
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
            msgRtrn.AppendLine("<div style='color:red;'>Planned_production_qty:" & ex.Message & "</div>")
        End Try

        For Each a As DataRow In dsAccess.Tables("UpdateTo").Rows

            If a.Item("flt_unallocate_qty") <= 0 Then
                a.Item("flt_unallocate_qty") = 0 'never allow negative number apearing coz this does not make sense
                a.Item("int_line_no") = valueOf("intDummyLine") 'production line is changed to dummy line
            End If

            a.Item("planned_production_qty") = a.Item("flt_unallocate_qty")

        Next

        Try

            dtUpdateTo.Update(dsAccess, "UpdateTo")
        Catch ex As Exception
            msgRtrn.AppendLine("<div style='color:red;'>Planned_production_qty:" & ex.Message & "</div>")
        End Try

        dtUpdateTo.Dispose()
        dsAccess.Dispose()

        Return msgRtrn.ToString

    End Function


    ''' <summary>
    ''' if price field is blank, assign 4 to it
    ''' </summary>
    ''' <param name="conn"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function assignValueToBlankPrice(ByRef conn As SqlConnection) As String

        Dim msgRtrn As New StringBuilder

        'planned production quantity will not change if the start time is within 16 hours
        Dim dtUpdateTo As SqlDataAdapter = New SqlDataAdapter("SELECT txt_order_key,flt_sales_price FROM Esch_Na_tbl_orders WHERE (flt_sales_price Is Null)", conn)
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
            msgRtrn.AppendLine("<div style='color:red;'>flt_sales_price:" & ex.Message & "</div>")
        End Try

        For Each a As DataRow In dsAccess.Tables("UpdateTo").Rows
            a.Item("flt_sales_price") = 4  'make the number exaggerated
        Next

        Try

            dtUpdateTo.Update(dsAccess, "UpdateTo")
        Catch ex As Exception
            msgRtrn.AppendLine("<div style='color:red;'>flt_sales_price:" & ex.Message & "</div>")
        End Try

        dtUpdateTo.Dispose()
        dsAccess.Dispose()

        Return msgRtrn.ToString

    End Function


    ''' <summary>
    ''' ABC
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Overloads Function DoCalculateRSD() As String
        Dim connstr As String = ConfigurationManager.ConnectionStrings(dbConnectionName).ConnectionString
        Dim conn As SqlConnection = New SqlConnection(connstr)
        Dim msg As String = calculateRSD(conn, " (dat_start_date > " & dateSeparator & Today & dateSeparator & ") ")
        conn.Dispose()

        Return msg
    End Function

    ''' <summary>
    ''' ABC
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Overloads Function DoCalculateRSD(ByVal extrCondition As String) As String
        Dim connstr As String = ConfigurationManager.ConnectionStrings(dbConnectionName).ConnectionString
        Dim conn As SqlConnection = New SqlConnection(connstr)
        Dim msg As String = calculateRSD(conn, " (dat_start_date > " & dateSeparator & Today & dateSeparator & ") " & extrCondition)
        conn.Dispose()

        Return msg
    End Function


    '''<summary>
    ''' calculate RSD for each revised  or new orders (also consider special order type like apple order certification time)
    ''' </summary>
    Public Function calculateRSD(ByRef conn As SqlConnection, Optional ByVal filterCondition As String = " (int_status_key Not In ('invoiced','cancelled','old')) ") As String

        Dim msgRtrn As New StringBuilder

        Try

            Dim dtUpdateTo As SqlDataAdapter = New SqlDataAdapter("SELECT txt_order_key,txt_item_no,dat_rdd,dat_etd,txt_currency,txt_destination,txt_ship_method,txt_end_user FROM Esch_Na_tbl_orders WHERE " & filterCondition, conn)
            Dim cmdbAccessCmdBuilder As New SqlCommandBuilder(dtUpdateTo)
            dtUpdateTo.UpdateCommand = cmdbAccessCmdBuilder.GetUpdateCommand()

            Dim dtTable As DataTable = New DataTable



            dtUpdateTo.Fill(dtTable)
            Dim keys(0) As DataColumn
            keys(0) = dtTable.Columns("txt_order_key")
            dtTable.PrimaryKey = keys


            msgRtrn.AppendLine(calculateRSD1(conn, dtTable.Select(Nothing)))



            dtUpdateTo.Update(dtTable)


            cmdbAccessCmdBuilder.Dispose()
            dtUpdateTo.Dispose()

        Catch ex As Exception
            msgRtrn.AppendLine("<div style='color:red;'>calculateRSD:" & ex.Message & "</div>")
        End Try

        Return msgRtrn.ToString

    End Function


    '''<summary>
    ''' calculate RSD for each revised  or new orders (also consider special order type like apple order certification time)
    ''' </summary>
    Public Function calculateRSD1(ByRef conn As SqlConnection, ByRef dtRows() As DataRow) As String

        Dim msgRtrn As New StringBuilder

        Dim connParam As SqlConnection = New SqlConnection(ConfigurationManager.ConnectionStrings(dbConnForParam).ConnectionString)
        connParam.Open()

        Try

            Dim dtUpdateFrom As SqlDataAdapter = New SqlDataAdapter("SELECT txt_currency,txt_destination,txt_ship_method,flt_transit FROM Esch_Na_tbl_transit", connParam)

            Dim dsAccess As DataSet = New DataSet

            dtUpdateFrom.Fill(dsAccess, "UpdateFrom")

            Dim frmRows() As DataRow, rsd As Date
            Dim consideringWeekend As Boolean = CBool(valueOf("bnlConsiderWeekendForRSD"))
            Dim defaultTransitDays As Integer = -CInt(valueOf("intDefaulTransit")) 'default transit days if no one matched in transit table
            'Dim extraDaysAddto As Integer = 0
            Dim frm1RowsCount As Integer = 0

            For Each a As DataRow In dtRows
                frmRows = dsAccess.Tables("UpdateFrom").Select("txt_currency = '" & a.Item("txt_currency") & "' And txt_destination = '" & a.Item("txt_destination") & "' And txt_ship_method = '" & a.Item("txt_ship_method") & "'")

                If frmRows.Count > 0 Then
                    rsd = CDate(a.Item("dat_rdd")).AddDays(-(frmRows(0).Item("flt_transit")))
                Else
                    rsd = CDate(a.Item("dat_rdd")).AddDays(defaultTransitDays) ' if no transit time set for this type of transit mode, give a default time
                End If

                'considering the case when calculated RSD fall into weekend period 
                If consideringWeekend Then
                    Select Case rsd.DayOfWeek
                        Case DayOfWeek.Sunday
                            rsd = rsd.AddDays(-2)
                        Case DayOfWeek.Saturday
                            rsd = rsd.AddDays(-1)
                    End Select
                End If

                a.Item("dat_etd") = rsd

            Next


            dtUpdateFrom.Dispose()
            dsAccess.Dispose()

        Catch ex As Exception
            msgRtrn.AppendLine("<div style='color:red;'>" & ex.Message & "</div>")
        End Try

        connParam.Close()
        connParam.Dispose()

        If msgRtrn.Length > 0 Then
            msgRtrn.Insert(0, "Calculate RSD1:")
        End If
        Return msgRtrn.ToString()

    End Function




    ''' <summary>
    ''' calculate working hours, finish time,explant date and span for each order which is not cancelled or shipped and its start time no earlier than 7 days before now
    ''' </summary>
    Public Function finishTime_exPlantDate_Span(ByRef conn As SqlConnection, ByRef hasException As Boolean, Optional ByVal whereClauseCondition As String = " ") As String

        Dim msgRtrn As New StringBuilder()



        Try
            'only do calculation for those order not cancelled or shipped and production start time not earlier than 7 days before now
            Dim dtUpdateTo As SqlDataAdapter = New SqlDataAdapter("SELECT * FROM Esch_Na_tbl_orders WHERE (int_status_key Not In ('invoiced','cancelled')) And (dat_start_date > " & dateSeparator & DateTime.Today.AddDays(-7) & dateSeparator & ") " & whereClauseCondition, conn)
            Dim cmdbAccessCmdBuilder As New SqlCommandBuilder(dtUpdateTo)
            dtUpdateTo.UpdateCommand = cmdbAccessCmdBuilder.GetUpdateCommand()

            Dim dtTable As DataTable = New DataTable

            dtUpdateTo.Fill(dtTable)
            Dim keys(0) As DataColumn
            keys(0) = dtTable.Columns("txt_order_key")
            dtTable.PrimaryKey = keys

            Dim packing_QA_time As Integer = CInt(valueOf("intDefautDaysAddForEx")) 'need time to complete packaging 

            For Each a As DataRow In dtTable.Select(" int_line_no = " & CInt(valueOf("intDummyLine")) & " Or dat_new_explant Is Null ")
                a.Item("dat_finish_date") = a.Item("dat_start_date")
                a.Item("dat_new_explant") = CDate(a.Item("dat_finish_date")).AddDays(packing_QA_time).Date
                a.Item("flt_working_hours") = 0
                a.Item("int_change_over_time") = CInt(valueOf("intDefaultCOT"))
                a.Item("int_span") = 0
            Next

            msgRtrn.AppendLine(finishTime_exPlantDate_Span1(conn, dtTable.Select("int_line_no <> " & CInt(valueOf("intDummyLine"))), hasException))

            'considering additional days to add on original ex-plant date
            msgRtrn.AppendLine(additionDaysOnExplantDate(conn, dtTable, DataViewRowState.ModifiedCurrent))

            dtUpdateTo.Update(dtTable)  'resolve changes back to database

            dtTable.Dispose()
            cmdbAccessCmdBuilder.Dispose()
            dtUpdateTo.Dispose()

        Catch ex As Exception
            msgRtrn.AppendLine("<div style='color:red;'>finishTime_exPlantDate_Span:" & ex.Message & "</div>")
        End Try

        Return msgRtrn.ToString

    End Function


    ''' <summary>
    ''' calculate working hours, finish time,explant date and span for each order which is not cancelled or shipped 
    ''' </summary>
    ''' <param name="limitedLine">used to load less data by narrowing down to specific production line</param>
    Public Function finishTime_exPlantDate_Span1(ByRef conn As SqlConnection, ByRef dtRows() As DataRow, ByRef hasException As Boolean, Optional ByVal limitedLine As String = " ", Optional ByVal whereClauseCondition As String = " ") As String

        Dim msgRtrn As New StringBuilder()

        Dim connParam As SqlConnection = New SqlConnection(ConfigurationManager.ConnectionStrings(dbConnForParam).ConnectionString)
        connParam.Open()

        Dim dtUpdateFrom0 As SqlDataAdapter = New SqlDataAdapter("SELECT * FROM Esch_Na_tbl_output_by_line_by_grade" & limitedLine, connParam)
        Dim dtUpdateFrom1 As SqlDataAdapter = New SqlDataAdapter("SELECT * FROM Esch_Na_tbl_output_by_line_by_item" & limitedLine, connParam)
        Dim dtUpdateFrom3 As SqlDataAdapter = New SqlDataAdapter("SELECT *  FROM Esch_Na_tbl_output_by_line_only", connParam)

        Dim dtUpdateFrom4 As SqlDataAdapter = New SqlDataAdapter("SELECT *  FROM Esch_Na_tbl_Validity", conn)
        Dim cmdbAccessCmdBuilder4 As New SqlCommandBuilder(dtUpdateFrom4)
        dtUpdateFrom4.DeleteCommand = cmdbAccessCmdBuilder4.GetDeleteCommand()
        dtUpdateFrom4.InsertCommand = cmdbAccessCmdBuilder4.GetInsertCommand()

        Dim dsAccess As DataSet = New DataSet
        Dim continues As Boolean = True

        dtUpdateFrom0.Fill(dsAccess, "by_line_and_grade")
        dtUpdateFrom1.Fill(dsAccess, "by_line_and_item")
        dtUpdateFrom3.Fill(dsAccess, "by_line_only")

        dtUpdateFrom4.Fill(dsAccess, "validity")

        Dim defaultRate As Integer = CInt(valueOf("intDefautRate"))  'a default run rate if no one be found in the two tables
        Dim defaultChangeOvertime As Integer = CInt(valueOf("intDefaultCOT"))  'a default change over time
        Dim packing_QA_time As Integer = CInt(valueOf("intDefautDaysAddForEx")) 'need time to complete packaging and 
        Dim extraTimeForSpecialEnduserItem As Integer = 0 'such as apple item certification time
        Dim extraTimeForSpecialEnduserItemCount As Integer = 0 '
        Dim txt_item_no As String = String.Empty 'use this variable to reduce the times of calculating a.item("txt_item_no") in order to get faster speed


        'delete records in table Esch_Na_tbl_Validity
        For Each a As DataRow In dsAccess.Tables("validity").Rows
            a.Delete()
        Next


        For Each a As DataRow In dtRows

            continues = True 'use this variable as a switch to decide whether continue to search for suitable capacity condition and  recorded special issue or remark
            txt_item_no = a.Item("txt_item_no")

            Dim byItem() As DataRow = dsAccess.Tables("by_line_and_item").Select("txt_item_no = '" & txt_item_no & "' And txt_line_no ='" & a.Item("int_line_no") & "'")



            If byItem.Count > 0 Then
                If byItem(0).Item("int_rate") > 0 Then
                    a.Item("flt_working_hours") = a.Item("planned_production_qty") / byItem(0).Item("int_rate") * 60
                    continues = False
                Else  'record some special issue in table Esch_Na_tbl_Validity due to no suitable capacity setting on this item
                    Dim vldtyNew As DataRow = dsAccess.Tables("validity").NewRow
                    vldtyNew.Item("txtCategory1") = a.Item("txt_item_no")
                    vldtyNew.Item("txtCategory2") = " Line:" & a.Item("int_line_no") & " txt_order_key:" & a.Item("txt_order_key")
                    vldtyNew.Item("txtRemark") = byItem(0).Item("txtRemark")
                    dsAccess.Tables("validity").Rows.Add(vldtyNew)
                End If
            End If

            If continues Then
                Dim byGrade() As DataRow = dsAccess.Tables("by_line_and_grade").Select("txt_grade = '" & a.Item("txt_grade") & "' And txt_line_no ='" & a.Item("int_line_no") & "'")
                If byGrade.Count > 0 Then
                    If byGrade(0).Item("int_rate") > 0 Then
                        a.Item("flt_working_hours") = a.Item("planned_production_qty") / byGrade(0).Item("int_rate") * 60
                        continues = False
                    Else  'record some special issue in table Esch_Na_tbl_Validity due to no suitable capacity setting on this grade

                        Dim vldtyNew As DataRow = dsAccess.Tables("validity").NewRow
                        vldtyNew.Item("txtCategory1") = a.Item("txt_grade")
                        vldtyNew.Item("txtCategory2") = " Line:" & a.Item("int_line_no") & " txt_order_key:" & a.Item("txt_order_key")
                        vldtyNew.Item("txtRemark") = byGrade(0).Item("txtRemark")
                        dsAccess.Tables("validity").Rows.Add(vldtyNew)
                    End If
                Else
                    Dim vldtyNew As DataRow = dsAccess.Tables("validity").NewRow
                    vldtyNew.Item("txtCategory1") = a.Item("txt_grade")
                    vldtyNew.Item("txtCategory2") = " Line:" & a.Item("int_line_no") & " txt_order_key:" & a.Item("txt_order_key")
                    vldtyNew.Item("txtRemark") = " No capacity data set up for this grade in this production line."
                    dsAccess.Tables("validity").Rows.Add(vldtyNew)
                End If
            End If

            If continues Then
                Dim byLineonly() As DataRow = dsAccess.Tables("by_line_only").Select("int_line_no  = " & a.Item("int_line_no"))
                If byLineonly.Count > 0 Then
                    If byLineonly(0).Item("int_rate") > 0 Then
                        a.Item("flt_working_hours") = a.Item("planned_production_qty") / byLineonly(0).Item("int_rate") * 60
                        continues = False
                    Else  'record some special issue in table Esch_Na_tbl_Validity due to no default capacity setting on this production line
                        Dim vldtyNew As DataRow = dsAccess.Tables("validity").NewRow
                        vldtyNew.Item("txtCategory1") = "No setting for the line"
                        vldtyNew.Item("txtCategory2") = " Line:" & a.Item("int_line_no") & " txt_order_key:" & a.Item("txt_order_key")
                        vldtyNew.Item("txtRemark") = byLineonly(0).Item("txtRemark")
                        dsAccess.Tables("validity").Rows.Add(vldtyNew)
                    End If
                End If
            End If


            If continues Then
                a.Item("flt_working_hours") = a.Item("planned_production_qty") / defaultRate * 60
            End If

            a.Item("int_change_over_time") = defaultChangeOvertime

            'a.Item("dat_finish_date") = CDate(a.Item("dat_start_date")).AddMinutes(a.Item("flt_working_hours") + a.Item("int_change_over_time"))
            'After FANAR+ GO-LIVE, need change it to calculate start time according to finished time, backward calculation
            a.Item("dat_start_date") = CDate(a.Item("dat_finish_date")).AddMinutes(-a.Item("flt_working_hours") - a.Item("int_change_over_time"))

            a.Item("dat_new_explant") = CDate(a.Item("dat_finish_date")).AddDays(packing_QA_time).Date
            a.Item("int_span") = DateDiff(DateInterval.Day, a.Item("dat_etd"), a.Item("dat_new_explant"))

        Next



        If dsAccess.Tables("validity").Select(Nothing, Nothing, System.Data.DataViewRowState.Added).Count > 0 Then
            hasException = True
            Try
                dtUpdateFrom4.Update(dsAccess, "validity")
            Catch ex As Exception
                msgRtrn.AppendLine("<div style='color:red;'>" & ex.Message & "</div>")
            End Try
        End If


        dsAccess.Dispose()


        dtUpdateFrom0.Dispose()
        dtUpdateFrom1.Dispose()
        dtUpdateFrom3.Dispose()

        cmdbAccessCmdBuilder4.Dispose()
        dtUpdateFrom4.Dispose()

        connParam.Close()
        connParam.Dispose()


        If msgRtrn.Length > 0 Then
            msgRtrn.Insert(0, "Calculate finish time & explant date & span:")
        End If

        Return msgRtrn.ToString()


    End Function


    ''' <summary>
    ''' additional days need to be added to explant date when consider new Shanghai customs system.also consider quality issue related part
    ''' </summary>
    Public Function additionDaysOnExplantDate(ByRef conn As SqlConnection, ByRef updateTo As DataTable, Optional ByVal dataviewRowstateInt As Integer = DataViewRowState.Added) As String


        Dim msgRtrn As New StringBuilder()

        Dim connParam As SqlConnection = New SqlConnection(ConfigurationManager.ConnectionStrings(dbConnForParam).ConnectionString)
        connParam.Open()
        Dim dsAccess As DataSet = New DataSet
        Dim defaultDaysToAdd As Integer = CInt(valueOf("intDefautDaysAddForEx")) 'default extra days to add on finish time to get ex-plant date
        Dim CurrencyCustomsOnTopOfqualityConcern As Boolean = False
        If Not String.IsNullOrEmpty(valueOf("bnlAccumlatedDaysForQualityAndCustoms")) Then
            CurrencyCustomsOnTopOfqualityConcern = CBool(valueOf("bnlAccumlatedDaysForQualityAndCustoms"))
        End If
        Dim WorkingDayCalendarForCustoms As Boolean = False
        If Not String.IsNullOrEmpty(valueOf("bnlWorkingDayCalendarForCustoms")) Then
            WorkingDayCalendarForCustoms = CBool(valueOf("bnlWorkingDayCalendarForCustoms"))
        End If



        'quality concern =================================================
        Dim dtUpdateFrom6 As SqlDataAdapter = New SqlDataAdapter("SELECT *  FROM Esch_Na_tbl_qualityConcern", connParam)


        dtUpdateFrom6.Fill(dsAccess, "qualityConcern")
        Dim listOfCurrencyGroup6 As New List(Of String)
        Dim listOfExtraDyas6 As New List(Of Integer)
        Dim groupNames6 = From c In dsAccess.Tables("qualityConcern").AsEnumerable() Order By c.Item("extraDays") Select c.Field(Of String)("groupName") Distinct
        Dim currencyConditions6 As StringBuilder = New StringBuilder()
        Dim extraChar6 As String = String.Empty
        Dim conditionValues6 As String = String.Empty
        For Each group1 In groupNames6
            currencyConditions6.Clear()
            Dim b() As DataRow = dsAccess.Tables("qualityConcern").Select("groupName = '" & group1 & "'")

            If b.Count > 0 Then

                For Each i As DataRow In b  'summarize up every condition to specific MOQ rule
                    extraChar6 = String.Empty
                    conditionValues6 = i.Item("conditionValue")
                    If i.Item("columnName").ToString().IndexOf("txt") >= 0 Then
                        If i.Item("relationOperator").ToString().ToLower().IndexOf("in") >= 0 Then
                            extraChar6 = ""
                            conditionValues6 = "('" & conditionValues6.Replace(",", "','") & "')"
                        Else
                            extraChar6 = "'"
                        End If

                    End If

                    If i.Item("columnName").ToString().IndexOf("dat") >= 0 Then
                        extraChar6 = dateSeparator
                    End If

                    currencyConditions6.Append(i.Item("columnName") & " " & i.Item("relationOperator") & " " & extraChar6 & conditionValues6 & extraChar6 & " And ")
                Next
                currencyConditions6.Remove(currencyConditions6.Length - 4 - 1, 5)  'eliminate the last operator ' And '
                listOfCurrencyGroup6.Add(currencyConditions6.ToString)
                listOfExtraDyas6.Add(b(0).Item("extraDays"))

            End If

        Next


        Try
            For i As Integer = 0 To listOfCurrencyGroup6.Count - 1
                Dim updateToRows() As DataRow = updateTo.Select(listOfCurrencyGroup6(i), Nothing, dataviewRowstateInt)
                Dim extraDays1 As Integer = listOfExtraDyas6(i)

                If False Then

                End If


                For Each dtRow As DataRow In updateToRows
                    dtRow.Item("dat_new_explant") = CDate(dtRow.Item("dat_new_explant")).AddDays(extraDays1).Date
                    dtRow.Item("int_span") = DateDiff(DateInterval.Day, dtRow.Item("dat_etd"), dtRow.Item("dat_new_explant"))
                Next

            Next

        Catch ex As Exception
            msgRtrn.AppendLine("<div style = 'color:red;'> additionalDaysOnTopOfOriginalExplantDate(quality concern part): " & ex.Message & "</div>")
        End Try


        dtUpdateFrom6.Dispose()
        'quality concern =================================================


        'considering currency and customs =================================
        Dim dtUpdateFrom5 As SqlDataAdapter = New SqlDataAdapter("SELECT *  FROM Esch_Na_tbl_currencyCustoms order by extraDays DESC", connParam)

        dtUpdateFrom5.Fill(dsAccess, "currencyCustoms")
        Dim listOfCurrencyGroup As New List(Of String)
        Dim listOfExtraDyas As New List(Of Integer)
        Dim groupNames = From c In dsAccess.Tables("currencyCustoms").AsEnumerable() Order By c.Item("extraDays") Select c.Field(Of String)("groupName") Distinct
        Dim currencyConditions As StringBuilder = New StringBuilder()
        Dim extraChar As String = String.Empty
        Dim conditionValues As String = String.Empty
        For Each group1 In groupNames
            currencyConditions.Clear()
            Dim b() As DataRow = dsAccess.Tables("currencyCustoms").Select("groupName = '" & group1 & "'")

            If b.Count > 0 Then
                For Each i As DataRow In b  'summarize up every condition to specific customs rule

                    extraChar = String.Empty
                    conditionValues = i.Item("conditionValue")
                    If i.Item("columnName").ToString().IndexOf("txt") >= 0 Then
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

                    currencyConditions.Append(i.Item("columnName") & " " & i.Item("relationOperator") & " " & extraChar & conditionValues & extraChar & " And ")
                Next
                currencyConditions.Remove(currencyConditions.Length - 4 - 1, 5)  'eliminate the last operator ' And '
                listOfCurrencyGroup.Add(currencyConditions.ToString)
                listOfExtraDyas.Add(b(0).Item("extraDays"))

            End If

        Next



        Try
            For i As Integer = 0 To listOfCurrencyGroup.Count - 1
                Dim updateToRows() As DataRow = updateTo.Select(listOfCurrencyGroup(i), Nothing, dataviewRowstateInt)

                Dim extraDays0 As Integer = listOfExtraDyas(i)

                Dim extraCalendarDays As Integer = extraDays0

                For Each dtRow As DataRow In updateToRows

                    If CurrencyCustomsOnTopOfqualityConcern Then
                        If WorkingDayCalendarForCustoms Then
                            extraCalendarDays = workingDaysToCalendarDays(CDate(dtRow.Item("dat_new_explant")).DayOfWeek, extraDays0)
                        Else
                            extraCalendarDays = extraDays0
                        End If
                        dtRow.Item("dat_new_explant") = CDate(dtRow.Item("dat_new_explant")).AddDays(extraCalendarDays).Date
                    Else
                        If WorkingDayCalendarForCustoms Then
                            extraCalendarDays = workingDaysToCalendarDays(CDate(dtRow.Item("dat_finish_date")).AddDays(defaultDaysToAdd).DayOfWeek, extraDays0)
                        Else
                            extraCalendarDays = extraDays0
                        End If
                        If DateDiff(DateInterval.Day, dtRow.Item("dat_finish_date"), dtRow.Item("dat_new_explant")) < (defaultDaysToAdd + extraCalendarDays) Then
                            dtRow.Item("dat_new_explant") = CDate(dtRow.Item("dat_finish_date")).AddDays(defaultDaysToAdd + extraCalendarDays).Date
                        End If
                    End If

                    dtRow.Item("int_span") = DateDiff(DateInterval.Day, dtRow.Item("dat_etd"), dtRow.Item("dat_new_explant"))
                Next

            Next

        Catch ex As Exception
            msgRtrn.AppendLine("<div style = 'color:red;'> additionalDaysOnTopOfOriginalExplantDate(currency&Customs part): " & ex.Message & "</div>")
        End Try

        dtUpdateFrom5.Dispose()
        'considering currency and customs =================================




        dsAccess.Dispose()

        connParam.Close()
        connParam.Dispose()


        Return msgRtrn.ToString()


    End Function

    ''' <summary>
    ''' get the count of calendar days from the count of working days
    ''' </summary>
    ''' <param name="dayOfWeek0"></param>
    ''' <param name="workingDays"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function workingDaysToCalendarDays(ByVal dayOfWeek0 As DayOfWeek, ByVal workingDays As Integer) As Integer

        'Dim multipleOfFive1 As Integer = workingDays \ 5
        'Dim modulusOfFive1 As Integer = workingDays Mod 5

        Dim iCount = 0
        Dim workingDays1 = workingDays
        While workingDays1 > 0
            iCount += 1
            workingDays1 -= workingDaysOfWeek((dayOfWeek0 + iCount) Mod 7)

        End While

        If workingDays < 0 Then  'while workingDays is negative number
            iCount = 0
            While workingDays < 0
                iCount -= 1
                workingDays += workingDaysOfWeek((dayOfWeek0 - 6 * iCount) Mod 7)
            End While
        End If



        Return iCount
    End Function


    ''' <summary>
    ''' Remark on the orders when there is any special package we need pay attention to. Or remark on the orders when there is any particular case
    ''' </summary>
    Public Function PackageOrMiscellaneous(ByRef conn As SqlConnection) As String


        Dim msgRtrn As New StringBuilder

        Dim connParam As SqlConnection = New SqlConnection(ConfigurationManager.ConnectionStrings(dbConnForParam).ConnectionString)
        connParam.Open()

        Dim dtUpdateFrom0 As SqlDataAdapter = New SqlDataAdapter("SELECT groupName,headerName,relationOperator,conditionValue,remarkToBeAdded FROM Esch_Na_tbl_remarkPerCondition", connParam)
        Dim dtUpdateFrom1 As SqlDataAdapter = New SqlDataAdapter("SELECT DISTINCT groupName FROM Esch_Na_tbl_remarkPerCondition", connParam)
        Dim dsAccess As DataSet = New DataSet

        dtUpdateFrom0.Fill(dsAccess, "remarkPerCondition")
        dtUpdateFrom1.Fill(dsAccess, "ColumnGroups")

        Dim columnNames = From c In dsAccess.Tables("remarkPerCondition").AsEnumerable() Select c.Field(Of String)("headerName") Distinct
        Dim sqlSelect As StringBuilder = New StringBuilder()
        For Each i In columnNames
            sqlSelect.Append("," & i.ToString())
        Next


        'Dim dtUpdateTo As SqlDataAdapter = New SqlDataAdapter("SELECT txt_order_key,int_line_no,dat_start_date,txt_remark " & sqlSelect.ToString & " FROM  Esch_Na_tbl_orders  WHERE (int_status_key = '" & newOrderStatus & "') ", conn)
        'FANAR+ GO-LIVE, on Nov.9.2016  need UPDATE these information according to order status, if they are 'NEWnew'
        Dim dtUpdateTo As SqlDataAdapter = New SqlDataAdapter("SELECT txt_order_key,int_line_no,dat_start_date,txt_remark " & sqlSelect.ToString & " FROM  Esch_Na_tbl_orders  WHERE (int_status_key = 'NEWnew') ", conn)
        Dim cmdbAccessCmdBuilder As New SqlCommandBuilder(dtUpdateTo)
        dtUpdateTo.UpdateCommand = cmdbAccessCmdBuilder.GetUpdateCommand()
        dtUpdateTo.InsertCommand = cmdbAccessCmdBuilder.GetInsertCommand()



        dtUpdateTo.Fill(dsAccess, "UpdateTo")
        Dim keys(0) As DataColumn
        keys(0) = dsAccess.Tables("UpdateTo").Columns("txt_order_key")
        dsAccess.Tables("UpdateTo").PrimaryKey = keys





        Dim dummyLateDate As DateTime = CDate(valueOf("datDummyVeryLateStartDate")), dummyLine As Integer = CInt(valueOf("intDummyLine"))
        Dim remarkCondition As StringBuilder = New StringBuilder()
        Dim extraChar As String


        'for each category in MOQ table
        For Each a As DataRow In dsAccess.Tables("ColumnGroups").Rows

            Dim b() As DataRow = dsAccess.Tables("remarkPerCondition").Select("groupName ='" & a.Item("groupName") & "'")
            remarkCondition.Clear()
            If b.Count > 0 Then
                For Each i As DataRow In b  'summarize up every condition to specific MOQ rule
                    extraChar = String.Empty

                    If i.Item("headerName").ToString().IndexOf("txt") >= 0 Then
                        extraChar = "'"
                    End If

                    If i.Item("headerName").ToString().IndexOf("dat") >= 0 Then
                        extraChar = dateSeparator
                    End If

                    remarkCondition.Append(i.Item("headerName") & " " & i.Item("relationOperator") & " " & extraChar & i.Item("conditionValue") & extraChar & " And ")
                Next
                remarkCondition.Remove(remarkCondition.Length - 4 - 1, 5)  'eliminate the last operator ' And '

                Dim c() As DataRow = dsAccess.Tables("UpdateTo").Select(remarkCondition.ToString())
                For Each d As DataRow In c
                    'If d.Item("txt_remark").ToString().IndexOf(b(0).Item("remarkToBeAdded")) = -1 Then 'no MOQ remark has ever been added
                    d.Item("txt_remark") &= " " & b(0).Item("remarkToBeAdded")
                    'End If
                Next
            End If

        Next



        Try
            dtUpdateTo.Update(dsAccess, "UpdateTo")  'resolve changes back to database
        Catch ex As Exception
            msgRtrn.AppendLine("<div style = 'color:red;'> PackageOrMiscellaneous(add remark): " & ex.Message & "</div>")
        End Try

        cmdbAccessCmdBuilder.Dispose()
        dsAccess.Dispose()
        dtUpdateTo.Dispose()
        dtUpdateFrom0.Dispose()
        dtUpdateFrom1.Dispose()

        connParam.Close()
        connParam.Dispose()

        Return msgRtrn.ToString


    End Function


    ''' <summary>
    ''' assign screw die and FDA
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function DoAssignScrewDieAndFDA(Optional ByVal strCondition As String = "") As String
        Dim connstr As String = ConfigurationManager.ConnectionStrings(dbConnectionName).ConnectionString
        Dim conn As SqlConnection = New SqlConnection(connstr)
        'Dim msg As String = AssignScrewDieAndFDA(conn, " (dat_start_date >= " & dateSeparator & Today & dateSeparator & ") " & strCondition)
        'FANAR+ GO-LIVE, on Nov.9.2016  need UPDATE these information according to order status, if they are 'NEWnew'
        Dim msg As String = AssignScrewDieAndFDA(conn, " (int_status_key = 'NEWnew') " & strCondition)
        conn.Dispose()

        Return msg

    End Function



    ''' <summary>
    ''' decide technical parameter like screw and die and FDA for each order line based on its grade or item name
    ''' </summary>
    Public Overloads Function AssignScrewDieAndFDA(ByRef conn As SqlConnection, Optional ByVal filterCondition As String = " (int_status_key = '" & newOrderStatus & "') ") As String

        Dim msgRtrn As New StringBuilder

        Dim startPoint As DateTime = Now.AddDays(1)

        Dim connParam As SqlConnection = New SqlConnection(ConfigurationManager.ConnectionStrings(dbConnForParam).ConnectionString)
        connParam.Open()

        Try



            Dim dtUpdateFrom0 As SqlDataAdapter = New SqlDataAdapter("SELECT * FROM Esch_Na_tbl_screw_die_FDA_by_grade", connParam)
            Dim dtUpdateFrom1 As SqlDataAdapter = New SqlDataAdapter("SELECT * FROM Esch_Na_tbl_screw_die_by_item", connParam)
            Dim dsAccess As DataSet = New DataSet

            dtUpdateFrom0.Fill(dsAccess, "byGrade")
            dtUpdateFrom1.Fill(dsAccess, "byItem")


            'only focus on new orders
            Dim dtUpdateTo As SqlDataAdapter = New SqlDataAdapter("SELECT txt_order_key,txt_grade,txt_item_no,txt_FDA,txt_process_technics  FROM  Esch_Na_tbl_orders  WHERE " & filterCondition, conn)
            Dim cmdbAccessCmdBuilder As New SqlCommandBuilder(dtUpdateTo)
            dtUpdateTo.UpdateCommand = cmdbAccessCmdBuilder.GetUpdateCommand()
            dtUpdateTo.InsertCommand = cmdbAccessCmdBuilder.GetInsertCommand()



            dtUpdateTo.Fill(dsAccess, "UpdateTo")
            Dim Esch_Na_tbl_orders As DataTable = dsAccess.Tables("UpdateTo")
            Dim keys(0) As DataColumn
            keys(0) = dsAccess.Tables("UpdateTo").Columns("txt_order_key")
            dsAccess.Tables("UpdateTo").PrimaryKey = keys


            'for every grade
            For Each a As DataRow In dsAccess.Tables("byGrade").Rows
                Dim b() As DataRow = Esch_Na_tbl_orders.Select("txt_grade = '" & a.Item("txt_grade") & "'")
                For Each c As DataRow In b  'update remark once we found any order matching the criteria
                    c.Item("txt_process_technics") = a.Item("txt_screw").ToString & " " & a.Item("txt_die").ToString
                    c.Item("txt_FDA") = a.Item("txt_FDA").ToString
                Next

            Next


            'for every item, potentially 
            'it will overwrite the value got from table by grade, the rule by item has higher priority to the one by grade
            For Each a As DataRow In dsAccess.Tables("byItem").Rows
                Dim b() As DataRow = Esch_Na_tbl_orders.Select("txt_item_no = '" & a.Item("txt_item_no") & "'")
                For Each c As DataRow In b  'update remark once we found any order matching the criteria
                    c.Item("txt_process_technics") = a.Item("txt_screw").ToString & " " & a.Item("txt_die").ToString
                Next
            Next


            dtUpdateTo.Update(dsAccess, "UpdateTo")  'resolve changes back to database

            dsAccess.Dispose()
            cmdbAccessCmdBuilder.Dispose()
            dtUpdateTo.Dispose()
            dtUpdateFrom0.Dispose()
            dtUpdateFrom1.Dispose()

        Catch ex As Exception
            msgRtrn.AppendLine("<div style='color:red;'> AssignScrewDieAndFDA:" & ex.Message & "</div>")

        End Try

        connParam.Close()
        connParam.Dispose()


        Return msgRtrn.ToString

    End Function


    ''' <summary>
    ''' decide technical parameter like screw and die and FDA for each order line based on its grade or item name
    ''' </summary>
    Public Overloads Function AssignScrewDieAndFDA(ByRef conn As SqlConnection, ByRef dtRows() As DataRow) As String

        Dim msgRtrn As New StringBuilder


        Dim connParam As SqlConnection = New SqlConnection(ConfigurationManager.ConnectionStrings(dbConnForParam).ConnectionString)
        connParam.Open()

        Try



            Dim dtUpdateFrom0 As SqlDataAdapter = New SqlDataAdapter("SELECT * FROM Esch_Na_tbl_screw_die_FDA_by_grade", connParam)
            Dim dtUpdateFrom1 As SqlDataAdapter = New SqlDataAdapter("SELECT * FROM Esch_Na_tbl_screw_die_by_item", connParam)
            Dim dsAccess As DataSet = New DataSet

            dtUpdateFrom0.Fill(dsAccess, "byGrade")
            dtUpdateFrom1.Fill(dsAccess, "byItem")


            'First check if we have a setting on item level, then will search for setting on grade level
            For Each orderline As DataRow In dtRows
                Dim byItem() As DataRow = dsAccess.Tables("byItem").Select("txt_item_no = '" & orderline.Item("txt_item_no") & "'")
                If byItem.Count > 0 Then
                    orderline.Item("txt_process_technics") = byItem(0).Item("txt_screw").ToString() & " " & byItem(0).Item("txt_die").ToString()
                Else
                    Dim byGrade() As DataRow = dsAccess.Tables("byGrade").Select("txt_grade = '" & orderline.Item("txt_grade") & "'")
                    If byGrade.Count > 0 Then
                        orderline.Item("txt_process_technics") = byGrade(0).Item("txt_screw").ToString() & " " & byGrade(0).Item("txt_die").ToString()
                    End If

                End If

            Next




            dsAccess.Dispose()

            dtUpdateFrom0.Dispose()
            dtUpdateFrom1.Dispose()

        Catch ex As Exception
            msgRtrn.AppendLine("<div style='color:red;'> AssignScrewDieAndFDA:" & ex.Message & "</div>")

        End Try

        connParam.Close()
        connParam.Dispose()


        Return msgRtrn.ToString

    End Function




    ''' <summary>
    ''' Update production progress for each order, how many percentage has the order been completed
    ''' flt_actual_completed = flt_actual_qty_from_qa / planned_production_qty
    ''' </summary>
    Public Function UpdateOrderCompletionPercentage(ByRef conn As SqlConnection) As String

        Dim msgRtrn As New StringBuilder

        Try

            'only focus on those orders with RDD not earlier than one week ago
            Dim dtUpdateTo As SqlDataAdapter = New SqlDataAdapter("SELECT txt_order_key,flt_actual_qty_man,flt_actual_completed,planned_production_qty  FROM  Esch_Na_tbl_orders  WHERE dat_start_date > " & dateSeparator & DateTime.Today.AddDays(-7).Date & dateSeparator & "", conn)
            Dim cmdbAccessCmdBuilder As New SqlCommandBuilder(dtUpdateTo)
            dtUpdateTo.UpdateCommand = cmdbAccessCmdBuilder.GetUpdateCommand()


            Dim dtTable As DataTable = New DataTable

            dtUpdateTo.Fill(dtTable)
            Dim keys(0) As DataColumn
            keys(0) = dtTable.Columns("txt_order_key")
            dtTable.PrimaryKey = keys

            If valueOf("strOrganization").ToUpper <> "PGNA" Then 'EXCLUDE Nansha from this practice
                msgRtrn.AppendLine(UpdateOrderCompletionPercentage1(conn, dtTable.Select(Nothing)))
            End If


            dtUpdateTo.Update(dtTable)  'resolve changes back to database


            cmdbAccessCmdBuilder.Dispose()
            dtUpdateTo.Dispose()

        Catch ex As Exception
            msgRtrn.AppendLine("<div style='color:red;'> UpdateOrderCompletionPercentage:" & ex.Message & "</div>")
        End Try

        Return msgRtrn.ToString

    End Function



    '''<summary>
    '''Update production progress for each order, how many percentage has the order been completed
    '''flt_actual_completed = flt_actual_qty_from_qa / planned_production_qty
    '''</summary>
    Function UpdateOrderCompletionPercentage1(ByRef conn As SqlConnection, ByRef dtRows() As DataRow) As String

        Dim msgRtrn As New StringBuilder()

        Dim a As DataRow

        Try

            'this is useless

            For Each a In dtRows
                If IsDBNull(a.Item("flt_actual_qty_man")) Then
                    a.Item("flt_actual_completed") = 0
                Else
                    If a.Item("flt_actual_qty_man") >= a.Item("planned_production_qty") OrElse a.Item("planned_production_qty") = 0 Then
                        a.Item("flt_actual_completed") = 100
                    Else
                        a.Item("flt_actual_completed") = CInt(a.Item("flt_actual_qty_man") * 100 \ a.Item("planned_production_qty"))
                    End If
                End If
            Next

        Catch ex As Exception
            msgRtrn.AppendLine("<div style='color:red;'>" & ex.Message & " txt_order_key : " & a.Item("txt_order_key") & "</div>")
        End Try

        If msgRtrn.Length > 0 Then
            msgRtrn.Insert(0, "UpdateOrderCompletionPercentage1:")
        End If
        Return msgRtrn.ToString()

    End Function





    ''' <summary>
    ''' Delete those orders which's status is cancelled or invoiced
    ''' </summary>
    Public Function deleteUponCondition(ByRef conn As SqlConnection, Optional ByVal condition As String = "int_status_key  ='cancelled'") As Integer

        deleteUponCondition = 0
        'only focus on those orders with RDD not earlier than one week ago
        Dim dtUpdateTo As SqlDataAdapter = New SqlDataAdapter("SELECT txt_order_key  FROM  Esch_Na_tbl_orders  WHERE  " & condition, conn)
        Dim cmdbAccessCmdBuilder As New SqlCommandBuilder(dtUpdateTo)
        dtUpdateTo.DeleteCommand = cmdbAccessCmdBuilder.GetDeleteCommand


        Dim dtTable As DataTable = New DataTable

        dtUpdateTo.Fill(dtTable)
        Dim keys(0) As DataColumn
        keys(0) = dtTable.Columns("txt_order_key")
        dtTable.PrimaryKey = keys


        deleteUponCondition = dtTable.Rows.Count


        For Each a As DataRow In dtTable.Rows
            a.Delete()
        Next



        dtUpdateTo.Update(dtTable)  'resolve changes back to database


        cmdbAccessCmdBuilder.Dispose()
        dtUpdateTo.Dispose()



    End Function





    ''' <summary>
    ''' initiate the list of production lines owned by valid user
    ''' </summary>
    ''' <remarks></remarks>
    Public Sub productionLinesOwnedByValiduser()
        'get all the production lines which are owned by this valid user

        If CacheFrom("Ls") Is Nothing Then 'see if we need refresh all the lines which are owned by valid user

            Dim lines As StringBuilder = New StringBuilder

            Dim connParam As SqlConnection = New SqlConnection(ConfigurationManager.ConnectionStrings(dbConnForParam).ConnectionString)
            connParam.Open()
            Dim command As New SqlCommand("SELECT int_line_no,validuser FROM Esch_Na_tbl_LinesAndOwners  ", connParam)
            Dim reader As SqlDataReader = command.ExecuteReader()


            lines.Append("'0','")
            Do While reader.Read()
                If ("," & reader("validuser") & ",").ToString.Contains(userIden() & ",") Then 'to judge whether the production line is owned by the valid user
                    lines.Append(reader("int_line_no") & "','")
                End If
            Loop
            lines.Append(valueOf("intDummyLine") & "','0'")


            reader.Close()
            command.Dispose()

            connParam.Close()
            connParam.Dispose()


            CacheInsert("Ls", lines.ToString())

        End If

    End Sub


    Public Function lineListOwnedByUser() As String


        productionLinesOwnedByValiduser()

        Return CType(CacheFrom("Ls"), String)

    End Function



    ''' <summary>
    ''' ?
    ''' </summary>
    ''' <remarks></remarks>
    Protected Overrides Sub Finalize()
        MyBase.Finalize()
    End Sub
End Class
