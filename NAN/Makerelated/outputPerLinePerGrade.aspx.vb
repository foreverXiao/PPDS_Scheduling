Imports System.IO
Imports System.Text
Imports System.Data
Imports System.Data.SqlClient
Imports System.Globalization
Imports Microsoft.VisualBasic
Imports System.Linq



Partial Class Makerelated_outputPerLinePerGrade
    Inherits InteracWithExcel

    Public Delegate Sub asychrSub() 'to run an asynchronous function,to generate a new table to show which grade can be produced in how many exact production lines


    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        SDS1.ConnectionString = ConfigurationManager.ConnectionStrings(dbConnForParam).ConnectionString


        maxRowNumber = 15001

        pageLoadInitiate(SDS1, DDL1, DDL2, filtercdtn1, Filter1, Download1, hiddenBT)

    End Sub


    Protected Sub Download1_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles Download1.Click

        downloadExcelFileFromSqlDataSource(SDS1, StatusLabel, hiddenBT)


    End Sub


    Protected Sub Filter1_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles Filter1.Click
        filterClickHnadler(filtercdtn1, DDL1, DDL2, SDS1, LV1)


    End Sub



    Protected Sub DDL1_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles DDL1.TextChanged
        DDLchangeSelection(DDL1, DDL2)
    End Sub

    Protected Sub clrfltr1_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles clrfltr1.Click
        clrfltrClickHnadler(SDS1, LV1)
    End Sub

    Protected Sub UpldDel_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles UpldDel.Click

        UpldDel_ClickHnadler(SDS1, LV1, StatusLabel, FileUpload1)
        asyAction()

    End Sub


    Protected Sub UpldInsrt_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles UpldInsrt.Click

        UpldInsrt_ClickHnadler(SDS1, LV1, StatusLabel, FileUpload1)
        asyAction()

    End Sub


    '''update table per the data in excel file
    Protected Sub UpldUpdate_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles UpldUpdate.Click

        UpldUpdate_ClickHnadler(SDS1, LV1, StatusLabel, FileUpload1)
        asyAction()
    End Sub

    Protected Sub overwrite_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles overwrite.Click

        overwrite_ClickHnadler(SDS1, LV1, StatusLabel, FileUpload1)
        asyAction()
    End Sub


    Public Overrides Function dataValidityCheck(ByRef dtTbl As DataTable) As String


        Dim msstring As String = String.Empty

        Try
            Dim i As Integer = 0
            Dim j As Integer = dtTbl.Columns.Count - 1

            For Each rowexcel As DataRow In dtTbl.Rows

                i = 0
                While i < j
                    If DBNull.Value.Equals(rowexcel.Item(i)) Then
                        If Not String.Equals(dtTbl.Columns(i).ColumnName, "txtRemark", StringComparison.OrdinalIgnoreCase) Then
                            msstring = "Empty value in column '" & dtTbl.Columns(i).ColumnName & "' is not allowed."
                            Exit For
                        End If
                    End If
                    i += 1
                End While

                If Not IsNumeric(rowexcel.Item("int_rate")) Then
                    msstring = "There is something wrong with row txt_grade,txt_line_no :" & rowexcel("txt_grade") & "," & rowexcel("txt_line_no") & "! Maybe int_rate is not a valid number."
                    Exit For
                End If

            Next
        Catch
            msstring = "There is something wrong with the data validity."
        End Try

        Return msstring

    End Function

    Protected Sub LV1_ItemInserting(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.ListViewInsertEventArgs) Handles LV1.ItemInserting
        ' Cancel the update operation if any of the fields is empty
        ' or null.
        Dim messageText As StringBuilder = New StringBuilder()

        For Each de As DictionaryEntry In e.Values
            ' Check if the value is null or empty except field txtRemark
            If (de.Value Is Nothing OrElse de.Value.ToString().Trim().Length = 0) And Not (de.Key.ToString = "txtRemark") Then
                messageText.Append("Cannot insert an empty value.<br />")
            End If
        Next


        If messageText.Length > 0 Then
            e.Cancel = True
            Message.Text = messageText.ToString()
        End If
    End Sub


    ''' <summary>
    ''' You can do some checking on data validity; if some exception is thrown out, the updating would be cancelled
    ''' </summary>
    ''' 
    Protected Sub LV1_ItemUpdating(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.ListViewUpdateEventArgs) Handles LV1.ItemUpdating

        Dim messageText As StringBuilder = New StringBuilder()

        If String.IsNullOrEmpty(e.NewValues("int_rate")) Then
            messageText.Append("Empty value in field int_rate is not allowed.")
        Else
            If Not IsNumeric(e.NewValues("int_rate")) Then messageText.Append("Number is needed in field int_rate.")
        End If

        If messageText.Length > 0 Then
            e.Cancel = True
            Message.Text = messageText.ToString()
        End If


    End Sub

    ''' <summary>
    ''' reveal insertion template
    ''' </summary>
    Protected Sub LV1_ItemCommand(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.ListViewCommandEventArgs) Handles LV1.ItemCommand
        If e.CommandName.Equals("New", StringComparison.OrdinalIgnoreCase) Then
            Dim me1 As ListView = CType(sender, ListView)
            me1.InsertItemPosition = InsertItemPosition.FirstItem
        End If
    End Sub


    Protected Sub LV1_ItemCanceling(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.ListViewCancelEventArgs) Handles LV1.ItemCanceling

        If e.CancelMode = ListViewCancelMode.CancelingInsert Then
            Dim me1 As ListView = CType(sender, ListView)
            me1.InsertItemPosition = InsertItemPosition.None
        End If

    End Sub


    Protected Sub LV1_ItemInserted(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.ListViewInsertedEventArgs) Handles LV1.ItemInserted
        Dim me1 As ListView = CType(sender, ListView)
        me1.InsertItemPosition = InsertItemPosition.None
    End Sub

    'trigger a asynchronous action to do some routine works
    Public Sub asyAction()

        Dim asySubroutine As asychrSub
        asySubroutine = New asychrSub(AddressOf GradeAndItsQualifiedProductionLines)
        asySubroutine.BeginInvoke(Nothing, Nothing)
    End Sub

    'generate a new table Esch_Na_tbl_assign_line_by_grade  which shows which production lines can produce specific grade
    Public Sub GradeAndItsQualifiedProductionLines()

        Dim connParam As SqlConnection = New SqlConnection(ConfigurationManager.ConnectionStrings(dbConnForParam).ConnectionString)
        connParam.Open()


        Dim dtUpdateFrom1 As SqlDataAdapter = New SqlDataAdapter("SELECT * FROM Esch_Na_tbl_output_by_line_by_grade", connParam)
        Dim dtUpdateTo As SqlDataAdapter = New SqlDataAdapter("SELECT * FROM Esch_Na_tbl_assign_line_by_grade", connParam)
        Dim cmdbAccessCmdBuilder As New SqlCommandBuilder(dtUpdateTo)
        dtUpdateTo.DeleteCommand = cmdbAccessCmdBuilder.GetDeleteCommand()
        dtUpdateTo.UpdateCommand = cmdbAccessCmdBuilder.GetUpdateCommand()
        dtUpdateTo.InsertCommand = cmdbAccessCmdBuilder.GetInsertCommand()
        'Dim dsAccess As DataSet = New DataSet
        Dim updateToTB As DataTable = New DataTable
        Dim updateFromTB As DataTable = New DataTable

        dtUpdateTo.Fill(updateToTB)
        Dim keys(0) As DataColumn
        keys(0) = updateToTB.Columns("txt_grade")
        updateToTB.PrimaryKey = keys

        dtUpdateFrom1.Fill(updateFromTB)

        For Each ROW1 As DataRow In updateToTB.Rows
            ROW1.Delete()
        Next

        Dim gradeNames = From c In updateFromTB.AsEnumerable() Select c.Field(Of String)("txt_grade") Distinct
        Dim sqlSelect As StringBuilder = New StringBuilder()
        For Each grade In gradeNames
            sqlSelect.Clear()
            For Each row2 In updateFromTB.Select("txt_grade = '" & grade & "'")
                sqlSelect.Append("," & row2.Item("txt_line_no"))
            Next

            If sqlSelect.Length > 0 Then
                sqlSelect.Remove(0, 1)

                Dim to2() As DataRow = updateToTB.Select("txt_grade = '" & grade & "'")
                If to2.Count > 0 Then
                    to2(0).Item("txt_line_no_group") = sqlSelect.ToString
                Else
                    Dim newRow As DataRow = updateToTB.NewRow()
                    newRow.Item("txt_grade") = grade
                    newRow.Item("txt_line_no_group") = sqlSelect.ToString
                    updateToTB.Rows.Add(newRow)
                End If
            End If

        Next


        Try
            dtUpdateTo.Update(updateToTB)  'resolve changes back to database
        Catch ex As Exception

        End Try


        cmdbAccessCmdBuilder.Dispose()

        updateToTB.Dispose()
        updateFromTB.Dispose()

        dtUpdateTo.Dispose()
        dtUpdateFrom1.Dispose()

        connParam.Close()
        connParam.Dispose()

    End Sub

    Protected Sub SDS1_Deleted(sender As Object, e As SqlDataSourceStatusEventArgs) Handles SDS1.Deleted
        asyAction()
    End Sub

    Protected Sub SDS1_Inserted(sender As Object, e As SqlDataSourceStatusEventArgs) Handles SDS1.Inserted
        asyAction()
    End Sub


    
    Protected Sub SDS1_Updated(sender As Object, e As SqlDataSourceStatusEventArgs) Handles SDS1.Updated
        asyAction()
    End Sub
End Class

