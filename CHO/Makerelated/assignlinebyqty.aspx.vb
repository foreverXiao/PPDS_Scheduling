Imports System.IO
Imports System.Text
Imports System.Data
Imports System.Data.SqlClient
Imports System.Globalization
Imports Microsoft.VisualBasic
Imports System.Linq



Partial Class Makerelated_assignlinebyqty
    Inherits InteracWithExcel



    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        SDS1.ConnectionString = ConfigurationManager.ConnectionStrings(dbConnForParam).ConnectionString
 

        maxRowNumber = 301

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

    End Sub


    Protected Sub UpldInsrt_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles UpldInsrt.Click

        UpldInsrt_ClickHnadler(SDS1, LV1, StatusLabel, FileUpload1)


    End Sub


    '''update table per the data in excel file
    Protected Sub UpldUpdate_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles UpldUpdate.Click

        UpldUpdate_ClickHnadler(SDS1, LV1, StatusLabel, FileUpload1)

    End Sub

    Protected Sub overwrite_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles overwrite.Click

        overwrite_ClickHnadler(SDS1, LV1, StatusLabel, FileUpload1)

    End Sub


    Public Overrides Function dataValidityCheck(ByRef dtTbl As DataTable) As String


        Dim msstring As String = String.Empty

        Try


            Dim validRows() As DataRow = dtTbl.Select(Nothing, "minimum,maximum")

            Dim countRows As Integer = validRows.Count - 2

            For i As Integer = 0 To countRows + 1
                If (Not IsNumeric(validRows(i).Item("minimum"))) OrElse (Not IsNumeric(validRows(i).Item("maximum"))) Then
                    msstring = "Illegal number with row minimum = " & validRows(i).Item("minimum") & " and maximum = " & validRows(i).Item("maximum")
                    Exit For
                End If


                If validRows(i).Item("minimum") > validRows(i).Item("maximum") Then
                    msstring = "minimum number is bigger than maximum number with row minimum = " & validRows(i).Item("minimum") & " and maximum = " & validRows(i).Item("maximum")
                    Exit For
                End If

            Next


            If (countRows >= 0) AndAlso String.IsNullOrEmpty(msstring) Then

                For i As Integer = 0 To countRows

                    If (validRows(i).Item("maximum") + 1) < validRows(i + 1).Item("minimum") Then
                        msstring = "Some order quantity would  not be covered by the rules due to the discontinuity between rows  with  the row minimum = " & validRows(i + 1).Item("minimum") & " and the previous row maximum = " & validRows(i).Item("maximum")
                        Exit For
                    End If

                Next

            End If

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
            ' Check if the value is null or empty
            If de.Value Is Nothing OrElse de.Value.ToString().Trim().Length = 0 Then
                messageText.Append("Cannot insert an empty value.<br />")
            End If
        Next

        ' Need numbers. If really happens, exception would be thrown before this updating event
        If IsNumeric(e.Values("minimum")) AndAlso IsNumeric(e.Values("maximum")) Then
            If CInt(e.Values("minimum")) > CInt(e.Values("maximum")) Then
                messageText.Append("minimum number should not exceed maximum one.<br />")
            End If
        Else
            messageText.Append("It need number in field minimum and maximum.<br />")
        End If


        If messageText.Length > 0 Then
            e.Cancel = True
            Message.Text = messageText.ToString()
        End If
    End Sub



    'do some checking on data validity
    Protected Sub LV1_ItemUpdating(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.ListViewUpdateEventArgs) Handles LV1.ItemUpdating
        ' Cancel the update operation if any of the fields is empty
        ' or null.
        Dim messageText As StringBuilder = New StringBuilder()
        For Each de As DictionaryEntry In e.NewValues
            ' Check if the value is null or empty
            If de.Value Is Nothing OrElse de.Value.ToString().Trim().Length = 0 Then
                messageText.Append("Cannot set a field to an empty value.<br />")
            End If
        Next

        ' Need numbers. If really happens, exception would be thrown before this updating event
     


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



    
End Class

