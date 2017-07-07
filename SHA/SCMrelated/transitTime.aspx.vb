Imports System.IO
Imports System.Text
Imports System.Data
Imports System.Data.SqlClient
Imports System.Globalization
Imports Microsoft.VisualBasic
Imports System.Linq



Partial Class SCMrelated_transitTime
    Inherits FrequentPlanActions



    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        SDS1.ConnectionString = ConfigurationManager.ConnectionStrings(dbConnForParam).ConnectionString
        
        maxRowNumber = 5001

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

            For Each rowexcel As DataRow In dtTbl.Rows
                If Not (IsNumeric(rowexcel("flt_transit")) AndAlso IsNumeric(rowexcel("flt_actual"))) Then
                    msstring = "There is something wrong with row txt_currency,txt_destination,txt_ship_method :" & rowexcel("txt_currency") & "," & rowexcel("txt_destination") & "," & rowexcel("txt_ship_method") & "" & "! Maybe flt_transit or flt_actual is not valid number."
                    Exit For
                End If

                If DBNull.Value.Equals(rowexcel("txt_currency")) OrElse rowexcel("txt_currency").ToString().Length > 10 Then
                    msstring = "There is something wrong with row txt_currency,txt_destination,txt_ship_method :" & rowexcel("txt_currency") & "," & rowexcel("txt_destination") & "," & rowexcel("txt_ship_method") & "" & "! For field txt_currency, the string is null or string length is greater than 10 ."
                    Exit For
                End If

                If DBNull.Value.Equals(rowexcel("txt_destination")) OrElse rowexcel("txt_destination").ToString().Length > 10 Then
                    msstring = "There is something wrong with row txt_currency,txt_destination,txt_ship_method :" & rowexcel("txt_currency") & "," & rowexcel("txt_destination") & "," & rowexcel("txt_ship_method") & "" & "! For field txt_destination, the string is null or string length is greater than 10 ."
                    Exit For
                End If

                If DBNull.Value.Equals(rowexcel("txt_ship_method")) OrElse rowexcel("txt_ship_method").ToString().Length > 10 Then
                    msstring = "There is something wrong with row txt_currency,txt_destination,txt_ship_method :" & rowexcel("txt_currency") & "," & rowexcel("txt_destination") & "," & rowexcel("txt_ship_method") & "" & "! For field txt_ship_method, the string is null or string length is greater than 10 ."
                    Exit For
                End If

                If rowexcel("txt_ship_to").ToString().Length > 50 Then
                    msstring = "There is something wrong with row txt_currency,txt_destination,txt_ship_method :" & rowexcel("txt_currency") & "," & rowexcel("txt_destination") & "," & rowexcel("txt_ship_method") & "" & "! For field txt_ship_to, string length is greater than 50 ."
                    Exit For
                End If


            Next

        Catch
            msstring = "There is something wrong with the data validity."
        End Try

        Return msstring


    End Function


    ''' <summary>
    ''' You can do some checking on data validity; if some exception is thrown out, the updating would be cancelled
    ''' </summary>
    ''' 
    Protected Sub LV1_ItemUpdating(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.ListViewUpdateEventArgs) Handles LV1.ItemUpdating

        Dim messageText As StringBuilder = New StringBuilder()

        If String.IsNullOrEmpty(e.NewValues("flt_transit")) Then
            messageText.Append("Empty value in field int_rate is not allowed.")
        Else
            If Not IsNumeric(e.NewValues("flt_transit")) Then messageText.Append("Number is needed in field int_rate.")
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





End Class

