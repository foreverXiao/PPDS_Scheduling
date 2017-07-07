Imports basepage1
Imports System.Data


Partial Class usermanage_deleteuser
    Inherits basepage1

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        'Dim sqldtsrcAccess123 As SqlDataSource = CType(Master.FindControl("CP1").FindControl("SDS1"), SqlDataSource)
        SDS1.ConnectionString = ConfigurationManager.ConnectionStrings(dbConnForParam).ConnectionString



    End Sub

 
    Protected Sub GV1_RowDataBound(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.GridViewRowEventArgs) Handles GV1.RowDataBound
        If e.Row.RowType = DataControlRowType.DataRow Then
            'e.Row.Cells(1).Text = "<i>" & e.Row.Cells(1).Text & "</i>"

            'If CInt(e.Row.Cells(2).Text) = 4 Then e.Row.Cells(2).Text = 1
            'If e.Row.Cells(2).Text = "8" Then e.Row.Cells(2).Text = "planner only"
        End If
    End Sub

    Protected Function formatData(ByVal a As String) As String
        formatData = a
        If a = "2" Then formatData = "anyone"
        If a = "4" Then formatData = "schedule checker"
        If a = "6" Then formatData = "PSI or OF"
        If a = "8" Then formatData = "Planner"
    End Function

    Protected Sub GV1_RowDeleting(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.GridViewDeleteEventArgs) Handles GV1.RowDeleting

    End Sub
End Class
