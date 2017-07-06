Imports basepage1

Partial Class _Default
    Inherits basepage1

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Response.Redirect("~/dragDrop/OrderDetail.aspx")
        Dim actionRequested As String = Request.Params("title")
        If Not String.IsNullOrEmpty(actionRequested) Then Page.Title = actionRequested
    End Sub
End Class
