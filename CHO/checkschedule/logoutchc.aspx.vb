
Partial Class checkschedule_logoutchc
    Inherits System.Web.UI.Page

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        If ((Request.Cookies("userInfo") Is Nothing) OrElse (CType(Server.HtmlEncode(Request.Cookies("userInfo")("level")), Integer) < 4)) Then
            'ClientScript.RegisterClientScriptBlock("OnLoad", "<script>alert('Welcome to ShotDev.Com')</script>")
            Response.Clear()
            Response.Redirect("loginchc.aspx")
        End If

        Dim userwelcome1 As Literal

        userwelcome1 = CType(Page.FindControl("Literal1"), Literal)
        If Not (userwelcome1 Is Nothing) Then
            userwelcome1.Text = "Welcome " & Server.HtmlEncode(Request.Cookies("userInfo")("id")) & " !"
        End If

    End Sub

    Protected Sub Button1_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles Button1.Click
        Response.Cookies("userInfo").Expires = DateTime.Now.AddDays(-1)
        Response.Redirect("loginchc.aspx")
    End Sub
End Class
