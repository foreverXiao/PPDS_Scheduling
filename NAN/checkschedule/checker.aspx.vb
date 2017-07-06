
Partial Class checkschedule_checker
    Inherits System.Web.UI.Page

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Dim cookiePath As String = Request.ApplicationPath

        If (Request.Cookies("userInfo") Is Nothing) OrElse (CType(Server.HtmlEncode(Request.Cookies("userInfo")("level")), Integer) < 4) OrElse (Request.Cookies(cookiePath) Is Nothing) Then
            '(Request.Cookies("userInfo") Is Nothing) OrElse String.IsNullOrEmpty(Request.Cookies("userInfo")("level")) OrElse (CType(Request.Cookies("userInfo")("level"), Integer) < accesslevel())) 
            Response.Redirect("loginchc.aspx?oldurl=checker.aspx")
        End If
    End Sub



End Class
