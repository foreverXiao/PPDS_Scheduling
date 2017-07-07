Imports basepage1

Partial Class dragDrop_logout
    Inherits basepage1

    Protected Sub Page_Init1(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Init
        accesslevel = 2 ' set a low security level for this page in order to allow people to log out
    End Sub


    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load


        Dim mpContentPlaceHolder As ContentPlaceHolder
        Dim userwelcome1 As Literal
        mpContentPlaceHolder = CType(Master.FindControl("CP1"), ContentPlaceHolder)
        If Not mpContentPlaceHolder Is Nothing Then
            userwelcome1 = CType(mpContentPlaceHolder.FindControl("Literal1"), Literal)
            If userwelcome1 IsNot Nothing AndAlso Not IsNothing(Request.Cookies("userInfo")) Then
                userwelcome1.Text = "User: " & Server.HtmlEncode(Request.Cookies("userInfo")("id")) & " "
            End If
        End If
    End Sub

    Protected Sub Button1_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles Button1.Click

        clearConsessions()

        Response.Redirect("login.aspx")

    End Sub


    Protected Sub clearConsessions()
        'clear SSO related

        Dim listOfSettingName As String = "sso,systemvariable,Ls,checker,schedule"
        Dim items() As String = listOfSettingName.Split(",".ToCharArray(), StringSplitOptions.RemoveEmptyEntries)

        For Each i As String In items
            If Not (CacheFrom(i) Is Nothing) Then
                CacheRemove(i)
            End If
        Next

        Dim cookiePath As String = Request.ApplicationPath
        Response.Cookies(cookiePath).Expires = DateTime.Now.AddDays(-1)
        Response.Cookies("userInfo").Expires = DateTime.Now.AddDays(-1)


    End Sub

End Class
