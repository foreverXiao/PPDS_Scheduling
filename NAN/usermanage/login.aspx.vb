Imports System.Data
Imports System.Data.SqlClient
Imports System.Configuration


Partial Class usermanage_login
    Inherits System.Web.UI.Page


    Protected Sub LoginButton_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles LoginButton.Click

        

        Dim userVerified As Boolean = False
        Dim nextpage As String = "login.aspx"
        If Not String.IsNullOrEmpty(Request.Params("oldurl")) Then
            nextpage = Request.Params("oldurl")
        Else

        End If

        Dim connstr As String = ConfigurationManager.ConnectionStrings("accessDB").ConnectionString

        'Dim connParam As SqlConnection = New SqlConnection(connstr)
        Dim connParam As SqlConnection = New SqlConnection(connstr)

        'Dim dtFrom As SqlDataAdapter = New SqlDataAdapter("Select user_name,rightlevel,password From  Esch_Na_tbl_userrole  ", connParam)
        dim dtFrom as SqlDataAdapter = new SqlDataAdapter("Select user_name,rightlevel,password From  Esch_Na_tbl_userrole  ", connParam)

        Dim dtTable As New DataTable()
        dtFrom.Fill(dtTable)

        Dim accessLvl As Integer = 1
        For Each r As DataRow In dtTable.Rows
            If r.Item("user_name").ToString.ToLower.Equals(UserName.Text.ToLower) AndAlso (True OrElse r.Item("password").ToString.ToLower.Equals(Password.Text.ToLower)) Then
                userVerified = True
                Response.Cookies("userInfo")("id") = r.Item("user_name").ToString
                Response.Cookies("userInfo")("level") = r.Item("rightlevel").ToString
                accessLvl = CInt(r.Item("rightlevel"))
                Dim cookiePath As String = Request.ApplicationPath
                Response.Cookies(cookiePath).Value = cookiePath

                Exit For
            End If
        Next

        dtTable.Dispose()
        dtFrom.Dispose()



        connParam.Dispose()



        If userVerified Then
            If Not String.IsNullOrEmpty(nextpage) Then
                Response.Redirect(nextpage)
            End If

        Else

            st.Text = "<div style='color:red;'>" & "Incorrect login. Incorrect user name, password or the page you do not have authority to visit." & "</div>"
        End If


    End Sub


    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        'Is last time login effective?
        Dim effectiveLogin As Boolean = False

        Dim cookiePath As String = Request.ApplicationPath

        'whether IIS application user login is valid
        If (Request.Cookies("userInfo") IsNot Nothing) AndAlso (Request.Cookies(cookiePath) IsNot Nothing) Then

            effectiveLogin = True
            UserNameRequired.Enabled = False
            UserName.Text = Server.HtmlEncode(Request.Cookies("userInfo")("id"))
            UserName.Enabled = False
            'UserName.BackColor = Drawing.Color.Gray
            PasswordRequired.Enabled = False
            Password.Enabled = False
            LoginButton.Enabled = False
            'Password.BackColor = Drawing.Color.Gray

        Else
            UserName.Enabled = True
            'UserName.BackColor = Drawing.Color.White
            Password.Enabled = True
            'Password.BackColor = Drawing.Color.White
            UserName.Focus()
        End If

        UserName.Text = Request.LogonUserIdentity.Name.Substring(Request.LogonUserIdentity.Name.LastIndexOf("\") + 1)
        Password.Attributes.Add("Value", "Password")

        'effetive login, then go to another web page
        If effectiveLogin Then
            If Not String.IsNullOrEmpty(Request.Params("accLvl")) AndAlso (CInt(Request.Params("accLvl")) <= CInt(Request.Cookies("userInfo")("level"))) Then
                If Not String.IsNullOrEmpty(Request.Params("oldurl")) Then
                    Response.Redirect(Request.Params("oldurl") & "?login=1")
                End If
            Else
                If Not String.IsNullOrEmpty(Request.Params("oldurl")) Then
                    st.Text = "<div style='color:red;'>You do not have authority to visit specific page. " & Request.Params("oldurl") & "</div>"
                End If
            End If
        End If




    End Sub

End Class
