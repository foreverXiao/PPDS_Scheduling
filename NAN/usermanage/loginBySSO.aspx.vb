Imports System.Data
Imports System.Data.OleDb
Imports System.Configuration



Partial Class usermanage_loginBySSO
    'Inherits System.Web.UI.Page
    Inherits basepage1

    Protected Sub LoginButton_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles LoginButton.Click

        'if sso and its password seem correct
        If Not String.IsNullOrEmpty(Trim(yoursso.Text)) AndAlso Not String.IsNullOrEmpty(Trim(ssopswrd.Text)) Then
            'whether domain user login is valid

            Dim validSSOuser As New impersonateFTP()
            Dim userOfSSOisValid = validSSOuser.impersonateValidUser(Trim(yoursso.Text), Trim(ssopswrd.Text), valueOf("strCompanyDomain"))
            validSSOuser = Nothing

            If userOfSSOisValid Then
                Dim ssoNV As New NameValueCollection
                ssoNV.Add("sso", Trim(yoursso.Text))
                ssoNV.Add("password", Trim(ssopswrd.Text))
                CacheInsert("sso", ssoNV, 480)

                If Not String.IsNullOrEmpty(Request.Params("oldurl")) Then
                    Response.Redirect(Request.Params("oldurl"))
                End If



            Else
                st.Text = "<div style='color:red;'>" & "Login is failed. Maybe SSO or password is incorrect." & "</div>"
            End If

        Else
            st.Text = "<div style='color:red;'>" & "SSO or password is not allowed empty value." & "</div>"
        End If


    End Sub


    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load


        'whether domain user login is valid

        If CacheFrom("sso") Is Nothing Then
            yoursso.Enabled = True
            ssopswrd.Enabled = True
            yoursso.Focus()

        Else
            yoursso.Enabled = False
            ssopswrd.Enabled = False
            LoginButton.Enabled = False
        End If




    End Sub

End Class
