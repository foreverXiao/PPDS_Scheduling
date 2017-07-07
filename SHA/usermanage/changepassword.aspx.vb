Imports basepage1
Imports System.Data.SqlClient



Partial Class usermanage_changepassword
    Inherits basepage1

    Protected Sub Page_Init1(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Init
        accesslevel = 2 ' set a low security level for this page in order to allow people to change their password
    End Sub




    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        If (Not Page.IsPostBack) AndAlso String.IsNullOrEmpty(userIden()) Then
            tbUsNm.Text = userIden()
        End If

    End Sub

    Protected Sub btSubmit_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btSubmit.Click
        Dim db As String, connstr As String
        db = ConfigurationManager.ConnectionStrings(dbConnForParam).ConnectionString
        connstr = db

        Dim conn As SqlConnection = New SqlConnection(connstr)
        Dim i As Integer = 0
        Dim cmd0 As SqlCommand = New SqlCommand()
        Dim cmd As SqlCommand = New SqlCommand()


        Try

            cmd0.Connection = conn
            cmd0.CommandText = "SELECT * FROM [Esch_Sh_tbl_userrole] WHERE [password] = '" & oldpsswrd.Text & "' AND [user_name] = '" & tbUsNm.Text & "'"


            cmd.Connection = conn
            cmd.CommandText = "UPDATE [Esch_Sh_tbl_userrole] SET [password] = '" & tbNewpswrdAgn.Text & "' WHERE [user_name] = '" & tbUsNm.Text & "'"
            

            conn.Open()

            If cmd0.ExecuteReader().Read() Then
                i = cmd.ExecuteNonQuery()
            Else
                msgPopUP("Password change failed.（Old password is incorrect）.", lbStatus, True, False)
                'lbStatus.Text = "Password change failed.（Old password is incorrect）"
            End If

            

        Finally
            If conn.State = 1 Then 'if conn is open, then close it
                conn.Close()
            End If

        End Try

        cmd0.Dispose()
        cmd.Dispose()
        conn.Dispose()


        If i = 1 Then
            msgPopUP(tbUsNm.Text & "'s password is changed sucessfully. Please remember your new password.", lbStatus, False, False)
            'lbStatus.Text = "Password changed sucessfully."
        End If



    End Sub

    Protected Sub ServerValidation(ByVal source As Object, ByVal args As ServerValidateEventArgs)
        args.IsValid = False

        Try

           
            If Len(tbUsNm.Text) <= 20 AndAlso Len(oldpsswrd.Text) <= 20 Then
                If Len(tbNewpswrd.Text) <= 20 AndAlso Len(tbNewpswrdAgn.Text) <= 20 Then
                    args.IsValid = True
                Else
                    msgPopUP("The number of characters in new password or confirmed password  exceeds 20 .", lbStatus, True, False)
                End If
            Else
                msgPopUP("The number of characters in user name or old password  exceeds 20 .", lbStatus, True, False)
            End If

        Catch ex As Exception

            args.IsValid = False

        End Try
    End Sub


End Class
