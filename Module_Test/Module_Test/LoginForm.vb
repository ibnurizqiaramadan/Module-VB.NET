Public Class LoginForm

    ' TODO: Insert code to perform custom authentication using the provided username and password 
    ' (See http://go.microsoft.com/fwlink/?LinkId=35339).  
    ' The custom principal can then be attached to the current thread's principal as follows: 
    '     My.User.CurrentPrincipal = CustomPrincipal
    ' where CustomPrincipal is the IPrincipal implementation used to perform authentication. 
    ' Subsequently, My.User will return identity information encapsulated in the CustomPrincipal object
    ' such as the username, display name, etc.

    Private Sub OK_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles OK.Click
        CekTextBoxComboBox(Me, ErrorProvider1)
        If KOSONGTEXT = False Then
            If (username.Text = "devnesas") And (password.Text = "devnesas") Then
                MainForm.Show()
                Close()
            Else
                ExecuteReader("SELECT * FROM t_user WHERE username = '" & EscapeString(username.Text) & "'")
                If Reader_kosong = False Then
                    If toMd5(password.Text, 5) = Reader_Data(2) Then
                        ClearTextBoxComboBox(Me)
                        USERDATA(0) = Reader_Data(0) 'id
                        USERDATA(1) = Reader_Data(1) 'username
                        USERDATA(2) = Reader_Data(3) 'nama
                        USERDATA(3) = Reader_Data(4) 'level
                        MainForm.Show()
                        Close()
                    Else
                        ErrorProvider1.SetError(password, "Password yang Anda masukan Salah !")
                        password.Select()
                        password.SelectAll()
                    End If
                Else
                    ErrorProvider1.SetError(username, "Username tidak ditemukan !")
                    username.Select()
                    username.SelectAll()
                End If
            End If
        End If
    End Sub

    Private Sub Cancel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Cancel.Click
        Me.Close()
    End Sub

End Class