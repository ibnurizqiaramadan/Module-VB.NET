Imports MySql.Data.MySqlClient
Public Class Form1
    Private Sub Button1_Click(sender As Object, e As EventArgs)

    End Sub

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        tampilkan_data() 'ini akan error karena anda belum membuat sub baru
    End Sub

    Sub tampilkan_data()
        Try
            SqlCon.Open()

            SQLCMD = New MySqlCommand("SELECT * FROM t_user", SqlCon)
            SQLDA = New MySqlDataAdapter(SQLCMD)
            SQLDS = New DataSet
            SQLDA.Fill(SQLDS)
            If SQLDS.Tables.Count > 0 Then
                DataGridView1.DataSource = SQLDS.Tables(0)
            End If

            SqlCon.Close()
        Catch ex As Exception
            MsgBox("Gagal : " + ex.Message, MsgBoxStyle.Critical, "Gagal")
            If SqlCon.State = ConnectionState.Open Then
                SqlCon.Close()
            End If
        End Try
    End Sub
End Class
