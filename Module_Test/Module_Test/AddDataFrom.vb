Public Class AddDataFrom

    Public EDIT = False
    Public USERNAME As String
    Public ID As String

    Private Sub btnTambah_Click(sender As Object, e As EventArgs) Handles btnTambah.Click
        CekTextBoxComboBox(GroupBox1, ErrorProvider1)
        If EDIT = False Then
            If (KOSONGTEXT = False) And (KOSONGCOMBO = False) Then
                Try
                    ExecuteReader("SELECT * FROM t_user WHERE username='" & EscapeString(TextBox1.Text) & "'")
                    If Reader_kosong = True Then
                        Dim values() = {
                            INSERT_S("id"),
                            INSERT_(TextBox1),
                            INSERT_MD5(TextBox2, 5), 'lima kali enkripsi/hash md5
                            INSERT_(nama),
                            INSERT_(ComboBox1)
                        }
                        If MasukanData("t_user", DATA_(values)) = True Then
                            msgBoxInfo("Berhasil Menambahkan Data !")
                            ClearTextBoxComboBox(GroupBox1)
                            MainForm.tampildata()
                            Close()
                        End If
                    Else
                        ErrorProvider1.SetError(TextBox1, "Username sudah Digunakan !")
                    End If
                Catch ex As Exception
                    ERRORDB(ex)
                End Try
            End If
        Else
            If (KOSONGCOMBO = False) And (KOSONGTEXT = False) Then
                If USERNAME <> TextBox1.Text Then
                    ExecuteReader("SELECT * FROM t_user WHERE username='" & EscapeString(TextBox1.Text) & "'")
                    If Reader_kosong = True Then
                        EditProses()
                    Else
                        ErrorProvider1.SetError(TextBox1, "Username Sudah Digunakan !")
                        TextBox1.Select()
                    End If
                Else
                    EditProses()
                End If
            End If
        End If
    End Sub

    Sub EditProses()
        Dim values() = {
            UPDATE_("username", TextBox1),
            UPDATE_("nama", nama),
            UPDATE_("level", ComboBox1)
        }
        If EditData("t_user", DATA_(values), "id", ID) = True Then
            msgBoxInfo("Berhasil Mengedit Data !")
            ClearTextBoxComboBox(Me)
            MainForm.tampildata()
            Close()
        End If
    End Sub

    Private Sub btnEdit_Click(sender As Object, e As EventArgs) Handles btnEdit.Click
        Close()
    End Sub

    Private Sub Panel2_Paint(sender As Object, e As PaintEventArgs) Handles Panel2.Paint

    End Sub

    Private Sub AddDataFrom_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        AcceptButton = btnTambah
    End Sub

    Private Sub clearText(sender As Object, e As FormClosingEventArgs) Handles MyBase.FormClosing
        ClearTextBoxComboBox(GroupBox1)
    End Sub

    Private Sub TextBox1_KeyPress(sender As Object, e As KeyPressEventArgs) Handles TextBox1.KeyPress
        Batasi_Hanya(",./\][{}`!@|~#$%^&*()_+:;'""", e)
    End Sub

    Private Sub texbox2pencet(sender As Object, e As KeyPressEventArgs) Handles TextBox2.KeyPress
        Batasi_Kecuali("1234567890./", e)
    End Sub

End Class