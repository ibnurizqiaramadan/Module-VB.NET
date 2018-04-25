Public Class MainForm
    Private Sub MainForm_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        tampildata()
        aturgrid()
    End Sub

    Sub tampildata(Optional where = "")
        Try
            QUERY = SELECT_("id, username, password, nama, level") +
                    FROM_("t_user") +
                    WHERE_(where)
            TampilDataQuery(QUERY, DataGridView1)
        Catch ex As Exception
            ERRORDB(ex)
        End Try
    End Sub

    Sub aturgrid()
        Dim HEADER() = {"ID", "USERNAME", "PASSWORD", "NAMA", "LEVEL"}
        Dim LEBAR() = {50, 150, 150, 200, 50}
        Dim HEADERAL() = {ML, MC, MC, MC, MC}
        Dim CONTENTAL() = {ML, ML, ML, ML, MC}
        For i = 0 To DataGridView1.ColumnCount - 1
            DGVHeaderTxt(DataGridView1, i, HEADER(i))
            DGVContentWidth(DataGridView1, i, LEBAR(i))
            DGVHeaderAlign(DataGridView1, i, HEADERAL(i))
            DGVContentAlign(DataGridView1, i, CONTENTAL(i))
        Next
        DataGridView1.Columns(0).Visible = False
    End Sub

    Private Sub btnTambah_Click(sender As Object, e As EventArgs) Handles btnTambah.Click
        With AddDataFrom
            .TextBox2.Enabled = True
            .Label1.Text = "Tambah Data User"
            .ShowDialog()
        End With
    End Sub

    Private Sub btnEdit_Click(sender As Object, e As EventArgs) Handles btnEdit.Click
        With AddDataFrom
            .Label1.Text = "Edit Data User"
            .TextBox1.Text = AmbilData(DataGridView1, 1)
            .TextBox2.Text = "Password . . ."
            .TextBox2.Enabled = False
            .nama.Text = AmbilData(DataGridView1, 3)
            .ComboBox1.Text = AmbilData(DataGridView1, 4)
            .USERNAME = AmbilData(DataGridView1, 1)
            .ID = AmbilData(DataGridView1, 0)
            .EDIT = True
            .ShowDialog()
        End With
    End Sub

    Private Sub MainForm_KeyPress(sender As Object, e As KeyPressEventArgs) Handles MyBase.KeyPress
        If e.KeyChar = Convert.ToChar(116) Then
            tampildata()
        End If
    End Sub

    Private Sub btnHapus_Click(sender As Object, e As EventArgs) Handles btnHapus.Click
        AmbilID(DataGridView1, 0) 'untuk mengambil ID yang akan di hapus
        If msgBoxKonfir("Apakah Anda yakin ingin menghapus Data Terpilih ?") = MsgBoxResult.Yes Then
            If HapusData("t_user", "id") = True Then
                msgBoxInfo("Berhasil Menghapus Data", "Berhasil")
                tampildata()
            End If
        End If
    End Sub

    Private Sub TextBox1_TextChanged(sender As Object, e As EventArgs) Handles TextBox1.TextChanged
        If TextBox1.Text.Trim <> "" Then
            tampildata(LIKE_("username", EscapeString(TextBox1.Text), 1) & OR_(LIKE_("level", EscapeString(TextBox1.Text), 1)))
        Else
            tampildata()
        End If
    End Sub
End Class