Imports MySql.Data.MySqlClient


'''     =====  MODULE MYSQL  =====    '''
'''                                   '''
''' Pembuat : Ibnu Rizqia Ramadan     '''
''' Jurusan : RPL angkatan 15         '''
''' VERSI   : 1.0.2 (BETA)            '''
'''                                   '''
'''        Terimakasih Kepada         '''
'''  Teman - teman yg sudah membantu  '''
'''     =====  MODULE MYSQL  =====    '''


Module SQLControl

    Dim IsMysql As Boolean = True
    Dim SERVER As String = "Localhost"
    Dim PORT As String = "3306"
    Dim UID As String = "root"
    Dim PASSWORD As String = ""
    Dim DATABASE As String = "db_koperasi_vs"
    Dim ZERO_DATETIME As Boolean = True

    Dim CONNECTIONSTRING As String = "SERVER=" & SERVER & ";PORT=" & PORT & ";UID=" & UID & ";PASSWORD=" & PASSWORD & ";DATABASE=" & DATABASE & ";Convert Zero Datetime=" & ZERO_DATETIME & ""

    Public MYSQLCONNECTION As New MySqlConnection(CONNECTIONSTRING)
    Public MYSQLCMD As MySqlCommand
    Public MYSQLDA As MySqlDataAdapter
    Public MYSQLDR As MySqlDataReader

    Public DS As DataSet
    Public DT As DataTable
    Public QUERY As String
    Public ID_DATA As String
    Public KOSONGTEXT As Boolean
    Public KOSONGCOMBO As Boolean
    Public DATA As String
    Public DATA1 As String
    Public DATA2 As String
    Public UdahText As String
    Public UdahCombo As String
    Public Helper As MySqlHelper

    Public USERDATA(3) As String

    Public Sub CLEARUSERDATA()
        Dim i As Integer
        For i = 0 To USERDATA.Length - 1
            USERDATA(i) = ""
        Next
    End Sub

    Public Sub buka_koneksi()
        MYSQLCONNECTION.Open()
    End Sub

    Public Sub tutup_koneksi()
        MYSQLCONNECTION.Close()
    End Sub

    Public Sub MunculkeunFormNoMaxMin(formna As Form, main As Form)
        formna.StartPosition = FormStartPosition.CenterParent
        formna.FormBorderStyle = FormBorderStyle.FixedSingle
        formna.MaximizeBox = False
        formna.MinimizeBox = False
        formna.Show()
        formna.MdiParent = main
    End Sub

    Public Sub MunculkeunFormNoMdi(formna As Form)
        formna.StartPosition = FormStartPosition.CenterParent
        formna.FormBorderStyle = FormBorderStyle.FixedSingle
        formna.MaximizeBox = False
        formna.MinimizeBox = False
        formna.ShowDialog()
    End Sub

    Public Sub MunculkeunForm(formna As Form)
        formna.StartPosition = FormStartPosition.CenterParent
        formna.FormBorderStyle = FormBorderStyle.FixedSingle
        formna.ShowDialog()
    End Sub

    Public Sub nampilkeunData(Tabelna As String, Grid As DataGridView)
        If IsMysql = True Then
            Try
                buka_koneksi()

                MYSQLCMD = New MySqlCommand("SELECT * FROM " & Tabelna, MYSQLCONNECTION)
                MYSQLDA = New MySqlDataAdapter(MYSQLCMD)
                DS = New DataSet
                MYSQLDA.Fill(DS)
                If DS.Tables.Count > 0 Then
                    Grid.DataSource = DS.Tables(0)
                    Grid.ReadOnly = True
                End If

                tutup_koneksi()
            Catch ex As Exception
                ERRORDB(ex)
            End Try
        End If
    End Sub

    ''' ========== NAMPILKEUNDATA ========== '''
    ''' 
    ''' CARA PENGGUNAAN
    '''
    ''' !!! nampilkeundata(parameter_pertama, parameter_kedua) !!!
    ''' 
    ''' parameter_pertama = diisi dengan nama tabel yang akan di tampilkan 
    ''' parameter_kedua   = diisi dengan nama DataGridView yang akan digunakan untuk menampilkan data yang ada ditabel parameter_pertama
    ''' 
    ''' !!! CONTOH !!!
    ''' 
    ''' nampilkeundata("t_user", DataGridView1)
    ''' 

    Public Sub nampilkeunDataWhere(Tabelna As String, where As String, Grid As DataGridView)
        Try
            buka_koneksi()

            MYSQLCMD = New MySqlCommand("SELECT * FROM " & Tabelna & " WHERE " & where, MYSQLCONNECTION)
            MYSQLDA = New MySqlDataAdapter(MYSQLCMD)
            DS = New DataSet
            MYSQLDA.Fill(DS)
            If DS.Tables.Count > 0 Then
                Grid.DataSource = DS.Tables(0)
                Grid.ReadOnly = True
                DATA = ""
            End If

            tutup_koneksi()
        Catch ex As Exception
            ERRORDB(ex)
        End Try
    End Sub

    ''' ========== NAMPILKEUNDATAWHERE ========== '''
    ''' 
    ''' CARA PENGGUNAAN
    '''
    ''' !!! nampilkeundata(parameter_pertama, parameter_kedua) !!!
    ''' 
    ''' parameter_pertama = diisi dengan nama tabel yang akan di tampilkan 
    ''' parameter_kedua   = diisi dengan nama DataGridView yang akan digunakan untuk menampilkan data yang ada ditabel parameter_pertama
    ''' 
    ''' !!! CONTOH !!!
    ''' 
    ''' nampilkeundata("t_user", DataGridView1)
    ''' 

    Public Sub nampilkeunDataLisView(Tabelna As String)
        Try

            MYSQLCMD = New MySqlCommand("SELECT * FROM " & Tabelna, MYSQLCONNECTION)
            MYSQLDA = New MySqlDataAdapter(MYSQLCMD)
            DT = New DataTable
            MYSQLDA.Fill(DT)

        Catch ex As Exception
            ERRORDB(ex)
        End Try
    End Sub

    Public Sub nampilkeunDataLisViewWhere(Tabelna As String, where As String)
        Try

            MYSQLCMD = New MySqlCommand("SELECT * FROM " & Tabelna & " WHERE " & where, MYSQLCONNECTION)
            MYSQLDA = New MySqlDataAdapter(MYSQLCMD)
            DT = New DataTable
            MYSQLDA.Fill(DT)

        Catch ex As Exception
            ERRORDB(ex)
        End Try
    End Sub

    Public Sub nampilkeunDatUrut(Tabelna As String, fieldna As String, Ngurutkeuna As String, Grid As DataGridView)
        Try
            buka_koneksi()

            MYSQLCMD = New MySqlCommand("SELECT * FROM " & Tabelna & " ORDER BY " & fieldna & "  " & Ngurutkeuna, MYSQLCONNECTION)
            MYSQLDA = New MySqlDataAdapter(MYSQLCMD)
            DS = New DataSet
            MYSQLDA.Fill(DS)
            If DS.Tables.Count > 0 Then
                Grid.DataSource = DS.Tables(0)
            End If

            tutup_koneksi()
        Catch ex As Exception
            ERRORDB(ex)
        End Try
    End Sub

    Public Function ExecuteQuery(QueryEx As String) As Boolean
        Try
            buka_koneksi()

            MYSQLCMD = New MySqlCommand(QueryEx, MYSQLCONNECTION)
            MYSQLCMD.ExecuteNonQuery()

            tutup_koneksi()
            Return True
        Catch ex As Exception
            ERRORDB(ex)
            Return False
        End Try
    End Function

    Public Sub ExecuteReader(QueryEx As String)
        Try

            MYSQLCMD = New MySqlCommand(QueryEx, MYSQLCONNECTION)
            MYSQLDR = MYSQLCMD.ExecuteReader
            MYSQLDR.Read()

        Catch ex As Exception
            ERRORDB(ex)
        End Try
    End Sub

    Public Function NgasupkeunData(Tabelna As String, Isina As String) As Boolean
        Try
            buka_koneksi()

            MYSQLCMD = New MySqlCommand("INSERT INTO " & Tabelna & " VALUES (" & Isina & ")", MYSQLCONNECTION)
            MYSQLCMD.ExecuteNonQuery()
            DATA = ""

            tutup_koneksi()
            Return True
        Catch ex As Exception
            DATA = ""
            ERRORDB(ex)
            Return False
        End Try
    End Function

    Public Function NgeditData(Tabelna As String, Isina As String, field As String, id As String) As Boolean
        Try
            buka_koneksi()

            MYSQLCMD = New MySqlCommand("UPDATE " & Tabelna & " SET " & Isina & " WHERE " & field & " = '" & Helper.EscapeString(id) & "'", MYSQLCONNECTION)
            MYSQLCMD.ExecuteNonQuery()
            DATA = ""

            tutup_koneksi()
            Return True
        Catch ex As Exception
            DATA = ""
            ERRORDB(ex)
            Return False
        End Try
    End Function

    Public Function NgahapusData(Tabelna As String, Fieldna As String, Valuena As String) As Boolean
        Try
            buka_koneksi()

            MYSQLCMD = New MySqlCommand("DELETE FROM " & Tabelna & " WHERE " & Fieldna & " = '" & Helper.EscapeString(Valuena) & "'", MYSQLCONNECTION)
            MYSQLCMD.ExecuteNonQuery()

            tutup_koneksi()
            Return True
        Catch ex As Exception
            ERRORDB(ex)
            Return False
        End Try
    End Function

    Public Sub ngasupkeunkaComboBoxASC(Tabelna As String, Fieldna As String, Comboboxna As ComboBox)
        Try
            buka_koneksi()

            MYSQLCMD = New MySqlCommand("SELECT " & Fieldna & " FROM " & Tabelna & " ORDER BY " & Fieldna & " ASC", MYSQLCONNECTION)
            MYSQLDR = MYSQLCMD.ExecuteReader

            While MYSQLDR.Read
                Comboboxna.Items.Add(MYSQLDR(Fieldna))
            End While
            tutup_koneksi()

        Catch ex As Exception
            ERRORDB(ex)
        End Try
    End Sub

    Public Sub ngasupkeunkaComboBoxDESC(Tabelna As String, Fieldna As String, Comboboxna As ComboBox)
        Try
            buka_koneksi()

            MYSQLCMD = New MySqlCommand("SELECT " & Fieldna & " FROM " & Tabelna & " ORDER BY " & Fieldna & " DESC", MYSQLCONNECTION)
            MYSQLDR = MYSQLCMD.ExecuteReader

            While MYSQLDR.Read
                Comboboxna.Items.Add(MYSQLDR(Fieldna))
            End While
            tutup_koneksi()

        Catch ex As Exception
            ERRORDB(ex)
        End Try
    End Sub

    Public Sub ngasupkeunkaComboBox(Tabelna As String, Fieldna As String, Comboboxna As ComboBox, orderby As String)
        Try
            buka_koneksi()

            MYSQLCMD = New MySqlCommand("SELECT " & Fieldna & " FROM " & Tabelna & " ORDER BY " & Fieldna & " " & orderby, MYSQLCONNECTION)
            MYSQLDR = MYSQLCMD.ExecuteReader

            While MYSQLDR.Read
                Comboboxna.Items.Add(MYSQLDR(Fieldna))
            End While
            tutup_koneksi()

        Catch ex As Exception
            ERRORDB(ex)
        End Try
    End Sub

    Public Function nyokotIDna(komponen As Control, Kolomkasabaraha As Integer)
        If TypeOf komponen Is DataGridView Then
            Dim gridna As DataGridView
            gridna = komponen
            Try
                Return gridna.Rows.Item(gridna.CurrentRow.Index).Cells(Kolomkasabaraha).Value
            Catch ex As Exception
                msgBoxInfo("Pilih Data terlebih dahulu !", "Pilih Data !")
                gridna.Select()
            End Try
        ElseIf TypeOf komponen Is ListView Then
            Dim listv As ListView
            listv = komponen
            Try
                Dim datana As String = listv.SelectedItems.Item(0).SubItems(Kolomkasabaraha).Text
                Return datana
            Catch ex As Exception
                msgBoxInfo("Pilih Data terlebih dahulu !", "Pilih Data !")
                listv.Select()
            End Try
        End If
    End Function

    Public Sub NgosongkeunComboBox(Formna As Form)
        For Each Control In Formna.Controls
            If TypeOf Control Is ComboBox Then
                Control.Text = ""
            End If
        Next Control
    End Sub

    Public Sub NgosongkeunTextBox(Formna As Form)
        For Each Control In Formna.Controls
            If TypeOf Control Is TextBox Then
                Control.Text = ""
            End If
        Next Control
    End Sub

    Public Sub NgumpulkeunData(arrayna As Array)
        For Each value As String In arrayna
            DATA = DATA + value
        Next
    End Sub

    Public Sub NgecekTextBox(Formna As Form, ErrorP As ErrorProvider)
        DATA1 = ""
        ErrorP.Dispose()
        For Each Control In Formna.Controls
            If TypeOf Control Is TextBox Then
                If (Control.Name <> "cari") And (Control.Name <> "keterangan") Then
                    If Control.Text = "" Then
                        Dim ss As String = "B"
                        Control.Select
                        ErrorP.SetError(Control, "Harap isi bidang ini !")
                        DATA1 = DATA1 + ss
                        UdahText = New String("U", DATA1.Length)
                        cekText()
                    Else
                        Dim ss As String = "U"
                        DATA1 = DATA1 + ss
                        UdahText = New String("U", DATA1.Length)
                        cekText()
                    End If
                End If
            End If
        Next Control
    End Sub

    Sub cekText()
        If DATA1 <> UdahText Then
            KOSONGTEXT = True
        Else
            KOSONGTEXT = False
        End If
    End Sub

    Public Sub NgecekComboBox(Formna As Form, ErrorP As ErrorProvider)
        DATA2 = ""
        For Each Control In Formna.Controls
            If TypeOf Control Is ComboBox Then
                If Control.Text = "" Then
                    Dim combos As Char = "B"
                    Control.Select
                    ErrorP.SetError(Control, "Harap isi bidang ini !")
                    DATA2 = DATA2 + combos
                    UdahCombo = New String("U", DATA2.Length)
                    cekCombo()
                Else
                    Dim combos As Char = "U"
                    DATA2 = DATA2 + combos
                    UdahCombo = New String("U", DATA2.Length)
                    cekCombo()
                End If
            End If
        Next Control
    End Sub

    Sub cekCombo()
        If DATA2 <> UdahCombo Then
            KOSONGCOMBO = True
        Else
            KOSONGCOMBO = False
        End If
    End Sub

    Public Sub msgBoxInfo(Pesan As String, Judul As String)
        MsgBox(Pesan, MsgBoxStyle.Information + MsgBoxStyle.OkOnly, Judul)
    End Sub

    Public Sub msgBoxPeringatan(Pesan As String, Judul As String)
        MsgBox(Pesan, MsgBoxStyle.Exclamation + MsgBoxStyle.OkOnly, Judul)
    End Sub

    Public Sub msgBoxError(Pesan As String, Judul As String)
        MsgBox(Pesan, MsgBoxStyle.Critical + MsgBoxStyle.OkOnly, Judul)
    End Sub

    Public Sub ERRORDB(ex As Exception)
        msgBoxError("GAGAL : " + ex.Message, "Gagal !")
        If MYSQLCONNECTION.State = ConnectionState.Open Then
            tutup_koneksi()
        End If
    End Sub

    Public Function toMd5(ByVal textToHash As String) As String
        Dim MD5 As New System.Security.Cryptography.MD5CryptoServiceProvider()
        Dim Bytes() As Byte = MD5.ComputeHash(System.Text.Encoding.ASCII.GetBytes(textToHash))
        Dim s As String = Nothing
        For Each by As Byte In Bytes
            s += by.ToString("x2")
        Next
        Return s
    End Function

    Public Function cokotdata(controlna As Control, subitem As Integer) As String
        If TypeOf controlna Is DataGridView Then
            Dim datagrid As DataGridView
            datagrid = controlna
            Return datagrid.Rows.Item(datagrid.CurrentRow.Index).Cells(subitem).Value
        ElseIf TypeOf controlna Is ListView Then
            Dim listvi As ListView
            listvi = controlna
            Dim datana As String = listvi.SelectedItems.Item(0).SubItems(subitem).Text
            Return datana
        End If
    End Function

    Public Function Base64Enc(text As String)
        Try
            Dim encript As Byte() = System.Text.Encoding.UTF8.GetBytes(text)
            Return Convert.ToBase64String(encript)
        Catch ex As Exception
            Return text
        End Try
    End Function

    Public Function Base64Dec(text As String)
        Try
            Dim decript As Byte() = Convert.FromBase64String(text)
            Return System.Text.Encoding.UTF8.GetString(decript)
        Catch ex As Exception
            Return text
        End Try
    End Function

End Module