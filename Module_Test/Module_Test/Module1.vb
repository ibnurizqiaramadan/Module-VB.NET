Imports MySql.Data.MySqlClient

'''     =====  MODULE MYSQL  =====    '''
'''                                   '''
''' Pembuat : Ibnu Rizqia Ramadan     '''
''' Jurusan : RPL angkatan 15         '''
''' VERSI   : 1.3.5 (BETA)            '''
'''                                   '''
'''        Terimakasih Kepada         '''
'''  Teman - teman yg sudah membantu  '''
'''     =====  MODULE MYSQL  =====    '''

Module Module1
    ''' DATABASE SETTING !!!
    Dim SERVER = "Localhost"
    Dim PORT = "3306"
    Dim UID = "root"
    Dim PASSWORD = ""
    Dim DATABASE = "db_module_test"
    Dim ZERO_DATETIME = True
    '' ' DATABASE SETTING !!!

    ''' ERROR SETTING !!!
    Dim SHOW_ERROR = True
    Dim SHOW_QUERY_ERROR = True
    ''' ERROR SETTING !!!
    '''
    ''' Pengecualian Validasi Textbox & Combobox dipisahkan menggunakan "," (nama textbox/combobox)
    Dim Pengecualian = "cari, keterangan, ket"
    ''' Pengecualian Validasi Textbox & Combobox

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
    Public ID_PENGGUNAAN As String
    Public Reader_Data(999) As String
    Public Reader_kosong As Boolean
    Public USERDATA(4) As String

    ''' DATAGRIDVIEW HEADER/CONTENT ALIGNMENT
    Private CAL As DataGridViewContentAlignment
    Public Const ML = CAL.MiddleLeft
    Public Const MC = CAL.MiddleCenter
    Public Const MR = CAL.MiddleRight
    Public Const BL = CAL.BottomLeft
    Public Const BC = CAL.BottomCenter
    Public Const BR = CAL.BottomRight
    Public Const TL = CAL.TopLeft
    Public Const TC = CAL.TopCenter
    Public Const TR = CAL.TopRight

    Sub CLEARUSERDATA()
        For i = 0 To USERDATA.Length - 1
            USERDATA(i) = ""
        Next
    End Sub

    Sub ClearReaderData()
        For i = 0 To Reader_Data.Length - 1
            Reader_Data(i) = ""
        Next
    End Sub

    Sub buka_koneksi()
        MYSQLCONNECTION.Open()
    End Sub

    Sub tutup_koneksi()
        MYSQLCONNECTION.Close()
    End Sub

    Sub KeDGV(Grid As DataGridView)
        Grid.DataSource = DS.Tables(0)
        Grid.ReadOnly = True
        Grid.SelectionMode = DataGridViewSelectionMode.FullRowSelect
        Grid.AllowUserToAddRows = False
    End Sub

    Sub DGVHeaderTxt(grid As DataGridView, columns As Integer, text As String)
        grid.Columns(columns).HeaderText = text
    End Sub

    Sub DGVContentWidth(grid As DataGridView, columns As Integer, width As Integer)
        grid.Columns(columns).Width = width
    End Sub

    Sub DGVContentAlign(grid As DataGridView, columns As Integer, align As DataGridViewContentAlignment)
        grid.Columns(columns).DefaultCellStyle.Alignment = align
    End Sub

    Sub DGVHeaderAlign(grid As DataGridView, columns As Integer, align As DataGridViewContentAlignment)
        grid.Columns(columns).HeaderCell.Style.Alignment = align
    End Sub

    Sub TampilDataTable(Tabelna As String, Grid As DataGridView)
        Try
            buka_koneksi()

            QUERY = "SELECT * FROM " & Tabelna
            MYSQLCMD = New MySqlCommand(QUERY, MYSQLCONNECTION)
            MYSQLDA = New MySqlDataAdapter(MYSQLCMD)
            DS = New DataSet
            MYSQLDA.Fill(DS)

            If DS.Tables.Count > 0 Then
                KeDGV(Grid)
            End If

            tutup_koneksi()
        Catch ex As Exception
            ERRORDB(ex)
        End Try
    End Sub

    Sub TampilDataQuery(queryna As String, Grid As DataGridView)
        Try
            buka_koneksi()

            QUERY = queryna
            MYSQLCMD = New MySqlCommand(QUERY, MYSQLCONNECTION)
            MYSQLDA = New MySqlDataAdapter(MYSQLCMD)
            DS = New DataSet
            MYSQLDA.Fill(DS)

            If DS.Tables.Count > 0 Then
                KeDGV(Grid)
            End If

            tutup_koneksi()
        Catch ex As Exception
            ERRORDB(ex)
        End Try
    End Sub

    Function SELECT_(field As String)
        Return "SELECT " & field
    End Function

    Function FROM_(table As String)
        Return " FROM " & table
    End Function

    Function WHERE_(where As String)
        If where <> "" Then
            Return " WHERE " & where & " "
        Else
            Return ""
        End If
    End Function

    Function JOIN_(table As String, joinn As String)
        Return " JOIN " & table & " ON (" & joinn & ")"
    End Function

    Function AND_(and__ As String)
        If and__ <> "" Then
            Return " AND " + Strings.Right(and__, (and__.Length - 6))
        Else
            Return ""
        End If
    End Function

    Function OR_(or__ As String)
        If or__ <> "" Then
            Return " OR " + or__
        Else
            Return ""
        End If
    End Function

    Function LIKE_(field As String, like__ As String, Optional mode As Integer = 0)
        Dim hasil = ""
        Select Case mode
            Case 0
                hasil = field & " LIKE " & "'" & like__ & "'"
            Case 1
                hasil = field & " LIKE " & "'" & like__ & "%'"
            Case 2
                hasil = field & " LIKE " & "'%" & like__ & "'"
            Case 3
                hasil = field & " LIKE " & "'%" & like__ & "%'"
            Case Else
                hasil = field & " LIKE " & "'" & like__ & "'"
        End Select
        Return hasil
    End Function

    Function INSERT_(control As Control, Optional format_ As String = "") As String
        If TypeOf control Is DateTimePicker Then
            Dim time As DateTimePicker = control
            Return "'" & EscapeString(Format(time.Value, format_)) & "', "
        ElseIf format_ <> "" Then
            Return "'" & EscapeString(Format(control.Text, format_)) & "', "
        Else
            Return "'" & EscapeString(control.Text) & "', "
        End If
    End Function

    Function INSERT_S(text As String, Optional format_ As String = "") As String
        If format_ <> "" Then
            Return "'" & EscapeString(Format(text, format_)) & "', "
        Else
            Return "'" & EscapeString(text) & "', "
        End If
    End Function

    Function INSERT_MD5(control As Control, Optional jumlahEncrypt As Integer = 1) As String
        Dim result As String = control.Text
        Dim i As Integer
        If jumlahEncrypt > 1 Then
            Return "'" & toMd5(result, jumlahEncrypt) & "', "
        Else
            Return "'" & toMd5(control.Text) & "', "
        End If
    End Function

    Function INSERT_BASE64ENC(control As Control, Optional jumlahEncrypt As Integer = 1) As String
        Dim result As String = control.Text
        Dim i As Integer
        If jumlahEncrypt > 1 Then
            Return "'" & Base64Enc(result, jumlahEncrypt) & "', "
        Else
            Return "'" & Base64Enc(control.Text) & "', "
        End If
    End Function

    Function UPDATE_MD5(field As String, control As Control, Optional jumlahEncrypt As Integer = 1) As String
        Dim result As String = control.Text
        If jumlahEncrypt > 1 Then
            Return field & " = '" & toMd5(result, jumlahEncrypt) & "', "
        Else
            Return field & " = '" & toMd5(control.Text) & "', "
        End If
    End Function

    Function UPDATE_BASE64ENC(field As String, control As Control, Optional jumlahEncrypt As Integer = 1) As String
        Dim result As String = control.Text
        If jumlahEncrypt > 1 Then
            Return field & " = '" & Base64Enc(result, jumlahEncrypt) & "', "
        Else
            Return field & " = '" & Base64Enc(control.Text) & "', "
        End If
    End Function

    Function UPDATE_(field As String, control As Control, Optional format_ As String = "") As String
        If TypeOf control Is DateTimePicker Then
            Dim time As DateTimePicker = control
            Return field & " = '" & EscapeString(Format(time.Value, format_)) & "', "
        ElseIf format_ <> "" Then
            Return field & " = '" & EscapeString(Format(control.Text, format_)) & "', "
        Else
            Return field & " = '" & EscapeString(control.Text) & "', "
        End If
    End Function

    Function UPDATE_S(field As String, text As String, Optional format_ As String = "") As String
        If format_ <> "" Then
            Return field & " = '" & EscapeString(Format(text, format_)) & "', "
        Else
            Return field & " = '" & EscapeString(text) & "', "
        End If
    End Function

    Function DATA_(array As Array)
        Dim DATA As String = ""
        For Each values In array
            DATA += values
        Next
        Return Strings.Left(DATA, DATA.Length - 2)
    End Function

    'Sub nampilkeunDataLisView(Tabelna As String)
    '    Try

    '        QUERY = "SELECT * FROM " & Tabelna
    '        MYSQLCMD = New MySqlCommand(QUERY, MYSQLCONNECTION)
    '        MYSQLDA = New MySqlDataAdapter(MYSQLCMD)
    '        DS = New DataSet
    '        MYSQLDA.Fill(DS)

    '    Catch ex As Exception
    '        ERRORDB(ex)
    '    End Try
    'End Sub

    Function ExecuteQuery(QueryEx As String) As Boolean
        Try
            buka_koneksi()

            QUERY = QueryEx
            MYSQLCMD = New MySqlCommand(QUERY, MYSQLCONNECTION)
            MYSQLCMD.ExecuteNonQuery()

            tutup_koneksi()
            Return True
        Catch ex As Exception
            ERRORDB(ex)
            Return False
        End Try
    End Function

    Function ExecuteQueryDataSet(QueryEx As String) As Boolean
        Try

            QUERY = QueryEx
            MYSQLCMD = New MySqlCommand(QUERY, MYSQLCONNECTION)
            MYSQLDA = New MySqlDataAdapter(MYSQLCMD)
            DS = New DataSet
            MYSQLDA.Fill(DS)

            Return True
        Catch ex As Exception
            ERRORDB(ex)
            Return False
        End Try
    End Function

    Sub ExecuteReader(queryread As String)
        ClearReaderData()
        Try
            buka_koneksi()
            QUERY = queryread
            MYSQLCMD = New MySqlCommand(QUERY, MYSQLCONNECTION)
            MYSQLDR = MYSQLCMD.ExecuteReader
            MYSQLDR.Read()
            If MYSQLDR.HasRows = True Then
                Reader_kosong = False
                For i = 0 To (MYSQLDR.FieldCount - 1)
                    Reader_Data(i) = MYSQLDR.Item(i)
                Next
            Else
                Reader_kosong = True
            End If
            tutup_koneksi()
        Catch ex As Exception
            ERRORDB(ex)
        End Try
    End Sub

    Function MasukanData(Tabelna As String, Isina As String) As Boolean
        Try
            buka_koneksi()

            QUERY = "INSERT INTO " & Tabelna & " VALUES (" & Isina & ")"
            MYSQLCMD = New MySqlCommand(QUERY, MYSQLCONNECTION)
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

    Function EditData(Tabelna As String, Isina As String, field As String, Optional id As String = "") As Boolean
        Try
            buka_koneksi()

            If id = "" Then
                QUERY = "UPDATE " & Tabelna & " SET " & Isina & " WHERE " & field & " = '" & EscapeString(ID_DATA) & "'"
            Else
                QUERY = "UPDATE " & Tabelna & " SET " & Isina & " WHERE " & field & " = '" & EscapeString(id) & "'"
            End If
            MYSQLCMD = New MySqlCommand(QUERY, MYSQLCONNECTION)
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

    Function HapusData(Tabelna As String, Fieldna As String, Optional Valuena As String = "") As Boolean
        Try
            buka_koneksi()

            If Valuena = "" Then
                QUERY = "DELETE FROM " & Tabelna & " WHERE " & Fieldna & " = '" & EscapeString(ID_DATA) & "'"
            Else
                QUERY = "DELETE FROM " & Tabelna & " WHERE " & Fieldna & " = '" & EscapeString(Valuena) & "'"
            End If
            MYSQLCMD = New MySqlCommand(QUERY, MYSQLCONNECTION)
            MYSQLCMD.ExecuteNonQuery()

            tutup_koneksi()
            Return True
        Catch ex As Exception
            ERRORDB(ex)
            Return False
        End Try
    End Function

    Sub MasukanDataKeComboBox(Tabelna As String, Fieldna As String, Comboboxna As ComboBox, Optional orderby As String = "ASC")
        Try
            buka_koneksi()

            QUERY = "SELECT " & Fieldna & " FROM " & Tabelna & " ORDER BY " & Fieldna & orderby
            MYSQLCMD = New MySqlCommand(QUERY, MYSQLCONNECTION)
            MYSQLDR = MYSQLCMD.ExecuteReader
            Comboboxna.Items.Clear()
            While MYSQLDR.Read
                Comboboxna.Items.Add(MYSQLDR(Fieldna))
            End While
            Comboboxna.AutoCompleteSource = AutoCompleteSource.ListItems
            Comboboxna.AutoCompleteMode = AutoCompleteMode.SuggestAppend

            tutup_koneksi()

        Catch ex As Exception
            ERRORDB(ex)
        End Try
    End Sub

    Sub AmbilID(komponen As Control, Optional Kolomkasabaraha As Integer = 0)
        If TypeOf komponen Is DataGridView Then
            Dim gridna As DataGridView
            gridna = komponen
            Try
                ID_DATA = gridna.Rows.Item(gridna.CurrentRow.Index).Cells(Kolomkasabaraha).Value
            Catch ex As Exception
                msgBoxInfo("Pilih Data terlebih dahulu !", "Pilih Data !")
                ID_DATA = ""
                gridna.Select()
            End Try
        ElseIf TypeOf komponen Is ListView Then
            Dim listv As ListView
            listv = komponen
            Try
                Dim datana As String = listv.SelectedItems.Item(0).SubItems(Kolomkasabaraha).Text
                ID_DATA = datana
            Catch ex As Exception
                msgBoxInfo("Pilih Data terlebih dahulu !", "Pilih Data !")
                ID_DATA = ""
                listv.Select()
            End Try
        End If
    End Sub

    Sub ClearTextBoxComboBox(control As Control)
        ClearTextBox(control)
        ClearComboBox(control)
    End Sub

    Private Sub ClearComboBox(control As Control)
        For Each control In control.Controls
            If TypeOf control Is ComboBox Then
                control.Text = ""
            End If
        Next control
    End Sub

    Private Sub ClearTextBox(control As Control)
        For Each control In control.Controls
            If TypeOf control Is TextBox Then
                control.Text = ""
            End If
        Next control
    End Sub

    Sub CekTextBoxComboBox(control As Control, ErrorP As ErrorProvider)
        ErrorP.Dispose()
        CekComboBox(control, ErrorP)
        CekTextBox(control, ErrorP)
    End Sub

    Private Sub CekTextBox(control As Control, ErrorP As ErrorProvider)
        DATA1 = ""
        For Each control In control.Controls
            If TypeOf control Is TextBox Then
                If Not CBool(InStr(Pengecualian, control.Name.ToLower)) Then
                    If control.Text.Trim = "" Then
                        Dim ss As String = "B"
                        control.Select()
                        ErrorP.SetError(control, "Harap isi bidang ini !")
                        DATA1 += ss
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
        Next control
    End Sub

    Private Sub cekText()
        If DATA1 <> UdahText Then
            KOSONGTEXT = True
        Else
            KOSONGTEXT = False
        End If
    End Sub

    Private Sub CekComboBox(control As Control, ErrorP As ErrorProvider)
        DATA2 = ""
        For Each control In control.Controls
            If TypeOf control Is ComboBox Then
                If Not CBool(InStr(Pengecualian, control.Name.ToLower)) Then
                    If control.Text.Trim = "" Then
                        Dim combos As Char = "B"
                        control.Select()
                        ErrorP.SetError(control, "Harap isi bidang ini !")
                        DATA2 += combos
                        UdahCombo = New String("U", DATA2.Length)
                        cekCombo()
                    Else
                        Dim combos As Char = "U"
                        DATA2 = DATA2 + combos
                        UdahCombo = New String("U", DATA2.Length)
                        cekCombo()
                    End If
                End If
            End If
        Next control
    End Sub

    Private Sub cekCombo()
        If DATA2 <> UdahCombo Then
            KOSONGCOMBO = True
        Else
            KOSONGCOMBO = False
        End If
    End Sub

    Sub msgBoxInfo(Pesan As String, Optional Judul As String = "Info")
        MsgBox(Pesan, MsgBoxStyle.Information + MsgBoxStyle.OkOnly, Judul)
    End Sub

    Sub msgBoxPeringatan(Pesan As String, Optional Judul As String = "Peringatan")
        MsgBox(Pesan, MsgBoxStyle.Exclamation + MsgBoxStyle.OkOnly, Judul)
    End Sub

    Sub msgBoxError(Pesan As String, Optional Judul As String = "Error")
        MsgBox(Pesan, MsgBoxStyle.Critical + MsgBoxStyle.OkOnly, Judul)
    End Sub

    Function msgBoxKonfir(pesan As String, Optional Judul As String = "Konfirmasi") As MsgBoxResult
        Return MsgBox(pesan, MsgBoxStyle.Question + MsgBoxStyle.YesNo, Judul)
    End Function

    Sub ERRORDB(ex As Exception)
        If SHOW_ERROR = True Then
            If SHOW_QUERY_ERROR = True Then
                If QUERY <> "" Then
                    msgBoxError("GAGAL, karena " + ex.Message + " | QUERY :  ( " + QUERY + " ) ", "Gagal !")
                    If MYSQLCONNECTION.State = ConnectionState.Open Then
                        tutup_koneksi()
                    End If
                Else
                    msgBoxError("GAGAL, karena " + ex.Message, "Gagal !")
                    If MYSQLCONNECTION.State = ConnectionState.Open Then
                        tutup_koneksi()
                    End If
                End If
            Else
                msgBoxError("GAGAL, karena " + ex.Message, "Gagal !")
                If MYSQLCONNECTION.State = ConnectionState.Open Then
                    tutup_koneksi()
                End If
            End If
        End If
    End Sub

    Function EscapeString(text As String) As String
        Try
            Return MySqlHelper.EscapeString(text)
        Catch ex As Exception
            Return ""
        End Try
    End Function

    Function EscapeString1(text As String) As String
        Dim result As String
        result = text.Replace("\", "\\")
        result = result.Replace("'", "\'")
        result = result.Replace("""", "\""")
        result = result.Replace("`", "\`")
        Return result
    End Function

    Function toMd5(ByVal textToHash As String, Optional jumlahEncrypt As Integer = 1) As String
        Dim result As String = textToHash
        Try
            If jumlahEncrypt > 1 Then
                For i As Integer = 1 To jumlahEncrypt
                    Dim MD5 As New System.Security.Cryptography.MD5CryptoServiceProvider()
                    Dim Bytes() As Byte = MD5.ComputeHash(System.Text.Encoding.ASCII.GetBytes(result))
                    Dim s As String = Nothing
                    For Each by As Byte In Bytes
                        s += by.ToString("x2")
                    Next
                    result = s
                Next
                Return result
            Else
                Dim MD5 As New System.Security.Cryptography.MD5CryptoServiceProvider()
                Dim Bytes() As Byte = MD5.ComputeHash(System.Text.Encoding.ASCII.GetBytes(textToHash))
                Dim s As String = Nothing
                For Each by As Byte In Bytes
                    s += by.ToString("x2")
                Next
                result = s
                Return result
            End If
        Catch ex As Exception
            Return textToHash
        End Try
    End Function

    Function AmbilData(control As Control, subitem As Integer) As String
        Try
            If TypeOf control Is DataGridView Then
                Dim datagrid As DataGridView
                datagrid = control
                Return datagrid.Rows.Item(datagrid.CurrentRow.Index).Cells(subitem).Value
            ElseIf TypeOf control Is ListView Then
                Dim listvi As ListView
                listvi = control
                Dim datana As String = listvi.SelectedItems.Item(0).SubItems(subitem).Text
                Return datana
            Else
                Return Nothing
            End If
        Catch ex As Exception
            Return Nothing
        End Try
    End Function

    Function Base64Enc(text As String, Optional JumlahEncrypt As Integer = 1) As String
        Dim result As String = text
        Try
            If JumlahEncrypt > 1 Then
                For i As Integer = 1 To JumlahEncrypt
                    Dim encript As Byte() = System.Text.Encoding.UTF8.GetBytes(result)
                    result = Convert.ToBase64String(encript)
                Next
                Return result
            Else
                Dim encript As Byte() = System.Text.Encoding.UTF8.GetBytes(result)
                Return Convert.ToBase64String(encript)
            End If
        Catch ex As Exception
            Return text
        End Try
    End Function

    Function Base64Dec(text As String, Optional JumlahDecrypt As Integer = 1)
        Dim result As String = text
        Try
            If JumlahDecrypt > 1 Then
                For i As Integer = 1 To JumlahDecrypt
                    Dim decript As Byte() = Convert.FromBase64String(result)
                    result = System.Text.Encoding.UTF8.GetString(decript)
                Next
                Return result
            Else
                Dim decript As Byte() = Convert.FromBase64String(result)
                Return System.Text.Encoding.UTF8.GetString(decript)
            End If
        Catch ex As Exception
            Return text
        End Try
    End Function

    Sub NumberOnly(e As KeyPressEventArgs)
        If Not ((e.KeyChar >= "0" And e.KeyChar <= "9") Or e.KeyChar = vbBack) Then
            e.Handled = True
        End If
    End Sub

    Sub Batasi_Kecuali(Character As String, e As KeyPressEventArgs)
        If Not (Character.Contains(e.KeyChar) Or e.KeyChar = vbBack Or e.KeyChar = Convert.ToChar(32)) Then
            e.Handled = True
        End If
    End Sub

    Sub Batasi_Hanya(Character As String, e As KeyPressEventArgs)
        If (Character.Contains(e.KeyChar)) Then
            e.Handled = True
        End If
    End Sub

End Module
