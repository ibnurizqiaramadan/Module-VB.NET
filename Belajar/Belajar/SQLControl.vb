Imports MySql.Data.MySqlClient
Module SQLControl

    Dim SERVER As String = "127.0.0.1"
    Dim PORT As String = "3306"
    Dim UID As String = "root"
    Dim PASSWORD As String = ""
    Dim DATABASE As String = "db_belajar"

    Dim ConnectionString As String = "SERVER=" & SERVER & ";PORT=" & PORT & ";UID=" & UID & ";PASSWORD=" & PASSWORD & ";DATABASE=" & DATABASE & ""

    Public SqlCon As New MySqlConnection(ConnectionString)

    Public SQLCMD As MySqlCommand
    Public SQLDA As MySqlDataAdapter
    Public SQLDS As DataSet


End Module
