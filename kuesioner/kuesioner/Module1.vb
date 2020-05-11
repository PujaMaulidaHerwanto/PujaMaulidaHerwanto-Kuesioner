Imports System.Data
Imports System.Data.OleDb

Module Module1
    Public conect As OleDbConnection
    Public cmd As OleDbCommand
    Public nds As New DataSet
    Public da As OleDbDataAdapter
    Public dr As OleDbDataReader
    Public lokasidata As String
    Public Sub konek()
        lokasidata = "provider = microsoft.jet.oledb.4.0; data source = kuesioner.mdb"
        conect = New OleDbConnection(lokasidata)
        If conect.State = ConnectionState.Closed Then
            conect.Open()
        End If

    End Sub
End Module
