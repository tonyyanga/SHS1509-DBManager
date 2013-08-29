Module Module1
    Public DBpath As String
    Public DBconnection As OleDb.OleDbConnection
    Public DBReader As OleDb.OleDbDataReader
    Public DBSQL As OleDb.OleDbCommand
    Public Function JetDBconnect(ByRef JetDBPath As String) As OleDb.OleDbConnection 'JET DB　Connection
        Dim temp As Byte
        On Error GoTo err
        JetDBconnect = New System.Data.OleDb.OleDbConnection("Provider=Microsoft.Jet.OLEDB.4.0; Data Source=" & JetDBPath)
        JetDBconnect.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0; Data Source=" & JetDBPath
        Exit Function
err:
        temp = MsgBox("DB database path error!" & Chr(10) & "Wrong path:" & DBpath, vbOKOnly, "DB Error")
        End
    End Function
End Module
