Module Module1
    Public DBpath As String
    Public DBconnection As odbc.odbcConnection
    Public DBReader As odbc.odbcDataReader
    Public DBSQL As odbc.odbcCommand
    Public Function JetDBconnect(ByRef Server As String, ByRef user As String, ByRef password As String) As Odbc.OdbcConnection 'JET DB　Connection
        Dim temp As Byte
        On Error GoTo err
        JetDBconnect = New System.Data.Odbc.OdbcConnection("DRIVER={MySQL ODBC 3.51 Driver};SERVER=" & Server & ";DATABASE=shs1509dbmanager;user=" & user & ";PASSWORD=" & password & ";OPTION=3;")
        JetDBconnect.ConnectionString = "DRIVER={MySQL ODBC 3.51 Driver};SERVER=" & Server & ";DATABASE=shs1509dbmanager;user=" & user & ";PASSWORD=" & password & ";OPTION=3;"
        Exit Function
err:
        temp = MsgBox("DB database path error!" & Chr(10) & "Wrong path:" & DBpath, vbOKOnly, "DB Error")
        End
    End Function
End Module
