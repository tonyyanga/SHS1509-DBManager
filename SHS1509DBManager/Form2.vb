Imports Excel = Microsoft.Office.Interop.Excel
Public Class Form2
    Private Sub Form2_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Call Tab1_Initialization()
        Call Tab2_Initialization()
    End Sub
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        If ListBox1.SelectedItem <> Nothing Then
            ListBox2.Items.Add(ListBox1.SelectedItem.ToString)
            ListBox1.Items.Remove(ListBox1.SelectedItem)
        End If
    End Sub
    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        If ListBox2.SelectedItem <> Nothing Then
            ListBox1.Items.Add(ListBox2.SelectedItem.ToString)
            ListBox2.Items.Remove(ListBox2.SelectedItem)
        End If
    End Sub
    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        Dim i As Integer
        For i = (ListBox1.Items.Count - 1) To 0 Step -1
            ListBox2.Items.Add(ListBox1.Items(i).ToString)
            ListBox1.Items.RemoveAt(i)
        Next i
    End Sub
    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        Dim i As Integer
        For i = (ListBox2.Items.Count - 1) To 0 Step -1
            ListBox1.Items.Add(ListBox2.Items(i).ToString)
            ListBox2.Items.RemoveAt(i)
        Next i
    End Sub
    Private Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click
        Dim StudentNo As String, day As String, i As Integer, j As Integer
        DBconnection.Open()
        day = Microsoft.VisualBasic.Left(MonthCalendar1.SelectionStart.ToString, Len(MonthCalendar1.SelectionStart.ToString) - 9)
        'Convert Room No to StudentName
        For i = (ListBox2.Items.Count - 1) To 0 Step -1
            If IsNumeric(ListBox2.Items(i).ToString) Then
                DBSQL = New OleDb.OleDbCommand("Select * From Room Where Room = " & ListBox2.Items(i).ToString, DBconnection)
                DBReader = DBSQL.ExecuteReader
                DBReader.Read()
                For j = 3 To 6
                    If DBReader.Item("StudentName" & CStr(j - 2)).ToString <> "" Then
                        ListBox2.Items.Add(DBReader.Item("StudentName" & CStr(j - 2)).ToString)
                    End If
                Next j
                ListBox2.Items.RemoveAt(i)
            End If
        Next i
        'Insert Records
        If ListBox2.Items.Count > 0 And IsNumeric(TextBox2.Text) Then
            For i = (ListBox2.Items.Count - 1) To 0 Step -1
                DBSQL = New OleDb.OleDbCommand("Select * From students Where studentname ='" _
                & ListBox2.Items(i).ToString & "'", DBconnection)
                DBReader = DBSQL.ExecuteReader
                DBReader.Read()
                StudentNo = DBReader.Item("studentno").ToString
                DBSQL = New OleDb.OleDbCommand("Insert into [general] ([Time], StudentNo, [Number], Reason, Person) Values (" _
                & day & "," & _
                StudentNo & "," & TextBox2.Text & ",'" & TextBox3.Text & "','" & TextBox4.Text & "')", DBconnection)
                DBSQL.ExecuteNonQuery()
            Next i
        End If
        DBconnection.Close()
        TextBox2.Text = ""
        Call Tab1_Initialization()
    End Sub
    Private Sub Tab1_Initialization()
        Dim i As Integer
        'Init tab1.
        MonthCalendar1.TodayDate = Now
        'MsgBox(DBconnection.ConnectionString)
        'Load Students' Names
        For i = (ListBox1.Items.Count - 1) To 0 Step -1
            ListBox1.Items.RemoveAt(i)
        Next
        For i = (ListBox2.Items.Count - 1) To 0 Step -1
            ListBox2.Items.RemoveAt(i)
        Next
        DBconnection.Open()
        DBSQL = New OleDb.OleDbCommand("Select * From students", DBconnection)
        DBReader = DBSQL.ExecuteReader
        Do While DBReader.Read()
            ListBox1.Items.Add(DBReader.Item("studentname"))
        Loop
        DBSQL = New OleDb.OleDbCommand("Select * From Room", DBconnection)
        DBReader = DBSQL.ExecuteReader
        Do While DBReader.Read()
            ListBox1.Items.Add(DBReader.Item("Room"))
        Loop
        DBconnection.Close()
    End Sub
    Private Sub Tab2_Initialization()
        Dim i As Integer
        'Init tab2.
        MonthCalendar2.TodayDate = Now
        'MsgBox(DBconnection.ConnectionString)
        'Load Students' Names
        For i = (ListBox3.Items.Count - 1) To 0 Step -1
            ListBox3.Items.RemoveAt(i)
        Next
        For i = (ListBox4.Items.Count - 1) To 0 Step -1
            ListBox4.Items.RemoveAt(i)
        Next
        DBconnection.Open()
        DBSQL = New OleDb.OleDbCommand("Select * From students", DBconnection)
        DBReader = DBSQL.ExecuteReader
        Do While DBReader.Read()
            ListBox4.Items.Add(DBReader.Item("studentname"))
        Loop
        DBconnection.Close()
    End Sub
    Private Sub Button10_Click(sender As Object, e As EventArgs) Handles Button10.Click
        If ListBox4.SelectedItem <> Nothing Then
            ListBox3.Items.Add(ListBox4.SelectedItem.ToString)
            ListBox4.Items.Remove(ListBox4.SelectedItem)
        End If
    End Sub
    Private Sub Button9_Click(sender As Object, e As EventArgs) Handles Button9.Click
        If ListBox3.SelectedItem <> Nothing Then
            ListBox4.Items.Add(ListBox3.SelectedItem.ToString)
            ListBox3.Items.Remove(ListBox3.SelectedItem)
        End If
    End Sub
    Private Sub Button8_Click(sender As Object, e As EventArgs) Handles Button8.Click
        Dim i As Integer
        For i = (ListBox4.Items.Count - 1) To 0 Step -1
            ListBox3.Items.Add(ListBox4.Items(i).ToString)
            ListBox4.Items.RemoveAt(i)
        Next i
    End Sub
    Private Sub Button7_Click(sender As Object, e As EventArgs) Handles Button7.Click
        Dim i As Integer
        For i = (ListBox3.Items.Count - 1) To 0 Step -1
            ListBox4.Items.Add(ListBox3.Items(i).ToString)
            ListBox3.Items.RemoveAt(i)
        Next i
    End Sub
    Private Sub Button6_Click(sender As Object, e As EventArgs) Handles Button6.Click

        Dim xlsxpath As String = ""
        'Dim strConn As String
        Dim SavingDialog As FileDialog = New SaveFileDialog
        With SavingDialog
            .AddExtension = True
            .DefaultExt = "xlsx"
            .Filter = "Excel 97-2003 Tables(*.xls)|*.xls|Excel Tables(*.xlsx)|*.xlsx"
            .Title = "Export to"
            If .ShowDialog = Windows.Forms.DialogResult.OK Then
                xlsxpath = .FileName
            End If
        End With
        If xlsxpath = "" Then
            Exit Sub
        End If

        Dim Excel As Excel.Application = New Excel.Application
        Dim Excelworkbook As Excel.Workbook
        Dim Excelsheet As Excel.Worksheet
        Excel.Visible = False
        Excelworkbook = Excel.Workbooks.Add()
        Excelsheet = Excelworkbook.Worksheets.Add()
        Excelsheet.Name = "Recent scores"
        Excelsheet.Activate()
        With Excelsheet
            .Range("A1").Value2 = "Records after"
            .Range("B1").Value2 = Microsoft.VisualBasic.Left(MonthCalendar2.SelectionStart.ToString, Len(MonthCalendar2.SelectionStart.ToString) - 9)
            .Range("A2").Value2 = "Student No"
            .Range("B2").Value2 = "Student Name"
            '.Columns.Range("A").ColumnWidth = 15
            '.Columns.Range("B").ColumnWidth = 12
        End With
        If RadioButton1.Checked = True Then
            Call sumrecord(Excelsheet)
        Else
            Call detailrecord(Excelsheet)
        End If
        Excelworkbook.SaveAs(xlsxpath)
        Excelworkbook.Close()
        MsgBox("Exported file saved in " & Chr(10) & xlsxpath, vbOKOnly, "Export Succeed")
        Excel.Visible = True
    End Sub
    Private Sub sumrecord(ByRef sheet As Excel.Worksheet)
        Dim i As Integer, temp As String, sum As Integer, j As Integer
        temp = Microsoft.VisualBasic.Left(MonthCalendar2.SelectionStart.ToString, Len(MonthCalendar2.SelectionStart.ToString) - 9)
        With sheet
            .Range("C2").Value2 = "Total score"
            DBconnection.Open()
            DBSQL = New OleDb.OleDbCommand("Select * From students", DBconnection)
            DBReader = DBSQL.ExecuteReader
            i = 3
            Do While DBReader.Read()
                .Range("A" & CStr(i)).Value2 = DBReader.Item("studentno")
                .Range("B" & CStr(i)).Value2 = DBReader.Item("studentname")
                i = i + 1
            Loop
            j = i - 1
            For i = 3 To j
                sum = 0
                DBSQL = New OleDb.OleDbCommand("Select * From [general] where studentno = " & .Range("A" & CStr(i)).Value2 & " and datediff('d',[Time]," & temp & ") <=0", DBconnection)
                DBReader = DBSQL.ExecuteReader
                Do While DBReader.Read()
                    sum = sum + Val(DBReader.Item("Number"))
                Loop
                .Range("C" & CStr(i)).Value2 = CStr(sum)
            Next
            DBconnection.Close()
        End With
    End Sub
    Private Sub detailrecord(ByRef sheet As Excel.Worksheet)
        Dim temp As String, i As Integer, j As Integer
        temp = Microsoft.VisualBasic.Left(MonthCalendar2.SelectionStart.ToString, Len(MonthCalendar2.SelectionStart.ToString) - 9)
        With sheet
            .Range("C2").Value2 = "Time"
            .Range("D2").Value2 = "Reason"
            .Range("E2").Value2 = "By whom"
            .Range("F2").Value2 = "Number"
            DBconnection.Open()
            DBSQL = New OleDb.OleDbCommand("Select * From [general] where datediff('d',[Time]," & temp & ") <=0", DBconnection)
            DBReader = DBSQL.ExecuteReader
            i = 3
            Do While DBReader.Read()
                .Range("A" & CStr(i)).Value2 = DBReader.Item("studentno")
                .Range("C" & CStr(i)).Value2 = DBReader.Item("Time")
                .Range("D" & CStr(i)).Value2 = DBReader.Item("Reason")
                .Range("E" & CStr(i)).Value2 = DBReader.Item("Person")
                .Range("F" & CStr(i)).Value2 = DBReader.Item("Number")
                i = i + 1
            Loop
            j = i - 1
            For i = 3 To j
                If .Range("A" & CStr(i)).Value2 <> 0 Then
                    DBSQL = New OleDb.OleDbCommand("Select * From students Where studentno =" & .Range("A" & CStr(i)).Value2, DBconnection)
                    DBReader = DBSQL.ExecuteReader
                    DBReader.Read()
                    .Range("B" & CStr(i)).Value2 = DBReader.Item("studentname")
                End If
            Next
            DBconnection.Close()
        End With
    End Sub
End Class