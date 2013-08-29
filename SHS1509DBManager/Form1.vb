Public Class Form1
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        If TextBox1.Text = "shs1509" And TextBox2.Text = "1509shs" Then 'Password Check
            If TextBox3.Text = "" Then
                MsgBox("Please type the path of the DB file.", vbOKOnly, "DB Error")
                Exit Sub
            ElseIf System.IO.File.Exists(TextBox3.Text) = False Then
                MsgBox("Database file doesn't exist.", vbOKOnly, "DB Error")
            Else
                DBpath = TextBox3.Text
                DBconnection = JetDBconnect(DBpath) 'DB Connect
                Try
                    DBconnection.Open()
                    DBconnection.Close()
                Catch ex As Exception
                    MsgBox("Connection failed! Unknown error.", vbOKOnly, "Failed")
                End Try
                MsgBox("Connection Succeeded!", vbOKOnly, "Succeed")
                Form2.Show()
                Me.Close()
            End If
        Else
            MsgBox("Incorrect username or password!" & Chr(10) & "Login failed", vbOKOnly, "ERROR!")
            End
        End If
    End Sub
    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Dim browsedialog As FileDialog = New OpenFileDialog
        With browsedialog
            .AddExtension = True
            .Filter = "Access DB File(*.mdb)|*.mdb"
            .CheckFileExists = True
            .Title = "Open DB File"
            .DefaultExt = "mdb"
            .ShowDialog()
            TextBox3.Text = browsedialog.FileName
        End With
    End Sub
End Class
