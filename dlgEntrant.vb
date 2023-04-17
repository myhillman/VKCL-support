Imports System.Windows.Forms
Imports Microsoft.Data.Sqlite

Public Class dlgEntrant

    Private Sub OK_Button_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles OK_Button.Click
        Me.DialogResult = System.Windows.Forms.DialogResult.OK
        Me.Close()
    End Sub

    Private Sub Cancel_Button_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Cancel_Button.Click
        Me.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.Close()
    End Sub

    Private Sub dlgEntrant_Load(sender As Object, e As EventArgs) Handles Me.Load
        Const CheckerDB = "data Source=Checker.db3"   ' name of the names file
        Dim sql As SqliteCommand, sqldr As SqliteDataReader, entrants As New ArrayList

        Using connect As New SqliteConnection(CheckerDB)
            connect.Open()
            sql = connect.CreateCommand
            sql.CommandText = $"SELECT * FROM Stations WHERE contestID={Me.Tag} ORDER BY station,section"     ' get list of entrants
            sqldr = sql.ExecuteReader
            While sqldr.Read
                If Not IsDBNull(sqldr("station")) And Not IsDBNull(sqldr("name")) Then
                    entrants.Add(New entrant(sqldr("station"), sqldr("section"), sqldr("name")))
                End If
            End While
            With DataGridView1
                .DataSource = entrants
            End With
        End Using
    End Sub
    Private Class entrant
        Property station As String
        Property section As String
        Property name As String
        Sub New(station As String, section As String, name As String)
            _station = station
            _section = section
            _name = name
        End Sub
    End Class
End Class
