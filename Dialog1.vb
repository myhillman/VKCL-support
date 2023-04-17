Imports System.Windows.Forms
Imports Microsoft.Data.Sqlite

Public Class dlgContest

    Private Sub OK_Button_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles OK_Button.Click
        Me.DialogResult = System.Windows.Forms.DialogResult.OK
        Me.Close()
    End Sub

    Private Sub Cancel_Button_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Cancel_Button.Click
        Me.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.Close()
    End Sub

    Private Sub dlgContest_Load(sender As Object, e As EventArgs) Handles Me.Load
        Const CheckerDB = "data Source=Checker.db3"   ' name of the names file
        Dim sql As SqliteCommand, sqldr As SqliteDataReader, contests As New ArrayList

        Using connect As New SqliteConnection(CheckerDB)
            connect.Open()
            sql = connect.CreateCommand
            sql.CommandText = "SELECT * FROM Contests ORDER BY Start"     ' get list of contests
            sqldr = sql.ExecuteReader
            While sqldr.Read
                contests.Add(New contest(sqldr("Name"), sqldr("contestID"), sqldr("start"), sqldr("finish"), sqldr("PermittedBands"), sqldr("PermittedModes")))
            End While
            With DataGridView1
                .DataSource = contests
                .Columns("id").Visible = False
            End With
        End Using
    End Sub
    Private Class contest
        Property id As Integer
        Property name As String
        Property start As String
        Property finish As String
        Property PermittedBands As String
        Property PermittedModes As String
        Sub New(name As String, id As Integer, start As String, finish As String, PermittedBands As String, PermittedModes As String)
            _name = name
            _id = id
            _start = start
            _finish = finish
            _PermittedBands = PermittedBands
            _PermittedModes = PermittedModes
        End Sub
    End Class
End Class

