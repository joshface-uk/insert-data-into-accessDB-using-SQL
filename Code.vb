Imports System.Data.OleDb
Public Class Form1

    Dim provider As String
    Dim dataFile As String
    Dim connString As String
    Dim myConnection As OleDbConnection = New OleDbConnection

    Private Sub btn_Add_Click(sender As Object, e As EventArgs) Handles btn_Add.Click
        provider = "Provider=Microsoft.ACE.OLEDB.12.0; Data Source="
        dataFile = "C:\Users\Josh\Desktop\Users.accdb"
        connString = provider & dataFile
        myConnection.ConnectionString = connString
        myConnection.Open()
        Dim str As String
        str = "Insert into Users([ID],[Forename],[Surname],[Age],[Gender]) Values (?,?,?,?,?)"
        Dim cmd As OleDbCommand = New OleDbCommand(str, myConnection)


        cmd.Parameters.Add(New OleDbParameter("ID", CType(NUD_IDNumber.Value, String)))
        cmd.Parameters.Add(New OleDbParameter("Forename", CType(tbx_Forename.Text, String)))
        cmd.Parameters.Add(New OleDbParameter("Surname", CType(tbx_Surname.Text, String)))
        cmd.Parameters.Add(New OleDbParameter("Age", CType(NUD_Age.Value, String)))
        cmd.Parameters.Add(New OleDbParameter("Gender", CType(cbx_Gender.SelectedItem, String)))
        Try
            cmd.ExecuteNonQuery()
            cmd.Dispose()
            myConnection.Close()
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        cbx_Gender.Items.Add("Male")
        cbx_Gender.Items.Add("Female")
        cbx_Gender.Items.Add("Other")
    End Sub

    Private Sub btn_Clear_Click(sender As Object, e As EventArgs) Handles btn_Clear.Click
        NUD_IDNumber.Value = 0
        tbx_Forename.Clear()
        tbx_Surname.Clear()
        NUD_Age.Value = 0
        cbx_Gender.SelectedIndex = -1
    End Sub
End Class
