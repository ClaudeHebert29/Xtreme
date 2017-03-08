Imports System.Data.OleDb
Public Class GestionBD
    Dim _cnConnexion As OleDbConnection

    Sub connexion(ByVal bd As String)

        Dim sConnexion As String
        sConnexion = "Provider=Microsoft.jet.oledb.4.0;"
        sConnexion &= "Password=;User Id=Admin;"
        sConnexion &= "Data Source=" & bd
        _cnConnexion = New OleDbConnection(sConnexion)

        Try
            _cnConnexion.Open()
        Catch ex As OleDbException
            MsgBox(ex.Message)
        End Try
    End Sub

    Sub Deconnexion()
        Try
            _cnConnexion.Close()
        Catch ex As OleDbException
            MsgBox(ex.Message)
        End Try
    End Sub
    ReadOnly Property cnconnexion() As OleDbConnection
        Get
            cnconnexion = _cnConnexion
        End Get

    End Property

End Class
