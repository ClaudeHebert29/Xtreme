Imports System.Data.OleDb
Public Class ChargerTable
    Function ChargerDataset(ds As DataSet, da As OleDbDataAdapter, commande As String, Nom As String)
        Dim bd As New GestionBD
        Dim cmdXtreme As New OleDbCommand
        ds = New DataSet
        cmdXtreme = bd.cnconnexion.CreateCommand
        cmdXtreme.CommandText = commande
        da.SelectCommand = cmdXtreme
        da.Fill(ds, Nom)
        Return ds
    End Function
End Class
