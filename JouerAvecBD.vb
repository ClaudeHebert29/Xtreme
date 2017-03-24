
Imports System.Data.OleDb
Module JouerAvecBD
    Function ChargerDs(da As OleDbDataAdapter, commande As String, Nom As String)
        Dim bd As New GestionBD
        bd.connexion("..\xtreme.mdb")
        bd.Deconnexion()
        Dim ds As DataSet
        Dim cmdXtreme As New OleDbCommand
        ds = New DataSet
        cmdXtreme = bd.cnconnexion.CreateCommand
        cmdXtreme.CommandText = commande
        da.SelectCommand = cmdXtreme
        da.Fill(ds, Nom)
        Return ds
    End Function

    Function fisrt(pos As Integer, ds As DataSet)
        Do
            If ds.Tables(0).Rows(pos).Item(ds.Tables(0).Columns.Count() - 1) = False Then
                Exit Do
            End If
            pos = pos + 1
        Loop Until pos > ds.Tables(0).Rows.Count() - 1
        Return pos
    End Function

    Function last(pos As Integer, ds As DataSet)
        Do
            If ds.Tables(0).Rows(pos).Item(ds.Tables(0).Columns.Count() - 1) = False Then
                Exit Do
            End If
            pos = pos - 1
        Loop Until pos < 0
        Return pos
    End Function
    Function suivant(pos As Integer, fin As Integer, ds As DataSet)
        Do
            If pos >= ds.Tables(0).Rows.Count() Then
                Exit Do
            ElseIf ds.Tables(0).Rows(pos).Item(ds.Tables(0).Columns.Count() - 1) = False Then
                Exit Do
            Else
                pos = pos + 1
            End If
        Loop Until pos > fin
        Return pos
    End Function
    Function suivantCache(pos As Integer, fin As Integer, ds As DataSet)
        Dim b As Boolean
        Do
            If ds.Tables(0).Rows.Count() <= pos Then
                b = False
                Exit Do
            ElseIf ds.Tables(0).Rows(pos).Item(ds.Tables(0).Columns.Count() - 1) = False Then
                b = True
                Exit Do
            Else
                b = False
                pos = pos + 1
            End If
        Loop Until pos > fin
        Return b
    End Function
    Function preview(pos As Integer, fin As Integer, ds As DataSet)
        Do
            If pos <= 0 Then
                Exit Do
            ElseIf ds.Tables(0).Rows(pos).Item(ds.Tables(0).Columns.Count() - 1) = False Then
                Exit Do
            Else
                pos = pos - 1
            End If
        Loop Until pos > 0
        Return pos
    End Function
    Function previewCache(pos As Integer, fin As Integer, ds As DataSet)
        Dim b As Boolean
        Do
            If 0 >= pos Then
                b = False
                Exit Do
            ElseIf ds.Tables(0).Rows(pos).Item(ds.Tables(0).Columns.Count() - 1) = False Then
                b = True
                Exit Do
            Else
                b = False
                pos = pos - 1
            End If
        Loop Until pos > 0
        Return b
    End Function
End Module
