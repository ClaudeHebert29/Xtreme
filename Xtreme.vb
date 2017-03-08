Imports System.Data.OleDb
Public Class Xtreme
    Dim bd As New GestionBD
    Dim daXtreme As New OleDbDataAdapter
    Dim dsXtreme As New DataSet
    Dim gestionoperation As New OleDbCommandBuilder
    Dim position, table, ctrTable As Integer
    Dim NomTable(), NomtableTout() As String
    Dim listeTXT_Client(), listeTXT_Four(), listeTXT_Produit(), listeTXT_Employes() As TextBox
    Dim listPanel() As Panel
    Private Sub Xtreme_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        NomtableTout = {"Achats", "Adresses des employés", "Clients", "Commandes", "Détails des commandes", "Employés", "Fournisseurs", "Info Xtreme", "Produits", "Régions", "Types de produit"}
        NomTable = {"Clients", "Fournisseurs", "Produits", "Employés"}
        listeTXT_Four = {txt_16, txt_17, Txt_18, Txt_19, Txt_20, Txt_21, Txt_22, Txt_23}
        listeTXT_Produit = {Txt_24, Txt_25, Txt_26, Txt_27, Txt_28, Txt_29, Txt_30}
        listeTXT_Employes = {Txt_31, Txt_32, Txt_33, Txt_34, Txt_35, Txt_36, Txt_37, Txt_39, Txt_40, Txt_41, Txt_42, Txt_43, Txt_44, Txt_45}
        listeTXT_Client = {txt_Clients_1, txt_Clients_2, txt_Clients_3, txt_Clients_4, txt_Clients_5, txt_Clients_6, txt_Clients_7, txt_Clients_8, txt_Clients_9, txt_Clients_10, txt_Clients_11, txt_Clients_12, txt_Clients_13, txt_Clients_14, txt_Clients_15}
        listPanel = {pan_clients, Pan_fournisseur, pan_produit, pan_produit}
        bd.connexion("..\xtreme.mdb")
        bd.Deconnexion()
        For Each s As String In NomTable
            cbx_Nomtable.Items.Add(s)
        Next
    End Sub
    Sub ChargerDataset()
        Dim cmdXtreme As New OleDbCommand
        dsXtreme = New DataSet
        cmdXtreme = bd.cnconnexion.CreateCommand
        cmdXtreme.CommandText = "Select * from " & NomTable(table)
        daXtreme.SelectCommand = cmdXtreme
        daXtreme.Fill(dsXtreme, NomTable(table))
    End Sub
    Sub RemplirControles()
        Dim ctr2 As Integer
        For ctr As Integer = 2 To dsXtreme.Tables(0).Columns.Count - 1
            If IsDBNull(dsXtreme.Tables(0).Rows(position).Item(ctr)) = False Then
                If table = 3 And ctr = 10 Then
                    ctr2 -= 1
                    pbx_Photo.BackgroundImage = dsXtreme.Tables(0).Rows(position).Item(ctr)
                Else
                    listeTXT_Client(ctr2).Text = dsXtreme.Tables(0).Rows(position).Item(ctr)

                End If
            Else
                listeTXT_Client(ctr2).Text = "Null"
            End If
            ctr2 += 1
        Next
    End Sub
    Private Sub btn_ElementLast_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btn_ElementLast.Click,
    btn_ElementNext.Click, btn_ElementPreview.Click, btn_ElementFirst.Click
        Select Case sender.text
            Case "<<"
                position = 0
                Btn_Element_Bloquer(False, False, True, True)
            Case "<"
                If position > 1 Then
                    position -= 1
                    Btn_Element_Bloquer(True, True, True, True)
                Else
                    position = 0
                    Btn_Element_Bloquer(False, False, True, True)
                End If
            Case ">"
                If position < dsXtreme.Tables(0).Rows.Count - 2 Then
                    position += 1
                    Btn_Element_Bloquer(True, True, True, True)
                Else
                    position += 1
                    Btn_Element_Bloquer(True, True, False, False)
                End If

            Case ">>"
                position = dsXtreme.Tables(0).Rows.Count - 1
                Btn_Element_Bloquer(True, True, False, False)
        End Select
        RemplirControles()
    End Sub

    Private Sub ChangerDeTable(ByVal sender As Object, ByVal e As System.EventArgs) Handles btn_ChangerTable.Click
        For ctr As Integer = 0 To 4
            If NomTable(ctr) Like cbx_Nomtable.Text Then
                changerPanel(ctr)
                table = ctr
                position = 0
                ChargerDataset()
                RemplirControles()
                Btn_Element_Bloquer(False, False, True, True)
                Exit For
            ElseIf ctr = 4 Then
                MsgBox("Erreur")
            End If
        Next
    End Sub
    Public Sub Btn_Element_Bloquer(a As Boolean, b As Boolean, c As Boolean, d As Boolean)
        btn_ElementFirst.Enabled = a
        btn_ElementPreview.Enabled = b
        btn_ElementLast.Enabled = c
        btn_ElementNext.Enabled = d
    End Sub
    Public Sub changerPanel(ctr As Integer)
        For c As Integer = 0 To 3
            If c = ctr Then
                listPanel(c).Visible = True
            Else
                listPanel(c).Visible = False
            End If
        Next
    End Sub
End Class
