Imports System.Data.OleDb
Public Class Xtreme
    Dim bd As New GestionBD
    Dim daXtreme As New OleDbDataAdapter
    Dim dsXtreme As New DataSet
    Dim gestionoperation As New OleDbCommandBuilder
    Dim position, table, ctrTable As Integer
    Dim NomTable(), NomtableTout() As String
    Dim listeTXT_Client(), listeTXT_Four(), listeTXT_Produit(), listeTXT_Employes() As TextBox
    Dim listeTXT As Object()
    Dim listPanel() As Panel
    Dim min, max As Integer

    Private Sub Pan_fournisseur_Paint(sender As Object, e As PaintEventArgs) Handles Pan_fournisseur.Paint

    End Sub

    Private Sub Xtreme_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        NomtableTout = {"Achats", "Adresses des employés", "Clients", "Commandes", "Détails des commandes", "Employés", "Fournisseurs", "Info Xtreme", "Produits", "Régions", "Types de produit"}
        NomTable = {"Clients", "Fournisseurs", "Produits", "Employés"}
        listeTXT_Four = {txt_four_1, txt_four_2, Txt_four_3, Txt_four_4, Txt_four_5, Txt_four_7, txt_four_6, Txt_four_8, Txt_prod_1}
        listeTXT_Produit = {Txt_prod_2, Txt_prod_3, Txt_prod_4, Txt_prod_5, Txt_prod_6, Txt_prod_7, Txt_prod_8}
        listeTXT_Employes = {Txt_Emp_1, Txt_Emp_2, Txt_Emp_3, Txt_Emp_4, Txt_Emp_5, Txt_Emp_6, Txt_Emp_7, txt_Emp_10, Txt_Emp_8, Txt_Emp_9, Txt_Emp_11, Txt_Emp_12, Txt_Emp_13, Txt_Emp_14, Txt_Emp_15}
        listeTXT_Client = {txt_Clients_1, txt_Clients_2, txt_Clients_3, txt_Clients_4, txt_Clients_5, txt_Clients_6, txt_Clients_7, txt_Clients_8, txt_Clients_9, txt_Clients_10, txt_Clients_11, txt_Clients_12, txt_Clients_13, txt_Clients_14, txt_Clients_15}
        listeTXT = {listeTXT_Client, listeTXT_Four, listeTXT_Produit, listeTXT_Employes}
        listPanel = {pan_clients, pan_produit, Pan_fournisseur, Pan_employer}
        bd.connexion("..\xtreme.mdb")
        bd.Deconnexion()
        Btn_Element_Bloquer(False, False, False, False)
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
        For ctr As Integer = min To max
            If IsDBNull(dsXtreme.Tables(0).Rows(position).Item(ctr)) = False Then
                If table = 3 And ctr = 9 Then
                    listeTXT_Client(ctr2).Text = "Null"
                Else
                    listeTXT(table)(ctr2).text = dsXtreme.Tables(0).Rows(position).Item(ctr)
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
                Btn_Element_Bloquer(False, False, True, True)
                Select Case cbx_Nomtable.Text
                    Case "Clients"
                        min = 2
                        max = dsXtreme.Tables(0).Columns.Count - 1
                    Case "Fournisseurs"
                        min = 1
                        max = dsXtreme.Tables(0).Columns.Count - 1
                    Case "Produits"
                        min = 1
                        max = dsXtreme.Tables(0).Columns.Count - 2
                    Case "Employés"
                        min = 1
                        max = dsXtreme.Tables(0).Columns.Count - 3
                End Select
                RemplirControles()
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
