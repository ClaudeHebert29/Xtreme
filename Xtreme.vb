Imports System.Data.OleDb
Public Class Xtreme
    Dim bd As New GestionBD
    Dim daXtreme, daTypeProduit, daFournisseur As New OleDbDataAdapter
    Dim dsXtreme, dsTypeProduit, dsFournisseur As New DataSet
    Dim gestionoperation As New OleDbCommandBuilder
    Dim position, table, ctrTable, min, max As Integer
    Dim NomTable(), NomtableTout() As String
    Dim listeTXT_Client(), listeTXT_Four(), listeTXT_Produit(), listeTXT_Employes(), listeTXT_Type_Produit() As TextBox
    Dim listeTXT As Object()
    Dim listPanel() As Panel
#Region "Load"
    Private Sub Xtreme_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        NomtableTout = {"Achats", "Adresses des employés", "Clients", "Commandes", "Détails des commandes", "Employés", "Fournisseurs", "Info Xtreme", "Produits", "Régions", "Types de produit"}
        NomTable = {"Clients", "Fournisseurs", "Produits", "Employés", "Types de produit"}
        listeTXT_Four = {txt_four_1, txt_four_2, Txt_four_3, Txt_four_4, Txt_four_5, Txt_four_7, txt_four_6, Txt_four_8}
        listeTXT_Produit = {Txt_prod_1, Txt_prod_2, Txt_prod_3, Txt_prod_4, Txt_prod_5, Txt_prod_6, Txt_prod_7, Txt_prod_8}
        listeTXT_Employes = {Txt_Emp_1, Txt_Emp_2, Txt_Emp_3, Txt_Emp_4, Txt_Emp_5, Txt_Emp_6, Txt_Emp_7, txt_Emp_10, Txt_Emp_8, Txt_Emp_9, Txt_Emp_11, Txt_Emp_12, Txt_Emp_16, Txt_Emp_13, Txt_Emp_14, Txt_Emp_15}
        listeTXT_Client = {txt_Clients_1, txt_Clients_2, txt_Clients_3, txt_Clients_4, txt_Clients_5, txt_Clients_6, txt_Clients_7, txt_Clients_8, txt_Clients_9, txt_Clients_10, txt_Clients_11, txt_Clients_12, txt_Clients_13, txt_Clients_14, txt_Clients_15}
        listeTXT = {listeTXT_Client, listeTXT_Four, listeTXT_Employes}
        listeTXT_Type_Produit = {Txt_Type_Prod_2, Txt_Type_Prod_3, Txt_Type_Prod_4}
        listeTXT = {listeTXT_Client, listeTXT_Four, listeTXT_Produit, listeTXT_Employes}
        listPanel = {pan_clients, pan_Fournisseur, Pan_produit, Pan_employer}
        bd.connexion("..\xtreme.mdb")
        bd.Deconnexion()
        Btn_Element_Bloquer(False, False, False, False)
        btnOption(False, False, False, False)
        For c As Integer = 0 To 3
            cbx_Nomtable.Items.Add(NomTable(c))
        Next
    End Sub
#End Region
#Region "BD"
    Sub ChargerDataset()
        Dim cmdXtreme As New OleDbCommand
        dsXtreme = New DataSet
        cmdXtreme = bd.cnconnexion.CreateCommand
        cmdXtreme.CommandText = "Select * from " & NomTable(table) ' & "where Actif = Oui"
        daXtreme.SelectCommand = cmdXtreme
        daXtreme.Fill(dsXtreme, NomTable(table))
        btnOption(True, True, True, False)
    End Sub
    Sub ChargerDataseTypeProduit()
        Dim cmdTypeProdui As New OleDbCommand
        dsTypeProduit = New DataSet
        cmdTypeProdui = bd.cnconnexion.CreateCommand
        cmdTypeProdui.CommandText = "Select * from " & "Types_de_produit"
        daTypeProduit.SelectCommand = cmdTypeProdui
        daTypeProduit.Fill(dsTypeProduit, "Types_de_produit")
    End Sub
    Sub ChargerDatasetFournisseur()
        Dim cmdFournisseur As New OleDbCommand
        dsFournisseur = New DataSet
        cmdFournisseur = bd.cnconnexion.CreateCommand
        cmdFournisseur.CommandText = "Select ID_fournisseur,Nom_du_fournisseur from Fournisseurs"
        daFournisseur.SelectCommand = cmdFournisseur
        daFournisseur.Fill(dsFournisseur, "Fournisseurs")
    End Sub

    Sub RemplirControles()
        Dim ctr2 As Integer
        For ctr As Integer = min To max
            If IsDBNull(dsXtreme.Tables(0).Rows(position).Item(ctr)) = False Then
                If table = 3 And ctr = 9 Then
                    listeTXT(table)(ctr2).text = "Null"
                Else
                    listeTXT(table)(ctr2).text = dsXtreme.Tables(0).Rows(position).Item(ctr)
                End If
            Else
                listeTXT(table)(ctr2).text = "Null"
            End If
            ctr2 += 1
        Next
        If table = 2 Then
            ChargerDataseTypeProduit()
            For c As Integer = 0 To dsTypeProduit.Tables(0).Rows.Count - 1
                If dsTypeProduit.Tables(0).Rows(c).Item(0) = listeTXT(table)(5).text Then
                    listeTXT(table)(5).text = dsTypeProduit.Tables(0).Rows(c).Item(1)
                    Exit For
                End If
            Next
            ChargerDatasetFournisseur()
            For c As Integer = 0 To dsFournisseur.Tables(0).Rows.Count - 1
                If dsFournisseur.Tables(0).Rows(c).Item(0) = listeTXT(table)(7).text Then
                    listeTXT(table)(7).text = dsFournisseur.Tables(0).Rows(c).Item(1)
                    Exit For
                End If
            Next

        End If

    End Sub
#End Region
#Region "Déplacement dans les tables"
    Private Sub btn_Deplacement(ByVal sender As Object, ByVal e As System.EventArgs) Handles btn_ElementLast.Click,
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
        TPVisiblePas(True)
        For ctr As Integer = 0 To 4
            If NomTable(ctr) Like cbx_Nomtable.Text And ctr < 4 Then
                changerPanel(ctr)
                table = ctr
                position = 0
                ChargerDataset()
                Btn_Element_Bloquer(False, False, True, True)
                Select Case cbx_Nomtable.Text
                    Case "Clients", "Employés"
                        min = 2
                    Case "Fournisseurs", "Produits"
                        min = 1
                End Select
                max = dsXtreme.Tables(0).Columns.Count - 2
                RemplirControles()
                Exit For
            ElseIf ctr = 4 Then
                MsgBox("Erreur")
                Exit For
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

#End Region
#Region "C.A.M.E."
#Region "Bouton ajouter"
    Private Sub btn_Ajouter_Click(sender As Object, e As EventArgs) Handles btn_Ajouter.Click
        Dim b As Boolean
        If sender.text = "Ajouter" Then
            EnableDurantoption(False)
            Select Case cbx_Nomtable.Text
                Case "Clients", "Employés"
                    vider(dsXtreme.Tables(0).Columns.Count - 4)
                Case "Fournisseurs", "Produits"
                    vider(dsXtreme.Tables(0).Columns.Count - 3)
            End Select
            sender.text = "Enregistrer"
            btnOption(True, False, False, True)
        Else
            EnableDurantoption(True)
            For c As Integer = 0 To dsXtreme.Tables(0).Columns.Count - 4
                If listeTXT(table)(c).text Like "" Then
                    b = False
                    Exit For
                Else
                    b = True
                End If
            Next
            If b = True Then
                Select Case cbx_Nomtable.Text
                    Case "Clients"
                        Ajouter(dsXtreme.Tables(0).Columns.Count - 4, 2)
                    Case "Fournisseurs", "Produits"
                        Ajouter(dsXtreme.Tables(0).Columns.Count - 3, 1)
                    Case "Employés"
                        Ajouter(dsXtreme.Tables(0).Columns.Count - 2, 2)
                End Select
                sender.text = "Ajouter"
                'miseAjourBD()
                btnOption(True, True, True, False)
            Else
                MsgBox("Des sections n'ont pas été remplies.")
            End If
        End If
    End Sub

    Sub Ajouter(nbr As Integer, min As Integer)
        Dim drnouvel As DataRow
        Dim c2 As Integer = min
        With dsXtreme.Tables(0)
            drnouvel = .NewRow()
            drnouvel(0) = dsXtreme.Tables(0).Rows.Count + 1
            If table = 0 Then
                drnouvel(1) = 0
            End If
            For c As Integer = 0 To nbr
                drnouvel(c2) = listeTXT(table)(c).text
                c2 += 1
            Next
            .Rows.Add(drnouvel)
        End With

    End Sub

#End Region
#Region "Bouton Modifier"

    Private Sub btnModifier_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btn_Modifier.Click
        EnableDurantoption(False)
        If sender.text = "Modifier" Then
            cbx_typeProduit.Items.Clear()

            If table = 2 Then
                For c As Integer = 0 To dsTypeProduit.Tables(0).Rows.Count - 1
                    cbx_typeProduit.Items.Add(dsTypeProduit.Tables(0).Rows(c).Item(1))
                Next
                For c As Integer = 0 To dsFournisseur.Tables(0).Rows.Count - 1
                    cbx_fournisseur.Items.Add(dsFournisseur.Tables(0).Rows(c).Item(1))
                Next
                listeTXT(table)(5).visible = False
                cbx_typeProduit.Visible = True
                cbx_typeProduit.Text = listeTXT(table)(5).text
                listeTXT(table)(7).visible = False
                cbx_fournisseur.Visible = True
                cbx_fournisseur.Text = listeTXT(table)(7).text
                btn_Ajouter_type_prod.Visible = True
            End If
            If table = 3 Then
                For c As Integer = 0 To dsXtreme.Tables(0).Rows.Count - 1
                    cbx_Sup.Items.Add(dsXtreme.Tables(0).Rows(c).Item(3))
                Next
                listeTXT(table)(9).visible = False
                cbx_Sup.Visible = True
                cbx_Sup.Text = listeTXT(table)(0).text
            End If
            btn_Modifier.Text = "Enregistrer"
            btnOption(False, True, False, True)
        Else
            If table = 2 Then
                listeTXT(table)(5).visible = True
                listeTXT(table)(7).visible = True
                listeTXT(table)(5).text = cbx_typeProduit.Text
                cbx_typeProduit.Visible = False
                cbx_fournisseur.Visible = False
                listeTXT(table)(7).text = cbx_fournisseur.Text
            End If
            btn_Modifier.Text = "Modifier"
            Select Case cbx_Nomtable.Text
                Case "Clients", "Employés"
                    modifier(dsXtreme.Tables(0).Columns.Count - 4, 2)
                Case "Fournisseurs", "Produits"
                    modifier(dsXtreme.Tables(0).Columns.Count - 3, 1)
            End Select
            EnableDurantoption(True)
            btnOption(True, True, True, False)
            'miseAjourBD()
        End If
    End Sub

    Sub modifier(nbr As Integer, min As Integer)
        For c As Integer = 0 To nbr
            If table = 3 And c = 7 Then
                dsXtreme.Tables(0).Rows(position).Item(min) = 0
            ElseIf table = 2 And c = 5 Then
                For c2 As Integer = 0 To dsTypeProduit.Tables(0).Rows.Count - 1
                    If dsTypeProduit.Tables(0).Rows(c2).Item(1) = listeTXT(table)(5).text Then
                        dsXtreme.Tables(0).Rows(position).Item(min) = dsTypeProduit.Tables(0).Rows(c2).Item(0)
                        Exit For
                    End If
                Next
            ElseIf table = 2 And c = 7 Then
                For c2 As Integer = 0 To dsFournisseur.Tables(0).Rows.Count - 1
                    If dsFournisseur.Tables(0).Rows(c2).Item(1) = listeTXT(table)(7).text Then
                        dsXtreme.Tables(0).Rows(position).Item(min) = dsFournisseur.Tables(0).Rows(c2).Item(0)
                        Exit For
                    End If
                Next
            ElseIf table = 3 And c = 9 Then
                For c2 As Integer = 0 To dsXtreme.Tables(0).Rows.Count - 1
                    If dsXtreme.Tables(0).Rows(c2).Item(2) = listeTXT(table)(9).text Then
                        dsXtreme.Tables(0).Rows(position).Item(min) = dsXtreme.Tables(0).Rows(c2).Item(0)
                        Exit For
                    End If
                Next
            Else
                dsXtreme.Tables(0).Rows(position).Item(min) = listeTXT(table)(c).text
            End If
            min += 1
        Next
    End Sub
#End Region
#Region "Bouton Supprimer"
    Private Sub btnSupprimer_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btn_supprimer.Click
        btnOption(False, False, False, True)
        gbx_sup.Visible = True
        EnableDurantoption(False)
    End Sub

    Private Sub btn_Oui_Click(sender As Object, e As EventArgs) Handles btn_Oui.Click, btn_Non.Click
        Select Case sender.text
            Case "Oui"
                dsXtreme.Tables(0).Rows(position).Item(dsXtreme.Tables(0).Columns.Count - 1) = False
                btnOption(True, True, True, False)
                'miseAjourBD()bite
            Case "Non"
                annuler()
        End Select
        gbx_sup.Visible = False
        EnableDurantoption(True)
    End Sub

#End Region
#Region "Bouton Annuler"
    Private Sub Annuler(ByVal sender As Object, ByVal e As System.EventArgs) Handles btn_annuler.Click
        annuler()
    End Sub

    Sub annuler()
        ChargerDataset()
        btnOption(True, True, True, False)
        RemplirControles()
        btn_supprimer.Text = "Supprimer"
        btn_Ajouter.Text = "Ajouter"
        btn_Modifier.Text = "Modifier"
        EnableDurantoption(True)
        gbx_sup.Visible = False
        If table = 2 Then
            listeTXT(table)(5).visible = True
            cbx_typeProduit.Visible = False
        End If
    End Sub
#End Region
#Region "Ajouter un type de produits"
    Private Sub Ajouter_Type_produit(sender As Object, e As EventArgs) Handles btn_Ajouter_type_prod.Click
        Pan_Type_produit.Visible = True
        Pan_produit.Visible = False
        btn_Ajouter_type_prod.Visible = False
        TPVisiblePas(False)
        For c As Integer = 0 To 2
            listeTXT_Type_Produit(c).Text = ""
        Next
    End Sub

    Private Sub btn_Enregistrer_Type_Click(sender As Object, e As EventArgs) Handles btn_Enregistrer_Type.Click
        Pan_Type_produit.Visible = False
        Pan_produit.Visible = True
        Dim drnouvel As DataRow
        Dim c2 As Integer = 1
        With dsTypeProduit.Tables(0)
            drnouvel = .NewRow()
            drnouvel(0) = dsTypeProduit.Tables(0).Rows.Count + 1
            If table = 0 Then
                drnouvel(1) = dsTypeProduit.Tables(0).Columns.Count + 1
            End If
            For c As Integer = 0 To 1
                drnouvel(c2) = listeTXT_Type_Produit(c).Text
                c2 += 1
            Next
            drnouvel(3) = 2
            .Rows.Add(drnouvel)
        End With
        cbx_typeProduit.Items.Add(listeTXT_Type_Produit(0).Text)
        cbx_typeProduit.Text = listeTXT_Type_Produit(0).Text
        TPVisiblePas(True)
        'miseAjourBD_TP()
    End Sub
    Sub miseAjourBD_TP()
        gestionoperation = New OleDbCommandBuilder(daTypeProduit)
        daTypeProduit.Update(dsTypeProduit, "Types_de_produit")
        dsTypeProduit.Clear()
        daTypeProduit.Fill(dsTypeProduit, "Types_de_produit")
    End Sub
    Private Sub btn_annuler_tp_Click(sender As Object, e As EventArgs) Handles btn_annuler_tp.Click
        TPVisiblePas(True)
        Pan_Type_produit.Visible = False
        Pan_produit.Visible = True
    End Sub

#End Region
#Region "Option"
    Sub miseAjourBD()
        gestionoperation = New OleDbCommandBuilder(daXtreme)
        daXtreme.Update(dsXtreme, NomTable(table))
        dsXtreme.Clear()
        daXtreme.Fill(dsXtreme, NomTable(table))
    End Sub
    Sub vider(nbr As Integer)
        For c As Integer = 0 To nbr
            listeTXT(table)(c).text = ""
        Next
    End Sub

    Sub btnOption(a As Boolean, b As Boolean, c As Boolean, d As Boolean)
        btn_Ajouter.Enabled = a
        btn_Modifier.Enabled = b
        btn_supprimer.Enabled = c
        btn_annuler.Enabled = d
    End Sub

    Sub TPVisiblePas(b As Boolean)
        btn_Ajouter.Visible = b
        btn_Modifier.Visible = b
        btn_supprimer.Visible = b
        btn_annuler.Visible = b
        btn_ElementFirst.Visible = b
        btn_ElementPreview.Visible = b
        btn_ElementLast.Visible = b
        btn_ElementNext.Visible = b
    End Sub
    Sub EnableDurantoption(b As Boolean)
        cbx_Nomtable.Visible = b
        btn_ChangerTable.Visible = b
        btn_ElementFirst.Visible = b
        btn_ElementPreview.Visible = b
        btn_ElementLast.Visible = b
        btn_ElementNext.Visible = b
    End Sub


    Private Sub cbx_Nomtable_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cbx_Nomtable.SelectedIndexChanged
        TPVisiblePas(False)
    End Sub
#End Region
#End Region
End Class

