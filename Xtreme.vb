Imports System.Data.OleDb
Public Class Xtreme
    Dim bd As New GestionBD
    Dim daXtreme, daTypeProduit, daFournisseur, daAdressesDesEmployes As New OleDbDataAdapter
    Dim dsXtreme, dsTypeProduit, dsFournisseur, dsAdressesDesEmployes As New DataSet
    Dim gestionoperation As New OleDbCommandBuilder
    Dim position, table, ctrTable, min, max, posAdresse, pos As Integer
    Dim NomTable(), NomtableTout(), nomColonne() As String
    Dim listeTXT_Client(), listeTXT_Four(), listeTXT_Produit(), listeTXT_Employes(), listeTXT_Type_Produit(), listTxt_adresse(), listeTXT_Commande() As TextBox
    Dim listeTXT As Object()
    Dim listPanel() As Panel
#Region "Load"
    Private Sub Xtreme_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        PictureBox1.BackgroundImage = Image.FromFile("travail.png")
        NomtableTout = {"Achats", "Adresses des Employes", "Clients", "Commandes", "Détails des commandes", "Employes", "Fournisseurs", "Info Xtreme", "Produits", "Régions", "Types de produit"}
        NomTable = {"Clients", "Fournisseurs", "Produits", "Employes", "Types de produit"}
        nomColonne = {"Nom_du_client", "Nom_du_fournisseur", "Nom_du_produit", "Nom"}
        listeTXT_Four = {txt_four_1, txt_four_2, Txt_four_3, Txt_four_4, Txt_four_5, Txt_four_7, txt_four_6, Txt_four_8}
        listeTXT_Produit = {Txt_prod_1, Txt_prod_2, Txt_prod_3, Txt_prod_4, Txt_prod_5, Txt_prod_6, Txt_prod_7, Txt_prod_8}
        listeTXT_Employes = {Txt_Emp_1, Txt_Emp_2, Txt_Emp_3, Txt_Emp_4, Txt_Emp_5, Txt_Emp_6, Txt_Emp_7, txt_Emp_10, Txt_Emp_8, Txt_Emp_9, Txt_Emp_11, Txt_Emp_12, Txt_Emp_16, Txt_Emp_13, Txt_Emp_14, Txt_Emp_15, Txt_Emp_17, Txt_Emp_18}
        listeTXT_Client = {txt_Clients_1, txt_Clients_2, txt_Clients_3, txt_Clients_4, txt_Clients_5, txt_Clients_6, txt_Clients_7, txt_Clients_8, txt_Clients_9, txt_Clients_10, txt_Clients_11, txt_Clients_12, txt_Clients_13, txt_Clients_14, txt_Clients_15}
        listeTXT_Commande = {Txt_ID_Commande_RGeneral_NC, Txt_montant_Commande_RGeneral_NC, Txt_Nom_Employer_RGeneral_NC, Txt_Date_de_commande_RGeneral_NC, TextBox8, TextBox9, TextBox10, TextBox11, TextBox12}
        listeTXT_Type_Produit = {Txt_Type_Prod_2, Txt_Type_Prod_3, Txt_Type_Prod_4}
        listeTXT = {listeTXT_Client, listeTXT_Four, listeTXT_Produit, listeTXT_Employes, listeTXT_Commande}
        listTxt_adresse = {txt_Adresse1, txt_Adresse2, txt_Adresse3, txt_Adresse4, txt_Adresse5, txt_Adresse6, txt_Adresse7, txt_Adresse8, txt_Adresse9, txt_Adresse10, txt_Adresse11, txt_Adresse12}
        listPanel = {pan_clients, pan_Fournisseur, Pan_produit, Pan_employer}
        bd.connexion("..\xtreme.mdb")
        bd.Deconnexion()
        Btn_Element_Bloquer(False, False, False, False)
        btnOption(False, False, False, False)
        TPVisiblePas(False)
        For c As Integer = 0 To 3
            cbx_Nomtable.Items.Add(NomTable(c))
        Next
        dtp_Naissance.MaxDate = Date.Today
        dtp_Embauche.MaxDate = Date.Today
    End Sub
#End Region
#Region "Gestion des table"
#Region "BD"
    Sub ChargerDataset()
        Dim cmdXtreme As New OleDbCommand
        dsXtreme = New DataSet
        cmdXtreme = bd.cnconnexion.CreateCommand
        cmdXtreme.CommandText = "Select * from " + NomTable(table) ' + " where innactif = false"
        daXtreme.SelectCommand = cmdXtreme
        daXtreme.Fill(dsXtreme, NomTable(table))
        btnOption(True, True, True, False)
    End Sub
    Sub ChargerDataseTypeProduit()
        Dim cmdTypeProdui As New OleDbCommand
        dsTypeProduit = New DataSet
        cmdTypeProdui = bd.cnconnexion.CreateCommand
        cmdTypeProdui.CommandText = "Select * from " & "TypesDeProduit"
        daTypeProduit.SelectCommand = cmdTypeProdui
        daTypeProduit.Fill(dsTypeProduit, "TypesDeProduit")
    End Sub
    Sub ChargerDatasetAdressesDesEmployes()
        Dim cmdAdressesDesEmployes As New OleDbCommand
        dsAdressesDesEmployes = New DataSet
        cmdAdressesDesEmployes = bd.cnconnexion.CreateCommand
        cmdAdressesDesEmployes.CommandText = "Select * from AdressesDesEmployes"
        daAdressesDesEmployes.SelectCommand = cmdAdressesDesEmployes
        daAdressesDesEmployes.Fill(dsAdressesDesEmployes, "AdressesDesEmployes")
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
        PosEcrireListBox(True)
        Dim ctr2 As Integer
        For ctr As Integer = min To max
            If IsDBNull(dsXtreme.Tables(0).Rows(position).Item(ctr)) = False Then
                If table = 3 And ctr = 9 Then
                    listeTXT(table)(ctr2).text = "-"
                Else
                    listeTXT(table)(ctr2).text = dsXtreme.Tables(0).Rows(position).Item(ctr)
                End If
            Else
                listeTXT(table)(ctr2).text = "-"
            End If
            ctr2 += 1
        Next
        If IsDBNull(dsXtreme.Tables(0).Rows(position).Item(1)) = False Then
            listeTXT(table)(17).text = dsXtreme.Tables(0).Rows(position).Item(1)
        Else
            listeTXT(table)(17).text = "-"
        End If

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
        ElseIf table = 3 Then
            For c As Integer = 0 To dsXtreme.Tables(0).Rows.Count - 1
                If listeTXT(table)(9).text = "-" Then
                    listeTXT(table)(9).text = "-"
                    Exit For
                ElseIf dsXtreme.Tables(0).Rows(c).Item(0) = listeTXT(table)(9).text Then
                    listeTXT(table)(9).text = dsXtreme.Tables(0).Rows(c).Item(2)
                    Exit For
                End If
            Next
            For c As Integer = 0 To dsXtreme.Tables(0).Rows.Count - 1
                If listeTXT(table)(17).text = "-" Then
                    listeTXT(table)(17).text = "-"
                    Exit For
                ElseIf dsXtreme.Tables(0).Rows(c).Item(0) = listeTXT(table)(17).text Then
                    listeTXT(table)(17).text = dsXtreme.Tables(0).Rows(c).Item(2)
                    Exit For
                End If
            Next
            ChargerDatasetAdressesDesEmployes()
            For c As Integer = 0 To dsAdressesDesEmployes.Tables(0).Rows.Count - 1
                If dsAdressesDesEmployes.Tables(0).Rows(c).Item(0) = dsXtreme.Tables(0).Rows(position).Item(0) Then
                    listeTXT(table)(16).text = dsAdressesDesEmployes.Tables(0).Rows(c).Item(1)
                    posAdresse = c
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
                pos = 0
                Do
                    If dsXtreme.Tables(0).Rows(pos).Item(dsXtreme.Tables(0).Columns.Count() - 1) = False Then
                        position = pos
                        Btn_Element_Bloquer(False, False, True, True)
                        Exit Do
                    End If
                    pos = pos + 1
                Loop Until pos > dsXtreme.Tables(0).Rows.Count() - 1
            Case "<"
                pos = position - 1
                Do
                    If dsXtreme.Tables(0).Rows(pos).Item(dsXtreme.Tables(0).Columns.Count() - 1) = False Then
                        Btn_Element_Bloquer(True, True, True, True)
                        position = pos
                        If pos < 1 Then
                            Btn_Element_Bloquer(False, False, True, True)
                        End If
                        Exit Do
                    End If
                    pos = pos - 1
                Loop Until pos < 0
                pos = position - 1
                Do
                    If 0 > pos Then
                        Btn_Element_Bloquer(False, False, True, True)
                        Exit Do
                    ElseIf 0 = pos And dsXtreme.Tables(0).Rows(pos).Item(dsXtreme.Tables(0).Columns.Count() - 1) = True Then
                        Btn_Element_Bloquer(False, False, True, True)
                        Exit Do
                    End If
                    If dsXtreme.Tables(0).Rows(pos).Item(dsXtreme.Tables(0).Columns.Count() - 1) = False Then
                        Exit Do
                    End If
                    pos = pos - 1
                Loop Until pos = 0
            Case ">"
                pos = position + 1
                Do
                    If dsXtreme.Tables(0).Rows(pos).Item(dsXtreme.Tables(0).Columns.Count() - 1) = False Then
                        Btn_Element_Bloquer(True, True, True, True)
                        position = pos
                        Exit Do
                    End If

                    pos = pos + 1
                Loop Until pos > dsXtreme.Tables(0).Rows.Count() - 1
                pos = position + 1
                Do
                    If dsXtreme.Tables(0).Rows.Count() = pos Then
                        Btn_Element_Bloquer(True, True, False, False)
                        Exit Do
                    ElseIf dsXtreme.Tables(0).Rows.Count() - 1 = pos And dsXtreme.Tables(0).rows(pos).Item(dsXtreme.Tables(0).Columns.Count() - 1) = True Then
                        Btn_Element_Bloquer(True, True, False, False)
                        Exit Do
                    End If
                    If dsXtreme.Tables(0).Rows(pos).Item(dsXtreme.Tables(0).Columns.Count() - 1) = False Then
                        Exit Do
                    End If
                    pos = pos + 1
                Loop Until pos = dsXtreme.Tables(0).Rows.Count()
            Case ">>"
                pos = dsXtreme.Tables(0).Rows.Count() - 1
                Do
                    If dsXtreme.Tables(0).Rows(pos).Item(dsXtreme.Tables(0).Columns.Count() - 1) = False Then
                        position = pos
                        Btn_Element_Bloquer(True, True, False, False)
                        Exit Do
                    End If
                    pos = pos - 1
                Loop Until pos < 0
        End Select
        RemplirControles()
    End Sub

    Private Sub ChangerDeTable(ByVal sender As Object, ByVal e As System.EventArgs) Handles btn_ChangerTable.Click
        TPVisiblePas(True)
        For ctr As Integer = 0 To 4
            If NomTable(ctr) Like cbx_Nomtable.Text And ctr < 4 Then
                changerPanel(ctr)
                table = ctr
                ChargerDataset()
                Btn_Element_Bloquer(False, False, True, True)
                Select Case cbx_Nomtable.Text
                    Case "Clients", "Employes"
                        min = 2
                    Case "Fournisseurs", "Produits"
                        min = 1
                End Select
                max = dsXtreme.Tables(0).Columns.Count - 2
                pos = 0
                Do
                    If dsXtreme.Tables(0).Rows(pos).Item(dsXtreme.Tables(0).Columns.Count() - 1) = False Then
                        position = pos
                        Btn_Element_Bloquer(False, False, True, True)
                        Exit Do
                    End If
                    pos = pos + 1
                Loop Until pos > dsXtreme.Tables(0).Rows.Count() - 1
                RemplirControles()
                Exit For
            ElseIf ctr = 4 Then
                MsgBox("Voyez selectionner une table.")
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
        If sender.text = "Ajouter" Then
            PosEcrireListBox(False)
            If table = 2 Then
                Remplir_cbx_table_2()
                option_Cbx_table_2(False, True, True)
            End If
            If table = 3 Then
                cacher_Adresse_Employer(False, True)
                Remplir_cbx_table_3()
                option_Cbx_table_3(False, True, True)
                cbx_Sup_h.Text = "-"
                dtp_Naissance.Text = Date.Today
                dtp_Embauche.Text = Date.Today
            End If
            EnableDurantoption(False)
            Select Case cbx_Nomtable.Text
                Case "Clients", "Employes"
                    vider(dsXtreme.Tables(0).Columns.Count - 4)
                Case "Fournisseurs", "Produits"
                    vider(dsXtreme.Tables(0).Columns.Count - 3)
            End Select
            sender.text = "Enregistrer"
            If table = 3 Then
                sender.text = "Suivant"
            End If
            btnOption(True, False, False, True)
        Else
            PosEcrireListBox(True)
            If table = 2 Then
                option_Cbx_table_2(True, False, False)
            End If
            EnableDurantoption(True)
            Select Case cbx_Nomtable.Text
                Case "Clients", "Employes"
                    Ajouter(dsXtreme.Tables(0).Columns.Count - 4, 2)
                Case "Fournisseurs", "Produits"
                    Ajouter(dsXtreme.Tables(0).Columns.Count - 3, 1)
            End Select
            sender.text = "Ajouter"
            miseAjourBD()
            If table = 3 Then
                pan_adresse.Visible = True
                Pan_employer.Visible = False
                btn_retour_Employer.Text = "Ajouter"
                position += 1
                btnOption(False, False, False, True)
                EnableDurantoption(False)
            Else
                btnOption(True, True, True, False)
            End If
        End If
    End Sub
    Sub cacher_Adresse_Employer(b As Boolean, bEtape As Boolean)
        btn_Adresse.Visible = b
        listeTXT(table)(16).Visible = b
        Label45.Visible = b
        lab_Etape1.Visible = bEtape
        lab_etape2.Visible = bEtape
    End Sub
    Sub Ajouter(nbr As Integer, min As Integer)
        Dim drnouvel As DataRow
        Dim c2 As Integer = min
        With dsXtreme.Tables(0)
            drnouvel = .NewRow()
            If table = 2 Then
                drnouvel(0) = dsXtreme.Tables(0).Rows.Count + 3
            Else
                drnouvel(0) = dsXtreme.Tables(0).Rows.Count + 3
            End If

            If table = 0 Then
                drnouvel(1) = DBNull.Value
            End If
            For c3 As Integer = 0 To nbr
                If listeTXT(table)(c3).text = Nothing Or listeTXT(table)(c3).text = "-" Then
                    drnouvel(c2) = DBNull.Value
                ElseIf table = 2 And c3 = 5 Then
                    For c As Integer = 0 To dsTypeProduit.Tables(0).Rows.Count - 1
                        If dsTypeProduit.Tables(0).Rows(c).Item(1) = cbx_typeProduit.Text Then
                            drnouvel(c2) = dsTypeProduit.Tables(0).Rows(c).Item(0)
                            Exit For
                        End If
                    Next
                ElseIf table = 2 And c3 = 7 Then
                    For c As Integer = 0 To dsFournisseur.Tables(0).Rows.Count - 1
                        If dsFournisseur.Tables(0).Rows(c).Item(1) = cbx_fournisseur.Text Then
                            drnouvel(c2) = dsFournisseur.Tables(0).Rows(c).Item(0)
                            Exit For
                        End If
                    Next
                ElseIf table = 3 And c3 = 9 Then
                    For c As Integer = 0 To dsXtreme.Tables(0).Rows.Count - 1
                        If cbx_Sup_h.Text = "-" Then
                            drnouvel(c2) = DBNull.Value
                            Exit For
                        ElseIf dsXtreme.Tables(0).Rows(c).Item(2) = cbx_Sup_h.Text Then
                            drnouvel(c2) = dsXtreme.Tables(0).Rows(c).Item(0)
                            Exit For
                        End If
                    Next
                    For c As Integer = 0 To dsXtreme.Tables(0).Rows.Count - 1
                        If cbx_Sup.Text = "-" Then
                            drnouvel(1) = DBNull.Value
                            Exit For
                        ElseIf dsXtreme.Tables(0).Rows(c).Item(2) = cbx_Sup.Text Then
                            drnouvel(1) = dsXtreme.Tables(0).Rows(c).Item(0)
                            Exit For
                        End If
                    Next
                ElseIf table = 3 And c3 = 7 Then
                    drnouvel(c2) = 0
                Else
                    drnouvel(c2) = listeTXT(table)(c3).text
                End If

                c2 += 1
            Next
            .Rows.Add(drnouvel)
        End With

    End Sub
    Sub miseAjourBD()
        gestionoperation = New OleDbCommandBuilder(daXtreme)
        gestionoperation.QuotePrefix = "["
        gestionoperation.QuoteSuffix = "]"
        daXtreme.Update(dsXtreme, NomTable(table))
        dsXtreme.Clear()
        daXtreme.Fill(dsXtreme, NomTable(table))
    End Sub

#End Region
#Region "Bouton Modifier"
    Private Sub btnModifier_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btn_Modifier.Click
        EnableDurantoption(False)
        If sender.text = "Modifier" Then
            PosEcrireListBox(False)
            If table = 2 Then
                Remplir_cbx_table_2()
                option_Cbx_table_2(False, True, True)
            End If
            If table = 3 Then
                Remplir_cbx_table_3()
                option_Cbx_table_3(False, True, True)
                btn_retour_Employer.Text = "Enregistrer"
            End If
            btn_Modifier.Text = "Enregistrer"
            btn_Adresse.Text = "Modifier"
            btnOption(False, True, False, True)
        Else
            PosEcrireListBox(True)
            If table = 2 Then
                option_Cbx_table_2(True, False, False)
            End If
            If table = 3 Then
                option_Cbx_table_3(True, False, False)
                btn_retour_Employer.Text = "Retour"
                btn_Adresse.Text = "Visualiser"
            End If
            btn_Modifier.Text = "Modifier"
            Select Case cbx_Nomtable.Text
                Case "Clients", "Employes"
                    modifier(dsXtreme.Tables(0).Columns.Count - 4, 2)
                Case "Fournisseurs", "Produits"
                    modifier(dsXtreme.Tables(0).Columns.Count - 3, 1)
            End Select
            EnableDurantoption(True)
            btnOption(True, True, True, False)
            miseAjourBD()
        End If
    End Sub

    Sub modifier(nbr As Integer, min As Integer)
        For c As Integer = 0 To nbr
            If listeTXT(table)(c).text = Nothing Or listeTXT(table)(c).text = "-" Then
                dsXtreme.Tables(0).Rows(position).Item(min) = DBNull.Value
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
            ElseIf table = 3 And c = 7 Then
                dsXtreme.Tables(0).Rows(position).Item(min) = 0
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
                dsXtreme.Tables(0).Rows(position).Item(dsXtreme.Tables(0).Columns.Count - 1) = True
                position = 0
                btnOption(True, True, True, False)
                Btn_Element_Bloquer(False, False, True, True)
                miseAjourBD()
                ChargerDataset()
            Case "Non"
                annuler()
        End Select
        gbx_sup.Visible = False
        EnableDurantoption(True)
        RemplirControles()
    End Sub

#End Region
#Region "Bouton Annuler"
    Private Sub Annuler(ByVal sender As Object, ByVal e As System.EventArgs) Handles btn_annuler.Click
        annuler()
    End Sub

    Sub annuler()
        cacher_Adresse_Employer(True, False)
        PosEcrireListBox(True)
        ChargerDataset()
        btnOption(True, True, True, False)
        RemplirControles()
        btn_supprimer.Text = "Supprimer"
        btn_Ajouter.Text = "Ajouter"
        btn_Modifier.Text = "Modifier"
        EnableDurantoption(True)
        gbx_sup.Visible = False
        If table = 2 Then
            option_Cbx_table_2(True, False, False)
        End If
        If table = 3 Then
            Vis_table_3(True, False)
        End If
        pan_adresse.Visible = False
        Pan_employer.Visible = True
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
            drnouvel(0) = dsTypeProduit.Tables(0).Rows.Count
            If listeTXT_Type_Produit(0).Text = Nothing Then
                drnouvel(1) = DBNull.Value
            Else
                drnouvel(1) = listeTXT_Type_Produit(0).Text
            End If
            If listeTXT_Type_Produit(0).Text = Nothing Then
                drnouvel(2) = DBNull.Value
            Else
                drnouvel(2) = listeTXT_Type_Produit(1).Text
            End If
            drnouvel(3) = DBNull.Value
            .Rows.Add(drnouvel)
        End With
        cbx_typeProduit.Items.Add(listeTXT_Type_Produit(0).Text)
        cbx_typeProduit.Text = listeTXT_Type_Produit(0).Text
        TPVisiblePas(True)
        miseAjourBD_TP()
    End Sub
    Sub miseAjourBD_TP()
        gestionoperation = New OleDbCommandBuilder(daTypeProduit)
        gestionoperation.QuotePrefix = "["
        gestionoperation.QuoteSuffix = "]"
        daTypeProduit.Update(dsTypeProduit, "TypesDeProduit")
        dsTypeProduit.Clear()
        daTypeProduit.Fill(dsTypeProduit, "TypesDeProduit")
    End Sub
    Private Sub btn_annuler_tp_Click(sender As Object, e As EventArgs) Handles btn_annuler_tp.Click
        TPVisiblePas(True)
        Pan_Type_produit.Visible = False
        Pan_produit.Visible = True
    End Sub

#End Region
#Region "Adresse Employé"
    Private Sub VoirPanAdresse(sender As Object, e As EventArgs) Handles btn_Adresse.Click
        Dim ctr2 As Integer
        If sender.text = "Voir" Then
            pan_adresse.Visible = True
            Pan_employer.Visible = False
        Else
            pan_adresse.Visible = True
            Pan_employer.Visible = False
            For c As Integer = 1 To dsAdressesDesEmployes.Tables(0).Columns.Count - 1
                If IsDBNull(dsAdressesDesEmployes.Tables(0).Rows(posAdresse).Item(c)) = True Then
                    listTxt_adresse(ctr2).Text = "-"
                Else
                    listTxt_adresse(ctr2).Text = CStr(dsAdressesDesEmployes.Tables(0).Rows(posAdresse).Item(c))
                End If
                ctr2 += 1
            Next
        End If
    End Sub
    Private Sub btn_retour_Employer_Click(sender As Object, e As EventArgs) Handles btn_retour_Employer.Click
        Select Case sender.text
            Case "Retour"
                pan_adresse.Visible = False
                Pan_employer.Visible = True
                btn_Adresse.Text = "Visualiser"
            Case "Enregistrer"
                pan_adresse.Visible = False
                Pan_employer.Visible = True
                btn_Adresse.Text = "Voir"
                Dim c2 As Integer = 1
                For c As Integer = 0 To dsAdressesDesEmployes.Tables(0).Columns.Count - 2
                    If listTxt_adresse(c).Text = Nothing Or listTxt_adresse(c).Text = "-" Then
                        dsAdressesDesEmployes.Tables(0).Rows(position).Item(c2) = DBNull.Value
                    Else
                        dsAdressesDesEmployes.Tables(0).Rows(position).Item(c2) = listTxt_adresse(c).Text
                    End If
                    c2 += 1
                Next
            Case "Ajouter"
                pan_adresse.Visible = False
                Pan_employer.Visible = True
                AjoutAdresse()
                btn_retour_Employer.Text = "Retour"
                miseAjourBDAdresseEmploye()
                btnOption(True, True, True, False)
                option_Cbx_table_3(True, False, False)
                cacher_Adresse_Employer(True, False)
                EnableDurantoption(False)
        End Select
    End Sub
    Sub miseAjourBDAdresseEmploye()
        gestionoperation = New OleDbCommandBuilder(daAdressesDesEmployes)
        gestionoperation.QuotePrefix = "["
        gestionoperation.QuoteSuffix = "]"
        daAdressesDesEmployes.Update(dsAdressesDesEmployes, "AdressesDesEmployes")
        dsAdressesDesEmployes.Clear()
        daAdressesDesEmployes.Fill(dsAdressesDesEmployes, "AdressesDesEmployes")
    End Sub
    Sub AjoutAdresse()
        Dim drnouvelAdresse As DataRow
        Dim c2 As Integer = 0
        With dsAdressesDesEmployes.Tables(0)
            drnouvelAdresse = .NewRow()
            drnouvelAdresse(0) = dsXtreme.Tables(0).Rows(dsXtreme.Tables(0).Rows.Count - 1).Item(0)
            For c As Integer = 1 To dsAdressesDesEmployes.Tables(0).Columns.Count - 1
                If listTxt_adresse(c2).Text = Nothing Or listTxt_adresse(c2).Text = "-" Then
                    drnouvelAdresse(c) = DBNull.Value
                Else
                    drnouvelAdresse(c) = listTxt_adresse(c2).Text
                End If
                c2 += 1
            Next
            .Rows.Add(drnouvelAdresse)
        End With
    End Sub
#End Region
#Region "Option"

    Sub vider(nbr As Integer)
        For c As Integer = 0 To nbr
            listeTXT(table)(c).text = ""
        Next
        If table = 3 Then
            listeTXT(table)(16).text = ""
            listeTXT(table)(17).text = ""
            cbx_Sup_h.Text = ""
            cbx_Sup.Text = ""
        End If
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
    Sub Remplir_cbx_table_2()
        cbx_typeProduit.Items.Clear()
        For c As Integer = 0 To dsTypeProduit.Tables(0).Rows.Count - 1
            cbx_typeProduit.Items.Add(dsTypeProduit.Tables(0).Rows(c).Item(1))
        Next
        For c As Integer = 0 To dsFournisseur.Tables(0).Rows.Count - 1
            cbx_fournisseur.Items.Add(dsFournisseur.Tables(0).Rows(c).Item(1))
        Next
    End Sub
    Sub Remplir_cbx_table_3()
        cbx_Sup_h.Items.Clear()
        cbx_Sup.Items.Clear()
        For c As Integer = 0 To dsXtreme.Tables(0).Rows.Count - 1
            cbx_Sup_h.Items.Add(dsXtreme.Tables(0).Rows(c).Item(3))
            cbx_Sup.Items.Add(dsXtreme.Tables(0).Rows(c).Item(3))
        Next
        cbx_Sup_h.Items.Add("-")
        cbx_Sup.Items.Add("-")
    End Sub
    Sub option_Cbx_table_2(b_liste As Boolean, b_cbx As Boolean, b As Boolean)
        listeTXT(table)(5).visible = b_liste
        listeTXT(table)(7).visible = b_liste
        cbx_typeProduit.Visible = b_cbx
        cbx_fournisseur.Visible = b_cbx
        btn_Ajouter_type_prod.Visible = b_cbx
        lab_TypeProd.Visible = b_cbx
        If b = True Then
            cbx_typeProduit.Text = listeTXT(table)(5).text
            cbx_fournisseur.Text = listeTXT(table)(7).text
        Else
            listeTXT(table)(5).text = cbx_typeProduit.Text
            listeTXT(table)(7).text = cbx_fournisseur.Text
        End If
    End Sub
    Sub option_Cbx_table_3(b_liste As Boolean, b_cbx As Boolean, b As Boolean)
        Vis_table_3(b_liste, b_cbx)
        If b = True Then
            dtp_Naissance.Text = listeTXT(table)(3).text
            dtp_Embauche.Text = listeTXT(table)(4).text
            cbx_Sup_h.Text = listeTXT(table)(9).text
            cbx_Sup.Text = listeTXT(table)(17).text
        Else
            listeTXT(table)(3).text = dtp_Naissance.Text
            listeTXT(table)(4).text = dtp_Embauche.Text
            listeTXT(table)(9).text = cbx_Sup_h.Text
            listeTXT(table)(17).text = cbx_Sup.Text
        End If
    End Sub
    Sub Vis_table_3(b_liste As Boolean, b_cbx As Boolean)
        listeTXT(table)(3).visible = b_liste
        listeTXT(table)(4).visible = b_liste
        listeTXT(table)(9).visible = b_liste
        listeTXT(table)(17).visible = b_liste
        cbx_Sup_h.Visible = b_cbx
        cbx_Sup.Visible = b_cbx
        dtp_Embauche.Visible = b_cbx
        dtp_Naissance.Visible = b_cbx
    End Sub
#End Region
#End Region
#Region "Menu Principal"
    Private Sub ChangerdepageDuMenu(sender As Object, e As EventArgs) Handles btn_Gestion_table.Click, btn_Recherche_Generales.Click, btn_Recherche_produit.Click, btn_Commandes.Click, btn_menu1.Click, btn_menu2.Click, btn_menu3.Click, Rech_prod.Click
        Select Case sender.tag
            Case "1"
                Rech_prod.SelectedIndex = 1
            Case "2"
                Rech_prod.SelectedIndex = 2
            Case "3"
                Rech_prod.SelectedIndex = 3
            Case "4"
                Rech_prod.SelectedIndex = 4
            Case "Menu"
                Rech_prod.SelectedIndex = 0
        End Select
    End Sub
    Sub PosEcrireListBox(b As Boolean)
        For Each c As Object In listeTXT(table)
            c.ReadOnly = b
        Next
    End Sub
#End Region


#End Region
#Region "Numéro de commande"
    Private Sub Btn_ok_NC_RGeneral_Click(sender As Object, e As EventArgs) Handles Btn_ok_NC_RGeneral.Click

    End Sub
#End Region

End Class

