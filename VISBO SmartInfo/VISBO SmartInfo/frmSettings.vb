Imports System
Imports System.Runtime.Serialization
Imports System.Xml
Imports System.Xml.Serialization
Imports System.IO
Imports ProjectBoardDefinitions
Imports MongoDbAccess


Public Class frmSettings


    Private Sub schriftSize_TextChanged(sender As Object, e As EventArgs)

        Try
            schriftGroesse = CDbl(txtboxSchriftGroesse.Text)
        Catch ex As Exception
            txtboxSchriftGroesse.Text = schriftGroesse.ToString
        End Try
    End Sub

    Private Sub abstandseinheit_SelectedIndexChanged(sender As Object, e As EventArgs) Handles txtboxAbstandsEinheit.SelectedIndexChanged

        If txtboxAbstandsEinheit.Text = "Tagen" Then
            absEinheit = pptAbsUnit.tage
        ElseIf txtboxAbstandsEinheit.Text = "Wochen" Then
            absEinheit = pptAbsUnit.wochen
        Else
            absEinheit = pptAbsUnit.monate
        End If

    End Sub

    Private Sub showInfoBC_CheckedChanged(sender As Object, e As EventArgs) Handles frmShowInfoBC.CheckedChanged

        showBreadCrumbField = frmShowInfoBC.Checked

    End Sub

    Private Sub extendedSearch_CheckedChanged(sender As Object, e As EventArgs) Handles frmExtendedSearch.CheckedChanged
        extSearch = frmExtendedSearch.Checked
    End Sub


    Private Sub languageSettings()
        If englishLanguage Then

            With Me
                .TabControl1.TabPages.Item(0).Text = "Basics"
                .TabControl1.TabPages.Item(1).Text = "Protection"
                .TabControl1.TabPages.Item(2).Text = "Language I/E"
                .lbl_schrift.Text = "Font Size"
                .Label5.Text = "distance unit"
                .lblLanguage.Text = "Language"
                .frmShowInfoBC.Text = "show full BreadCrumb"
                .frmExtendedSearch.Text = "extended search"
                .btnChangeLanguage.Text = "Translate Annotations"
                .rdbPWD.Text = "Password"
                .lblProtectField1.Text = "Password:"
                .btnProtect.Text = "Protect"
                .Label1.Text = "Language File"
                .btnLanguageImp.Text = "Import"
                .btnLanguageExp.Text = "Export"
            End With

        End If

    End Sub
    Private Sub frmSettings_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        frmShowInfoBC.Checked = showBreadCrumbField
        frmExtendedSearch.Checked = extSearch

        Call languageSettings()

        rdbPWD.Checked = True
        lblProtectField1.Text = "Password:"



        lblProtectField2.Visible = False
        frmProtectField2.Visible = False
        frmProtectField2.Text = ""


        If languages.count > 1 Then
            ' jetzt wird die txtboxLanguage aktualisiert
            txtboxLanguage.Visible = True
            lblLanguage.Visible = True
            txtboxLanguage.Items.Clear()

            For i As Integer = 1 To languages.count
                Dim sprache As String = languages.getLanguageName(i)
                txtboxLanguage.Items.Add(sprache)
            Next

            txtboxLanguage.SelectedItem = selectedLanguage

        Else
            txtboxLanguage.Visible = False
            lblLanguage.Visible = False
            selectedLanguage = defaultSprache
        End If


    End Sub


    

    Private Sub btnProtect_Click(sender As Object, e As EventArgs) Handles btnProtect.Click

        VisboProtected = True

        If rdbPWD.Checked Then
            pptAPP.ActivePresentation.Tags.Add(protectionTag, "PWD")
            pptAPP.ActivePresentation.Tags.Add(protectionValue, frmProtectField1.Text)
        ElseIf rdbDBLogin.Checked Then
            pptAPP.ActivePresentation.Tags.Add(protectionTag, "DATABASE")
            pptAPP.ActivePresentation.Tags.Add(protectionValue, "")
        Else
            pptAPP.ActivePresentation.Tags.Add(protectionTag, "COMPUTER")
            pptAPP.ActivePresentation.Tags.Add(protectionValue, frmProtectField1.Text & "\" & frmProtectField2.Text)
        End If

        Call makeVisboShapesVisible(False)

    End Sub

    Private Sub rdbPWD_CheckedChanged(sender As Object, e As EventArgs) Handles rdbPWD.CheckedChanged
        If rdbPWD.Checked = True Then

            lblProtectField1.Text = "Passwort:"
            lblProtectField1.Visible = True
            frmProtectField1.Visible = True

            lblProtectField2.Visible = False
            frmProtectField2.Visible = False
            frmProtectField2.Text = ""
        
        End If
    End Sub

    Private Sub rdbDBLogin_CheckedChanged(sender As Object, e As EventArgs) Handles rdbDBLogin.CheckedChanged

        If rdbDBLogin.Checked = True Then
            lblProtectField1.Text = ""
            lblProtectField1.Visible = False
            frmProtectField1.Visible = False

            lblProtectField2.Visible = False
            frmProtectField2.Visible = False
            frmProtectField2.Text = ""
        End If

    End Sub

    Private Sub rdbUserName_CheckedChanged(sender As Object, e As EventArgs) Handles rdbUserName.CheckedChanged

        If rdbUserName.Checked = True Then
            lblProtectField1.Text = ""
            lblProtectField1.Visible = True
            frmProtectField1.Visible = True

            lblProtectField2.Visible = True
            frmProtectField2.Visible = True
            frmProtectField2.Text = ""


            frmProtectField1.PasswordChar = ""
        End If
    End Sub

    Private Sub lbl_schrift_Click(sender As Object, e As EventArgs) Handles lbl_schrift.Click

    End Sub

    Private Sub frmSchriftGroesse_TextChanged(sender As Object, e As EventArgs) Handles txtboxSchriftGroesse.TextChanged
        Try
            schriftGroesse = CDbl(txtboxSchriftGroesse.Text)
        Catch ex As Exception
            Call MsgBox("unzulässiger Wert für Schriftgröße ...")
            txtboxSchriftGroesse.Text = schriftGroesse.ToString
        End Try

    End Sub

    Private Sub btnLanguageExp_Click(sender As Object, e As EventArgs) Handles btnLanguageExp.Click
        Try
            Dim tmpCollection = smartSlideLists.getElementNamen
            Call languages.addLanguage(defaultSprache, tmpCollection)
            Call languages.exportLanguages()
            Call MsgBox("ok, nach Desktop exportiert ...")
        Catch ex As Exception
            Call MsgBox("Fehler bei Export: " & ex.Message)
        End Try

    End Sub

    ''' <summary>
    ''' Bedingungen beim Import: 
    ''' es muss die Default Sprache geben, es muss jeweils eine einein-deutige Übersetzung existieren ....
    ''' also die Anzahl der Default-Sprachen Elemente muss gleich sein der Anzahl der anderen Sprachen-Elemente
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub btnLanguageImp_Click(sender As Object, e As EventArgs) Handles btnLanguageImp.Click

        Try
            Dim tmpLanguages As New clsLanguages
            Dim tmpCollection = smartSlideLists.getElementNamen
            Dim xmlFileName As String = My.Computer.FileSystem.SpecialDirectories.Desktop & "\" & "PPTlanguages.xml"
            Dim xmlResult As String = ""

            Call languages.addLanguage(defaultSprache, tmpCollection)
            Call languages.importLanguages()
            Dim anzahlLanguages As Integer = languages.count
            Call MsgBox("ok, " & anzahlLanguages - 1 & " weitere Sprachen importiert ...")

            ' jetzt wird die txtboxLanguage aktualisiert
            txtboxLanguage.Items.Clear()
            For i As Integer = 1 To languages.count
                Dim sprache As String = languages.getLanguageName(i)
                txtboxLanguage.Items.Add(sprache)
            Next

            txtboxLanguage.Visible = True
            lblLanguage.Visible = True

            txtboxLanguage.SelectedItem = defaultSprache
            selectedLanguage = defaultSprache

            ' --- nicht mehr benötigt
            ''Dim serializer = New DataContractSerializer(GetType(clsLanguages))
            ''Dim file As New FileStream(xmlFileName, FileMode.Create)
            ''serializer.WriteObject(file, languages)
            ''file.Close()


            '
            ' --- alte customXMLPart für Languages, falls vorhanden,  löschen
            '
            Dim oldlangGUID As String = pptAPP.ActivePresentation.Tags.Item("langGUID")
            If oldlangGUID.Length > 0 Then
                pptAPP.ActivePresentation.CustomXMLParts.SelectByID(oldlangGUID).Delete()
            End If

            '
            ' --- neues CustomXMLPart für Languages hinzufügen
            '
            ' der folgende Befehl embedded eine XML Struktur - in einem String -  in die aktive PPT Datei 
            ' Beschreibung zum Konzept der customXMLParts siehe: 
            ' siehe https://msdn.microsoft.com/en-us/library/bb608612.aspx 
            '
            Dim langXMLstring As String = xml_serialize(languages)
            Dim languageXMLPart As Office.CustomXMLPart = pptAPP.ActivePresentation.CustomXMLParts.Add(langXMLstring)

            '
            ' Setzen einen Tags zum Merken der GUID des CustomXMLPart - Language
            '
            pptAPP.ActivePresentation.Tags.Add("langGUID", languageXMLPart.Id)

            Dim anzXMLParts As Integer = pptAPP.ActivePresentation.CustomXMLParts.Count


        Catch ex As Exception
            Call MsgBox("Fehler bei Import: " & ex.Message)
        End Try

    End Sub

    Private Sub txtboxLanguage_SelectedIndexChanged(sender As Object, e As EventArgs) Handles txtboxLanguage.SelectedIndexChanged
        selectedLanguage = CStr(txtboxLanguage.SelectedItem)
    End Sub

    Private Sub btnChangeLanguage_Click(sender As Object, e As EventArgs) Handles btnChangeLanguage.Click
        Call changeLanguageInAnnotations()
    End Sub


    Private Sub TabPage4_Click(sender As Object, e As EventArgs) Handles TabPage4.Click
        Call MsgBox("jetzt in TabPage4")
    End Sub


End Class