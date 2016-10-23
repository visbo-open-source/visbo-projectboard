Imports System
Imports System.Runtime.Serialization
Imports System.Xml
Imports System.Xml.Serialization
Imports System.IO


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

    Private Sub frmSettings_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        frmShowInfoBC.Checked = showBreadCrumbField
        frmExtendedSearch.Checked = extSearch

        frmUserName.Text = userName
        frmUserPWD.Text = ""

        rdbPWD.Checked = True
        lblProtectField1.Text = "Passwort:"

        lblProtectField2.Visible = False
        frmProtectField2.Visible = False
        frmProtectField2.Text = ""


    End Sub


    Private Sub dbLoginButton_Click(sender As Object, e As EventArgs) Handles btnDBLogin.Click

        userName = frmUserName.Text
        userPWD = frmUserPWD.Text

    End Sub

    Private Sub btnProtect_Click(sender As Object, e As EventArgs) Handles btnProtect.Click

        protectContents = Not protectContents

        For Each tmpShape As PowerPoint.Shape In currentSlide.Shapes
            If tmpShape.Tags.Count > 0 Then
                If isRelevantShape(tmpShape) Then
                    ' Sichtbarkeit setzen ....
                    tmpShape.Visible = protectContents
                End If
            End If
        Next

    End Sub

    Private Sub rdbPWD_CheckedChanged(sender As Object, e As EventArgs) Handles rdbPWD.CheckedChanged
        If rdbPWD.Checked = True Then
            lblProtectField1.Text = "Passwort:"

            lblProtectField2.Visible = False
            frmProtectField2.Visible = False
            frmProtectField2.Text = ""
        Else
            lblProtectField1.Text = "Domain-Name:"

            lblProtectField2.Visible = True

            frmProtectField2.Visible = True
            frmProtectField2.Text = ""
        End If
    End Sub

    Private Sub lbl_schrift_Click(sender As Object, e As EventArgs) Handles lbl_schrift.Click

    End Sub

    Private Sub frmSchriftGroesse_TextChanged(sender As Object, e As EventArgs) Handles txtboxSchriftGroesse.TextChanged
        Try
            schriftGroesse = CDbl(txtboxSchriftGroesse.Text)
        Catch ex As Exception
            Call MsgBox("unzuzlässiger Wert für Schriftgröße ...")
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
            Dim xmlFileName As String = My.Computer.FileSystem.SpecialDirectories.Desktop & "\" & "PPTlanguages.xlm"
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

            txtboxLanguage.SelectedItem = defaultSprache
            selectedLanguage = defaultSprache

            Dim sprachenArray As clsPrepLanguagesForXML = languages.getSprachenKlasse

            ' jetzt wird ein CustomXMLPart hinzugefügt 
            'Dim serializer = New DataContractSerializer(GetType(clsLanguages))
            Dim serializer = New DataContractSerializer(GetType(clsPrepLanguagesForXML))

            Dim file As New FileStream(xmlFileName, FileMode.Create)

            serializer.WriteObject(file, sprachenArray)
            file.Close()

            'Dim settings As New XmlWriterSettings()
            'settings.Indent = True
            'settings.IndentChars = (ControlChars.Tab)
            'settings.OmitXmlDeclaration = True


            'Dim writer As XmlWriter = XmlWriter.Create(xmlFileName, settings)
            'serializer.WriteObject(writer, languages)
            'serializer.WriteObject(writer, sprachenArray)

            'xmlResult = writer.ToString
            'writer.Flush()
            'writer.Close()

            'pptAPP.ActivePresentation.CustomXMLParts.Add(xmlResult)
        Catch ex As Exception
            Call MsgBox("Fehler bei Import: " & ex.Message)
        End Try

    End Sub
End Class