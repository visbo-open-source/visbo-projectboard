<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmSettings
    Inherits System.Windows.Forms.Form

    'Das Formular überschreibt den Löschvorgang, um die Komponentenliste zu bereinigen.
    <System.Diagnostics.DebuggerNonUserCode()> _
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        Try
            If disposing AndAlso components IsNot Nothing Then
                components.Dispose()
            End If
        Finally
            MyBase.Dispose(disposing)
        End Try
    End Sub

    'Wird vom Windows Form-Designer benötigt.
    Private components As System.ComponentModel.IContainer

    'Hinweis: Die folgende Prozedur ist für den Windows Form-Designer erforderlich.
    'Das Bearbeiten ist mit dem Windows Form-Designer möglich.  
    'Das Bearbeiten mit dem Code-Editor ist nicht möglich.
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Me.frmExtendedSearch = New System.Windows.Forms.CheckBox()
        Me.frmShowInfoBC = New System.Windows.Forms.CheckBox()
        Me.btnDBLogin = New System.Windows.Forms.Button()
        Me.frmUserPWD = New System.Windows.Forms.TextBox()
        Me.frmUserName = New System.Windows.Forms.TextBox()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.rdbPWD = New System.Windows.Forms.RadioButton()
        Me.rdbUserName = New System.Windows.Forms.RadioButton()
        Me.lblProtectField1 = New System.Windows.Forms.Label()
        Me.TabControl1 = New System.Windows.Forms.TabControl()
        Me.TabPage1 = New System.Windows.Forms.TabPage()
        Me.btnChangeLanguage = New System.Windows.Forms.Button()
        Me.txtboxLanguage = New System.Windows.Forms.ComboBox()
        Me.lblLanguage = New System.Windows.Forms.Label()
        Me.Label5 = New System.Windows.Forms.Label()
        Me.txtboxSchriftGroesse = New System.Windows.Forms.TextBox()
        Me.lbl_schrift = New System.Windows.Forms.Label()
        Me.txtboxAbstandsEinheit = New System.Windows.Forms.ComboBox()
        Me.TabPage2 = New System.Windows.Forms.TabPage()
        Me.btnProtect = New System.Windows.Forms.Button()
        Me.frmProtectField2 = New System.Windows.Forms.TextBox()
        Me.frmProtectField1 = New System.Windows.Forms.TextBox()
        Me.lblProtectField2 = New System.Windows.Forms.Label()
        Me.TabPage3 = New System.Windows.Forms.TabPage()
        Me.btnDBabbr = New System.Windows.Forms.Button()
        Me.TabPage4 = New System.Windows.Forms.TabPage()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.btnLanguageExp = New System.Windows.Forms.Button()
        Me.btnLanguageImp = New System.Windows.Forms.Button()
        Me.TabControl1.SuspendLayout()
        Me.TabPage1.SuspendLayout()
        Me.TabPage2.SuspendLayout()
        Me.TabPage3.SuspendLayout()
        Me.TabPage4.SuspendLayout()
        Me.SuspendLayout()
        '
        'frmExtendedSearch
        '
        Me.frmExtendedSearch.AutoSize = True
        Me.frmExtendedSearch.Location = New System.Drawing.Point(8, 108)
        Me.frmExtendedSearch.Name = "frmExtendedSearch"
        Me.frmExtendedSearch.Size = New System.Drawing.Size(106, 17)
        Me.frmExtendedSearch.TabIndex = 30
        Me.frmExtendedSearch.Text = "erweiterte Suche"
        Me.frmExtendedSearch.UseVisualStyleBackColor = True
        '
        'frmShowInfoBC
        '
        Me.frmShowInfoBC.AutoSize = True
        Me.frmShowInfoBC.Location = New System.Drawing.Point(8, 85)
        Me.frmShowInfoBC.Name = "frmShowInfoBC"
        Me.frmShowInfoBC.Size = New System.Drawing.Size(129, 17)
        Me.frmShowInfoBC.TabIndex = 29
        Me.frmShowInfoBC.Text = "Breadcrumb anzeigen"
        Me.frmShowInfoBC.UseVisualStyleBackColor = True
        '
        'btnDBLogin
        '
        Me.btnDBLogin.Location = New System.Drawing.Point(38, 81)
        Me.btnDBLogin.Name = "btnDBLogin"
        Me.btnDBLogin.Size = New System.Drawing.Size(91, 23)
        Me.btnDBLogin.TabIndex = 28
        Me.btnDBLogin.Text = "Login"
        Me.btnDBLogin.UseVisualStyleBackColor = True
        '
        'frmUserPWD
        '
        Me.frmUserPWD.Location = New System.Drawing.Point(92, 38)
        Me.frmUserPWD.Name = "frmUserPWD"
        Me.frmUserPWD.PasswordChar = Global.Microsoft.VisualBasic.ChrW(42)
        Me.frmUserPWD.Size = New System.Drawing.Size(187, 20)
        Me.frmUserPWD.TabIndex = 27
        Me.frmUserPWD.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        Me.frmUserPWD.UseSystemPasswordChar = True
        Me.frmUserPWD.WordWrap = False
        '
        'frmUserName
        '
        Me.frmUserName.Location = New System.Drawing.Point(92, 14)
        Me.frmUserName.Name = "frmUserName"
        Me.frmUserName.Size = New System.Drawing.Size(187, 20)
        Me.frmUserName.TabIndex = 26
        Me.frmUserName.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Location = New System.Drawing.Point(3, 41)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(56, 13)
        Me.Label4.TabIndex = 21
        Me.Label4.Text = "Password:"
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Location = New System.Drawing.Point(3, 17)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(83, 13)
        Me.Label3.TabIndex = 20
        Me.Label3.Text = "Benutzer-Name:"
        '
        'rdbPWD
        '
        Me.rdbPWD.AutoSize = True
        Me.rdbPWD.Checked = True
        Me.rdbPWD.Location = New System.Drawing.Point(6, 6)
        Me.rdbPWD.Name = "rdbPWD"
        Me.rdbPWD.Size = New System.Drawing.Size(68, 17)
        Me.rdbPWD.TabIndex = 32
        Me.rdbPWD.TabStop = True
        Me.rdbPWD.Text = "Passwort"
        Me.rdbPWD.UseVisualStyleBackColor = True
        '
        'rdbUserName
        '
        Me.rdbUserName.AutoSize = True
        Me.rdbUserName.Location = New System.Drawing.Point(96, 6)
        Me.rdbUserName.Name = "rdbUserName"
        Me.rdbUserName.Size = New System.Drawing.Size(117, 17)
        Me.rdbUserName.TabIndex = 33
        Me.rdbUserName.Text = "Domain-/Username"
        Me.rdbUserName.UseVisualStyleBackColor = True
        '
        'lblProtectField1
        '
        Me.lblProtectField1.AutoSize = True
        Me.lblProtectField1.Location = New System.Drawing.Point(6, 41)
        Me.lblProtectField1.Name = "lblProtectField1"
        Me.lblProtectField1.Size = New System.Drawing.Size(53, 13)
        Me.lblProtectField1.TabIndex = 34
        Me.lblProtectField1.Text = "Passwort:"
        '
        'TabControl1
        '
        Me.TabControl1.Controls.Add(Me.TabPage1)
        Me.TabControl1.Controls.Add(Me.TabPage2)
        Me.TabControl1.Controls.Add(Me.TabPage3)
        Me.TabControl1.Controls.Add(Me.TabPage4)
        Me.TabControl1.Location = New System.Drawing.Point(12, 12)
        Me.TabControl1.Name = "TabControl1"
        Me.TabControl1.SelectedIndex = 0
        Me.TabControl1.Size = New System.Drawing.Size(290, 167)
        Me.TabControl1.TabIndex = 35
        '
        'TabPage1
        '
        Me.TabPage1.Controls.Add(Me.btnChangeLanguage)
        Me.TabPage1.Controls.Add(Me.txtboxLanguage)
        Me.TabPage1.Controls.Add(Me.lblLanguage)
        Me.TabPage1.Controls.Add(Me.Label5)
        Me.TabPage1.Controls.Add(Me.txtboxSchriftGroesse)
        Me.TabPage1.Controls.Add(Me.lbl_schrift)
        Me.TabPage1.Controls.Add(Me.txtboxAbstandsEinheit)
        Me.TabPage1.Controls.Add(Me.frmExtendedSearch)
        Me.TabPage1.Controls.Add(Me.frmShowInfoBC)
        Me.TabPage1.Location = New System.Drawing.Point(4, 22)
        Me.TabPage1.Name = "TabPage1"
        Me.TabPage1.Padding = New System.Windows.Forms.Padding(3)
        Me.TabPage1.Size = New System.Drawing.Size(282, 141)
        Me.TabPage1.TabIndex = 0
        Me.TabPage1.Text = "Allgemein"
        Me.TabPage1.UseVisualStyleBackColor = True
        '
        'btnChangeLanguage
        '
        Me.btnChangeLanguage.Location = New System.Drawing.Point(183, 85)
        Me.btnChangeLanguage.Name = "btnChangeLanguage"
        Me.btnChangeLanguage.Size = New System.Drawing.Size(83, 40)
        Me.btnChangeLanguage.TabIndex = 39
        Me.btnChangeLanguage.Text = "Namen übersetzen"
        Me.btnChangeLanguage.UseVisualStyleBackColor = True
        '
        'txtboxLanguage
        '
        Me.txtboxLanguage.FormattingEnabled = True
        Me.txtboxLanguage.Location = New System.Drawing.Point(183, 51)
        Me.txtboxLanguage.Name = "txtboxLanguage"
        Me.txtboxLanguage.Size = New System.Drawing.Size(83, 21)
        Me.txtboxLanguage.TabIndex = 38
        Me.txtboxLanguage.Text = "deutsch"
        '
        'lblLanguage
        '
        Me.lblLanguage.AutoSize = True
        Me.lblLanguage.Location = New System.Drawing.Point(3, 54)
        Me.lblLanguage.Name = "lblLanguage"
        Me.lblLanguage.Size = New System.Drawing.Size(50, 13)
        Me.lblLanguage.TabIndex = 37
        Me.lblLanguage.Text = "Sprache:"
        '
        'Label5
        '
        Me.Label5.AutoSize = True
        Me.Label5.Location = New System.Drawing.Point(2, 31)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(106, 13)
        Me.Label5.TabIndex = 36
        Me.Label5.Text = "Abstand anzeigen in:"
        '
        'txtboxSchriftGroesse
        '
        Me.txtboxSchriftGroesse.Location = New System.Drawing.Point(183, 6)
        Me.txtboxSchriftGroesse.Name = "txtboxSchriftGroesse"
        Me.txtboxSchriftGroesse.Size = New System.Drawing.Size(83, 20)
        Me.txtboxSchriftGroesse.TabIndex = 19
        Me.txtboxSchriftGroesse.Text = "8"
        Me.txtboxSchriftGroesse.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'lbl_schrift
        '
        Me.lbl_schrift.AutoSize = True
        Me.lbl_schrift.Location = New System.Drawing.Point(2, 9)
        Me.lbl_schrift.Name = "lbl_schrift"
        Me.lbl_schrift.Size = New System.Drawing.Size(67, 13)
        Me.lbl_schrift.TabIndex = 18
        Me.lbl_schrift.Text = "Schriftgröße:"
        '
        'txtboxAbstandsEinheit
        '
        Me.txtboxAbstandsEinheit.FormattingEnabled = True
        Me.txtboxAbstandsEinheit.Items.AddRange(New Object() {"Tagen", "Wochen", "Monaten"})
        Me.txtboxAbstandsEinheit.Location = New System.Drawing.Point(183, 28)
        Me.txtboxAbstandsEinheit.Name = "txtboxAbstandsEinheit"
        Me.txtboxAbstandsEinheit.Size = New System.Drawing.Size(83, 21)
        Me.txtboxAbstandsEinheit.TabIndex = 25
        Me.txtboxAbstandsEinheit.Text = "Tagen"
        '
        'TabPage2
        '
        Me.TabPage2.Controls.Add(Me.btnProtect)
        Me.TabPage2.Controls.Add(Me.frmProtectField2)
        Me.TabPage2.Controls.Add(Me.frmProtectField1)
        Me.TabPage2.Controls.Add(Me.lblProtectField2)
        Me.TabPage2.Controls.Add(Me.rdbPWD)
        Me.TabPage2.Controls.Add(Me.lblProtectField1)
        Me.TabPage2.Controls.Add(Me.rdbUserName)
        Me.TabPage2.Location = New System.Drawing.Point(4, 22)
        Me.TabPage2.Name = "TabPage2"
        Me.TabPage2.Padding = New System.Windows.Forms.Padding(3)
        Me.TabPage2.Size = New System.Drawing.Size(282, 141)
        Me.TabPage2.TabIndex = 1
        Me.TabPage2.Text = "Schutz"
        Me.TabPage2.UseVisualStyleBackColor = True
        '
        'btnProtect
        '
        Me.btnProtect.Location = New System.Drawing.Point(96, 100)
        Me.btnProtect.Name = "btnProtect"
        Me.btnProtect.Size = New System.Drawing.Size(75, 23)
        Me.btnProtect.TabIndex = 38
        Me.btnProtect.Text = "Schützen"
        Me.btnProtect.UseVisualStyleBackColor = True
        '
        'frmProtectField2
        '
        Me.frmProtectField2.Location = New System.Drawing.Point(96, 62)
        Me.frmProtectField2.Name = "frmProtectField2"
        Me.frmProtectField2.Size = New System.Drawing.Size(180, 20)
        Me.frmProtectField2.TabIndex = 37
        '
        'frmProtectField1
        '
        Me.frmProtectField1.Location = New System.Drawing.Point(96, 38)
        Me.frmProtectField1.Name = "frmProtectField1"
        Me.frmProtectField1.PasswordChar = Global.Microsoft.VisualBasic.ChrW(42)
        Me.frmProtectField1.Size = New System.Drawing.Size(180, 20)
        Me.frmProtectField1.TabIndex = 36
        '
        'lblProtectField2
        '
        Me.lblProtectField2.AutoSize = True
        Me.lblProtectField2.Location = New System.Drawing.Point(6, 65)
        Me.lblProtectField2.Name = "lblProtectField2"
        Me.lblProtectField2.Size = New System.Drawing.Size(55, 13)
        Me.lblProtectField2.TabIndex = 35
        Me.lblProtectField2.Text = "Username"
        '
        'TabPage3
        '
        Me.TabPage3.Controls.Add(Me.btnDBabbr)
        Me.TabPage3.Controls.Add(Me.Label3)
        Me.TabPage3.Controls.Add(Me.Label4)
        Me.TabPage3.Controls.Add(Me.btnDBLogin)
        Me.TabPage3.Controls.Add(Me.frmUserName)
        Me.TabPage3.Controls.Add(Me.frmUserPWD)
        Me.TabPage3.Location = New System.Drawing.Point(4, 22)
        Me.TabPage3.Name = "TabPage3"
        Me.TabPage3.Size = New System.Drawing.Size(282, 141)
        Me.TabPage3.TabIndex = 2
        Me.TabPage3.Text = "Datenbank"
        Me.TabPage3.UseVisualStyleBackColor = True
        '
        'btnDBabbr
        '
        Me.btnDBabbr.Location = New System.Drawing.Point(153, 81)
        Me.btnDBabbr.Name = "btnDBabbr"
        Me.btnDBabbr.Size = New System.Drawing.Size(91, 23)
        Me.btnDBabbr.TabIndex = 29
        Me.btnDBabbr.Text = "Abbrechen"
        Me.btnDBabbr.UseVisualStyleBackColor = True
        '
        'TabPage4
        '
        Me.TabPage4.Controls.Add(Me.Label1)
        Me.TabPage4.Controls.Add(Me.btnLanguageExp)
        Me.TabPage4.Controls.Add(Me.btnLanguageImp)
        Me.TabPage4.Location = New System.Drawing.Point(4, 22)
        Me.TabPage4.Name = "TabPage4"
        Me.TabPage4.Size = New System.Drawing.Size(282, 141)
        Me.TabPage4.TabIndex = 3
        Me.TabPage4.Text = "Sprachen I/E"
        Me.TabPage4.UseVisualStyleBackColor = True
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(37, 28)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(81, 13)
        Me.Label1.TabIndex = 2
        Me.Label1.Text = "Sprachen-Datei"
        '
        'btnLanguageExp
        '
        Me.btnLanguageExp.Location = New System.Drawing.Point(153, 59)
        Me.btnLanguageExp.Name = "btnLanguageExp"
        Me.btnLanguageExp.Size = New System.Drawing.Size(75, 23)
        Me.btnLanguageExp.TabIndex = 1
        Me.btnLanguageExp.Text = "Exportieren"
        Me.btnLanguageExp.UseVisualStyleBackColor = True
        '
        'btnLanguageImp
        '
        Me.btnLanguageImp.Location = New System.Drawing.Point(37, 59)
        Me.btnLanguageImp.Name = "btnLanguageImp"
        Me.btnLanguageImp.Size = New System.Drawing.Size(75, 23)
        Me.btnLanguageImp.TabIndex = 0
        Me.btnLanguageImp.Text = "Importieren"
        Me.btnLanguageImp.UseVisualStyleBackColor = True
        '
        'frmSettings
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(323, 202)
        Me.Controls.Add(Me.TabControl1)
        Me.Name = "frmSettings"
        Me.Text = "Einstellungen"
        Me.TabControl1.ResumeLayout(False)
        Me.TabPage1.ResumeLayout(False)
        Me.TabPage1.PerformLayout()
        Me.TabPage2.ResumeLayout(False)
        Me.TabPage2.PerformLayout()
        Me.TabPage3.ResumeLayout(False)
        Me.TabPage3.PerformLayout()
        Me.TabPage4.ResumeLayout(False)
        Me.TabPage4.PerformLayout()
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents frmExtendedSearch As System.Windows.Forms.CheckBox
    Friend WithEvents frmShowInfoBC As System.Windows.Forms.CheckBox
    Friend WithEvents btnDBLogin As System.Windows.Forms.Button
    Friend WithEvents frmUserPWD As System.Windows.Forms.TextBox
    Friend WithEvents frmUserName As System.Windows.Forms.TextBox
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents rdbPWD As System.Windows.Forms.RadioButton
    Friend WithEvents rdbUserName As System.Windows.Forms.RadioButton
    Friend WithEvents lblProtectField1 As System.Windows.Forms.Label
    Friend WithEvents TabControl1 As System.Windows.Forms.TabControl
    Friend WithEvents TabPage1 As System.Windows.Forms.TabPage
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents txtboxSchriftGroesse As System.Windows.Forms.TextBox
    Friend WithEvents lbl_schrift As System.Windows.Forms.Label
    Friend WithEvents txtboxAbstandsEinheit As System.Windows.Forms.ComboBox
    Friend WithEvents TabPage2 As System.Windows.Forms.TabPage
    Friend WithEvents TabPage3 As System.Windows.Forms.TabPage
    Friend WithEvents btnDBabbr As System.Windows.Forms.Button
    Friend WithEvents btnProtect As System.Windows.Forms.Button
    Friend WithEvents frmProtectField2 As System.Windows.Forms.TextBox
    Friend WithEvents frmProtectField1 As System.Windows.Forms.TextBox
    Friend WithEvents lblProtectField2 As System.Windows.Forms.Label
    Friend WithEvents txtboxLanguage As System.Windows.Forms.ComboBox
    Friend WithEvents lblLanguage As System.Windows.Forms.Label
    Friend WithEvents TabPage4 As System.Windows.Forms.TabPage
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents btnLanguageExp As System.Windows.Forms.Button
    Friend WithEvents btnLanguageImp As System.Windows.Forms.Button
    Friend WithEvents btnChangeLanguage As System.Windows.Forms.Button
End Class
