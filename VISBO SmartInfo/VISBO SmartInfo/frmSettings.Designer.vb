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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmSettings))
        Me.TabPage4 = New System.Windows.Forms.TabPage()
        Me.btnLanguageImp = New System.Windows.Forms.Button()
        Me.btnLanguageExp = New System.Windows.Forms.Button()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.TabPage1 = New System.Windows.Forms.TabPage()
        Me.txtboxAbstandsEinheit = New System.Windows.Forms.ComboBox()
        Me.lbl_schrift = New System.Windows.Forms.Label()
        Me.txtboxSchriftGroesse = New System.Windows.Forms.TextBox()
        Me.Label5 = New System.Windows.Forms.Label()
        Me.lblLanguage = New System.Windows.Forms.Label()
        Me.txtboxLanguage = New System.Windows.Forms.ComboBox()
        Me.btnChangeLanguage = New System.Windows.Forms.Button()
        Me.TabControl1 = New System.Windows.Forms.TabControl()
        Me.TabPage4.SuspendLayout()
        Me.TabPage1.SuspendLayout()
        Me.TabControl1.SuspendLayout()
        Me.SuspendLayout()
        '
        'TabPage4
        '
        Me.TabPage4.Controls.Add(Me.Label1)
        Me.TabPage4.Controls.Add(Me.btnLanguageExp)
        Me.TabPage4.Controls.Add(Me.btnLanguageImp)
        Me.TabPage4.Location = New System.Drawing.Point(4, 22)
        Me.TabPage4.Name = "TabPage4"
        Me.TabPage4.Size = New System.Drawing.Size(282, 156)
        Me.TabPage4.TabIndex = 3
        Me.TabPage4.Text = "Sprachen I/E"
        Me.TabPage4.UseVisualStyleBackColor = True
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
        'btnLanguageExp
        '
        Me.btnLanguageExp.Location = New System.Drawing.Point(153, 59)
        Me.btnLanguageExp.Name = "btnLanguageExp"
        Me.btnLanguageExp.Size = New System.Drawing.Size(75, 23)
        Me.btnLanguageExp.TabIndex = 1
        Me.btnLanguageExp.Text = "Exportieren"
        Me.btnLanguageExp.UseVisualStyleBackColor = True
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
        'TabPage1
        '
        Me.TabPage1.Controls.Add(Me.btnChangeLanguage)
        Me.TabPage1.Controls.Add(Me.txtboxLanguage)
        Me.TabPage1.Controls.Add(Me.lblLanguage)
        Me.TabPage1.Controls.Add(Me.Label5)
        Me.TabPage1.Controls.Add(Me.txtboxSchriftGroesse)
        Me.TabPage1.Controls.Add(Me.lbl_schrift)
        Me.TabPage1.Controls.Add(Me.txtboxAbstandsEinheit)
        Me.TabPage1.Location = New System.Drawing.Point(4, 22)
        Me.TabPage1.Name = "TabPage1"
        Me.TabPage1.Padding = New System.Windows.Forms.Padding(3)
        Me.TabPage1.Size = New System.Drawing.Size(282, 156)
        Me.TabPage1.TabIndex = 0
        Me.TabPage1.Text = "Allgemein"
        Me.TabPage1.UseVisualStyleBackColor = True
        '
        'txtboxAbstandsEinheit
        '
        Me.txtboxAbstandsEinheit.FormattingEnabled = True
        Me.txtboxAbstandsEinheit.Items.AddRange(New Object() {"Days", "Weeks", "Months"})
        Me.txtboxAbstandsEinheit.Location = New System.Drawing.Point(183, 28)
        Me.txtboxAbstandsEinheit.Name = "txtboxAbstandsEinheit"
        Me.txtboxAbstandsEinheit.Size = New System.Drawing.Size(83, 21)
        Me.txtboxAbstandsEinheit.TabIndex = 25
        Me.txtboxAbstandsEinheit.Text = "Days"
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
        'txtboxSchriftGroesse
        '
        Me.txtboxSchriftGroesse.Location = New System.Drawing.Point(183, 6)
        Me.txtboxSchriftGroesse.Name = "txtboxSchriftGroesse"
        Me.txtboxSchriftGroesse.Size = New System.Drawing.Size(83, 20)
        Me.txtboxSchriftGroesse.TabIndex = 19
        Me.txtboxSchriftGroesse.Text = "8"
        Me.txtboxSchriftGroesse.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'Label5
        '
        Me.Label5.AutoSize = True
        Me.Label5.Location = New System.Drawing.Point(2, 31)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(85, 13)
        Me.Label5.TabIndex = 36
        Me.Label5.Text = "Abstand messen"
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
        'txtboxLanguage
        '
        Me.txtboxLanguage.FormattingEnabled = True
        Me.txtboxLanguage.Location = New System.Drawing.Point(183, 51)
        Me.txtboxLanguage.Name = "txtboxLanguage"
        Me.txtboxLanguage.Size = New System.Drawing.Size(83, 21)
        Me.txtboxLanguage.TabIndex = 38
        Me.txtboxLanguage.Text = "Original"
        '
        'btnChangeLanguage
        '
        Me.btnChangeLanguage.Location = New System.Drawing.Point(183, 104)
        Me.btnChangeLanguage.Name = "btnChangeLanguage"
        Me.btnChangeLanguage.Size = New System.Drawing.Size(83, 40)
        Me.btnChangeLanguage.TabIndex = 39
        Me.btnChangeLanguage.Text = "Namen übersetzen"
        Me.btnChangeLanguage.UseVisualStyleBackColor = True
        '
        'TabControl1
        '
        Me.TabControl1.Controls.Add(Me.TabPage1)
        Me.TabControl1.Controls.Add(Me.TabPage4)
        Me.TabControl1.Location = New System.Drawing.Point(12, 12)
        Me.TabControl1.Name = "TabControl1"
        Me.TabControl1.SelectedIndex = 0
        Me.TabControl1.Size = New System.Drawing.Size(290, 182)
        Me.TabControl1.TabIndex = 35
        '
        'frmSettings
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(96.0!, 96.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Dpi
        Me.BackColor = System.Drawing.SystemColors.ControlLightLight
        Me.ClientSize = New System.Drawing.Size(323, 206)
        Me.Controls.Add(Me.TabControl1)
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.Name = "frmSettings"
        Me.Text = "Einstellungen"
        Me.TabPage4.ResumeLayout(False)
        Me.TabPage4.PerformLayout()
        Me.TabPage1.ResumeLayout(False)
        Me.TabPage1.PerformLayout()
        Me.TabControl1.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub

    Friend WithEvents TabPage4 As Windows.Forms.TabPage
    Friend WithEvents Label1 As Windows.Forms.Label
    Friend WithEvents btnLanguageExp As Windows.Forms.Button
    Friend WithEvents btnLanguageImp As Windows.Forms.Button
    Friend WithEvents TabPage1 As Windows.Forms.TabPage
    Friend WithEvents btnChangeLanguage As Windows.Forms.Button
    Public WithEvents txtboxLanguage As Windows.Forms.ComboBox
    Friend WithEvents lblLanguage As Windows.Forms.Label
    Friend WithEvents Label5 As Windows.Forms.Label
    Friend WithEvents txtboxSchriftGroesse As Windows.Forms.TextBox
    Friend WithEvents lbl_schrift As Windows.Forms.Label
    Friend WithEvents txtboxAbstandsEinheit As Windows.Forms.ComboBox
    Friend WithEvents TabControl1 As Windows.Forms.TabControl
End Class
