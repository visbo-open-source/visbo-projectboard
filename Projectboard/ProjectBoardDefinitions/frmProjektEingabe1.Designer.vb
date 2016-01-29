<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmProjektEingabe1
    Inherits System.Windows.Forms.Form

    'Form overrides dispose to clean up the component list.
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

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Me.OKButton = New System.Windows.Forms.Button()
        Me.AbbrButton = New System.Windows.Forms.Button()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.projectName = New System.Windows.Forms.TextBox()
        Me.pName = New System.Windows.Forms.Label()
        Me.Erloes = New System.Windows.Forms.TextBox()
        Me.sFit = New System.Windows.Forms.TextBox()
        Me.risiko = New System.Windows.Forms.TextBox()
        Me.vorlagenDropbox = New System.Windows.Forms.ComboBox()
        Me.Label5 = New System.Windows.Forms.Label()
        Me.volume = New System.Windows.Forms.TextBox()
        Me.DateTimeStart = New System.Windows.Forms.DateTimePicker()
        Me.dauerUnverändert = New System.Windows.Forms.CheckBox()
        Me.DateTimeEnde = New System.Windows.Forms.DateTimePicker()
        Me.lbl_BusinessUnit = New System.Windows.Forms.Label()
        Me.businessUnitDropBox = New System.Windows.Forms.ComboBox()
        Me.lbl_Referenz1 = New System.Windows.Forms.Label()
        Me.startMilestoneDropbox = New System.Windows.Forms.ComboBox()
        Me.lbl_Referenz2 = New System.Windows.Forms.Label()
        Me.endMilestoneDropbox = New System.Windows.Forms.ComboBox()
        Me.lbl_Laufzeit = New System.Windows.Forms.Label()
        Me.propRessourcenAnpassung = New System.Windows.Forms.CheckBox()
        Me.SuspendLayout()
        '
        'OKButton
        '
        Me.OKButton.DialogResult = System.Windows.Forms.DialogResult.OK
        Me.OKButton.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.OKButton.Location = New System.Drawing.Point(114, 278)
        Me.OKButton.Margin = New System.Windows.Forms.Padding(2, 3, 2, 3)
        Me.OKButton.Name = "OKButton"
        Me.OKButton.Size = New System.Drawing.Size(119, 22)
        Me.OKButton.TabIndex = 1
        Me.OKButton.Text = "OK"
        Me.OKButton.UseVisualStyleBackColor = True
        '
        'AbbrButton
        '
        Me.AbbrButton.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.AbbrButton.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.AbbrButton.Location = New System.Drawing.Point(672, 278)
        Me.AbbrButton.Margin = New System.Windows.Forms.Padding(2, 3, 2, 3)
        Me.AbbrButton.Name = "AbbrButton"
        Me.AbbrButton.Size = New System.Drawing.Size(119, 22)
        Me.AbbrButton.TabIndex = 10
        Me.AbbrButton.Text = "Abbrechen"
        Me.AbbrButton.UseVisualStyleBackColor = True
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Enabled = False
        Me.Label1.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label1.Location = New System.Drawing.Point(13, 130)
        Me.Label1.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(78, 16)
        Me.Label1.TabIndex = 5
        Me.Label1.Text = "Budget (T€)"
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Enabled = False
        Me.Label2.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label2.Location = New System.Drawing.Point(199, 130)
        Me.Label2.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(104, 16)
        Me.Label2.TabIndex = 6
        Me.Label2.Text = "Strategischer Fit"
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Enabled = False
        Me.Label3.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label3.Location = New System.Drawing.Point(403, 130)
        Me.Label3.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(124, 16)
        Me.Label3.TabIndex = 7
        Me.Label3.Text = "Umsetzungs-Risiko"
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Enabled = False
        Me.Label4.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label4.Location = New System.Drawing.Point(447, 28)
        Me.Label4.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(56, 16)
        Me.Label4.TabIndex = 14
        Me.Label4.Text = "Vorlage"
        '
        'projectName
        '
        Me.projectName.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.projectName.Location = New System.Drawing.Point(114, 25)
        Me.projectName.Margin = New System.Windows.Forms.Padding(2)
        Me.projectName.Name = "projectName"
        Me.projectName.Size = New System.Drawing.Size(254, 22)
        Me.projectName.TabIndex = 0
        '
        'pName
        '
        Me.pName.AutoSize = True
        Me.pName.Enabled = False
        Me.pName.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.pName.Location = New System.Drawing.Point(13, 28)
        Me.pName.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.pName.Name = "pName"
        Me.pName.Size = New System.Drawing.Size(91, 16)
        Me.pName.TabIndex = 16
        Me.pName.Text = "Projekt-Name"
        '
        'Erloes
        '
        Me.Erloes.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Erloes.Location = New System.Drawing.Point(114, 127)
        Me.Erloes.Margin = New System.Windows.Forms.Padding(2)
        Me.Erloes.Name = "Erloes"
        Me.Erloes.Size = New System.Drawing.Size(60, 22)
        Me.Erloes.TabIndex = 20
        '
        'sFit
        '
        Me.sFit.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.sFit.Location = New System.Drawing.Point(307, 127)
        Me.sFit.Margin = New System.Windows.Forms.Padding(2)
        Me.sFit.Name = "sFit"
        Me.sFit.Size = New System.Drawing.Size(60, 22)
        Me.sFit.TabIndex = 21
        '
        'risiko
        '
        Me.risiko.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.risiko.Location = New System.Drawing.Point(531, 127)
        Me.risiko.Margin = New System.Windows.Forms.Padding(2)
        Me.risiko.Name = "risiko"
        Me.risiko.Size = New System.Drawing.Size(60, 22)
        Me.risiko.TabIndex = 22
        '
        'vorlagenDropbox
        '
        Me.vorlagenDropbox.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.vorlagenDropbox.FormattingEnabled = True
        Me.vorlagenDropbox.Location = New System.Drawing.Point(506, 25)
        Me.vorlagenDropbox.Margin = New System.Windows.Forms.Padding(2)
        Me.vorlagenDropbox.Name = "vorlagenDropbox"
        Me.vorlagenDropbox.Size = New System.Drawing.Size(285, 24)
        Me.vorlagenDropbox.TabIndex = 23
        '
        'Label5
        '
        Me.Label5.AutoEllipsis = True
        Me.Label5.AutoSize = True
        Me.Label5.Enabled = False
        Me.Label5.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label5.Location = New System.Drawing.Point(628, 130)
        Me.Label5.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(61, 16)
        Me.Label5.TabIndex = 24
        Me.Label5.Text = "Volumen"
        Me.Label5.Visible = False
        '
        'volume
        '
        Me.volume.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.volume.Location = New System.Drawing.Point(693, 127)
        Me.volume.Margin = New System.Windows.Forms.Padding(2)
        Me.volume.Name = "volume"
        Me.volume.Size = New System.Drawing.Size(98, 22)
        Me.volume.TabIndex = 25
        Me.volume.Visible = False
        '
        'DateTimeStart
        '
        Me.DateTimeStart.CustomFormat = ""
        Me.DateTimeStart.Format = System.Windows.Forms.DateTimePickerFormat.[Short]
        Me.DateTimeStart.Location = New System.Drawing.Point(672, 194)
        Me.DateTimeStart.Margin = New System.Windows.Forms.Padding(2)
        Me.DateTimeStart.Name = "DateTimeStart"
        Me.DateTimeStart.Size = New System.Drawing.Size(119, 22)
        Me.DateTimeStart.TabIndex = 26
        '
        'dauerUnverändert
        '
        Me.dauerUnverändert.AutoSize = True
        Me.dauerUnverändert.Checked = True
        Me.dauerUnverändert.CheckState = System.Windows.Forms.CheckState.Checked
        Me.dauerUnverändert.Location = New System.Drawing.Point(114, 162)
        Me.dauerUnverändert.Margin = New System.Windows.Forms.Padding(2)
        Me.dauerUnverändert.Name = "dauerUnverändert"
        Me.dauerUnverändert.Size = New System.Drawing.Size(138, 20)
        Me.dauerUnverändert.TabIndex = 27
        Me.dauerUnverändert.Text = "Dauer wie Vorlage"
        Me.dauerUnverändert.UseVisualStyleBackColor = True
        '
        'DateTimeEnde
        '
        Me.DateTimeEnde.Format = System.Windows.Forms.DateTimePickerFormat.[Short]
        Me.DateTimeEnde.Location = New System.Drawing.Point(672, 225)
        Me.DateTimeEnde.Margin = New System.Windows.Forms.Padding(2)
        Me.DateTimeEnde.Name = "DateTimeEnde"
        Me.DateTimeEnde.Size = New System.Drawing.Size(119, 22)
        Me.DateTimeEnde.TabIndex = 29
        '
        'lbl_BusinessUnit
        '
        Me.lbl_BusinessUnit.AutoSize = True
        Me.lbl_BusinessUnit.Enabled = False
        Me.lbl_BusinessUnit.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lbl_BusinessUnit.Location = New System.Drawing.Point(13, 65)
        Me.lbl_BusinessUnit.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.lbl_BusinessUnit.Name = "lbl_BusinessUnit"
        Me.lbl_BusinessUnit.Size = New System.Drawing.Size(89, 16)
        Me.lbl_BusinessUnit.TabIndex = 33
        Me.lbl_BusinessUnit.Text = "Business Unit"
        '
        'businessUnitDropBox
        '
        Me.businessUnitDropBox.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.businessUnitDropBox.FormattingEnabled = True
        Me.businessUnitDropBox.Location = New System.Drawing.Point(114, 60)
        Me.businessUnitDropBox.Margin = New System.Windows.Forms.Padding(2)
        Me.businessUnitDropBox.Name = "businessUnitDropBox"
        Me.businessUnitDropBox.Size = New System.Drawing.Size(254, 24)
        Me.businessUnitDropBox.TabIndex = 34
        '
        'lbl_Referenz1
        '
        Me.lbl_Referenz1.AutoSize = True
        Me.lbl_Referenz1.Enabled = False
        Me.lbl_Referenz1.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lbl_Referenz1.Location = New System.Drawing.Point(13, 196)
        Me.lbl_Referenz1.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.lbl_Referenz1.Name = "lbl_Referenz1"
        Me.lbl_Referenz1.Size = New System.Drawing.Size(86, 16)
        Me.lbl_Referenz1.TabIndex = 35
        Me.lbl_Referenz1.Text = "Meilenstein 1"
        '
        'startMilestoneDropbox
        '
        Me.startMilestoneDropbox.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.startMilestoneDropbox.FormattingEnabled = True
        Me.startMilestoneDropbox.Location = New System.Drawing.Point(114, 192)
        Me.startMilestoneDropbox.Margin = New System.Windows.Forms.Padding(2)
        Me.startMilestoneDropbox.Name = "startMilestoneDropbox"
        Me.startMilestoneDropbox.Size = New System.Drawing.Size(532, 24)
        Me.startMilestoneDropbox.TabIndex = 36
        '
        'lbl_Referenz2
        '
        Me.lbl_Referenz2.AutoSize = True
        Me.lbl_Referenz2.Enabled = False
        Me.lbl_Referenz2.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lbl_Referenz2.Location = New System.Drawing.Point(13, 227)
        Me.lbl_Referenz2.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.lbl_Referenz2.Name = "lbl_Referenz2"
        Me.lbl_Referenz2.Size = New System.Drawing.Size(86, 16)
        Me.lbl_Referenz2.TabIndex = 37
        Me.lbl_Referenz2.Text = "Meilenstein 2"
        '
        'endMilestoneDropbox
        '
        Me.endMilestoneDropbox.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.endMilestoneDropbox.FormattingEnabled = True
        Me.endMilestoneDropbox.Location = New System.Drawing.Point(114, 223)
        Me.endMilestoneDropbox.Margin = New System.Windows.Forms.Padding(2)
        Me.endMilestoneDropbox.Name = "endMilestoneDropbox"
        Me.endMilestoneDropbox.Size = New System.Drawing.Size(532, 24)
        Me.endMilestoneDropbox.TabIndex = 38
        '
        'lbl_Laufzeit
        '
        Me.lbl_Laufzeit.AutoSize = True
        Me.lbl_Laufzeit.Enabled = False
        Me.lbl_Laufzeit.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lbl_Laufzeit.Location = New System.Drawing.Point(447, 68)
        Me.lbl_Laufzeit.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.lbl_Laufzeit.Name = "lbl_Laufzeit"
        Me.lbl_Laufzeit.Size = New System.Drawing.Size(53, 16)
        Me.lbl_Laufzeit.TabIndex = 39
        Me.lbl_Laufzeit.Text = "Laufzeit"
        '
        'propRessourcenAnpassung
        '
        Me.propRessourcenAnpassung.AutoSize = True
        Me.propRessourcenAnpassung.Checked = True
        Me.propRessourcenAnpassung.CheckState = System.Windows.Forms.CheckState.Checked
        Me.propRessourcenAnpassung.Location = New System.Drawing.Point(400, 254)
        Me.propRessourcenAnpassung.Margin = New System.Windows.Forms.Padding(2)
        Me.propRessourcenAnpassung.Name = "propRessourcenAnpassung"
        Me.propRessourcenAnpassung.Size = New System.Drawing.Size(428, 20)
        Me.propRessourcenAnpassung.TabIndex = 40
        Me.propRessourcenAnpassung.Text = "Ressourcen und Kostenbedarfe der Vorlage proportional anpassen"
        Me.propRessourcenAnpassung.UseVisualStyleBackColor = True
        '
        'frmProjektEingabe1
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(96.0!, 96.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Dpi
        Me.BackColor = System.Drawing.SystemColors.Window
        Me.ClientSize = New System.Drawing.Size(827, 329)
        Me.Controls.Add(Me.propRessourcenAnpassung)
        Me.Controls.Add(Me.lbl_Laufzeit)
        Me.Controls.Add(Me.endMilestoneDropbox)
        Me.Controls.Add(Me.lbl_Referenz2)
        Me.Controls.Add(Me.startMilestoneDropbox)
        Me.Controls.Add(Me.lbl_Referenz1)
        Me.Controls.Add(Me.businessUnitDropBox)
        Me.Controls.Add(Me.lbl_BusinessUnit)
        Me.Controls.Add(Me.DateTimeEnde)
        Me.Controls.Add(Me.dauerUnverändert)
        Me.Controls.Add(Me.DateTimeStart)
        Me.Controls.Add(Me.volume)
        Me.Controls.Add(Me.Label5)
        Me.Controls.Add(Me.vorlagenDropbox)
        Me.Controls.Add(Me.risiko)
        Me.Controls.Add(Me.sFit)
        Me.Controls.Add(Me.Erloes)
        Me.Controls.Add(Me.pName)
        Me.Controls.Add(Me.projectName)
        Me.Controls.Add(Me.Label4)
        Me.Controls.Add(Me.Label3)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.AbbrButton)
        Me.Controls.Add(Me.OKButton)
        Me.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Margin = New System.Windows.Forms.Padding(2, 3, 2, 3)
        Me.Name = "frmProjektEingabe1"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.Manual
        Me.Text = "Neues Projekt anlegen"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Public WithEvents OKButton As System.Windows.Forms.Button
    Public WithEvents AbbrButton As System.Windows.Forms.Button
    Public WithEvents Label1 As System.Windows.Forms.Label
    Public WithEvents Label2 As System.Windows.Forms.Label
    Public WithEvents Label3 As System.Windows.Forms.Label
    Public WithEvents Label4 As System.Windows.Forms.Label
    Public WithEvents projectName As System.Windows.Forms.TextBox
    Public WithEvents pName As System.Windows.Forms.Label
    Public WithEvents Erloes As System.Windows.Forms.TextBox
    Public WithEvents sFit As System.Windows.Forms.TextBox
    Public WithEvents risiko As System.Windows.Forms.TextBox
    Public WithEvents vorlagenDropbox As System.Windows.Forms.ComboBox
    Public WithEvents Label5 As System.Windows.Forms.Label
    Public WithEvents volume As System.Windows.Forms.TextBox
    Public WithEvents DateTimeStart As System.Windows.Forms.DateTimePicker
    Friend WithEvents dauerUnverändert As System.Windows.Forms.CheckBox
    Friend WithEvents DateTimeEnde As System.Windows.Forms.DateTimePicker
    Public WithEvents lbl_BusinessUnit As System.Windows.Forms.Label
    Public WithEvents businessUnitDropBox As System.Windows.Forms.ComboBox
    Public WithEvents lbl_Referenz1 As System.Windows.Forms.Label
    Public WithEvents startMilestoneDropbox As System.Windows.Forms.ComboBox
    Public WithEvents lbl_Referenz2 As System.Windows.Forms.Label
    Public WithEvents endMilestoneDropbox As System.Windows.Forms.ComboBox
    Public WithEvents lbl_Laufzeit As System.Windows.Forms.Label
    Friend WithEvents propRessourcenAnpassung As System.Windows.Forms.CheckBox
End Class
