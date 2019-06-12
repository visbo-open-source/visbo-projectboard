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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmProjektEingabe1))
        Me.OKButton = New System.Windows.Forms.Button()
        Me.AbbrButton = New System.Windows.Forms.Button()
        Me.lblBudget = New System.Windows.Forms.Label()
        Me.lblVorlage = New System.Windows.Forms.Label()
        Me.projectName = New System.Windows.Forms.TextBox()
        Me.lbl_pName = New System.Windows.Forms.Label()
        Me.Erloes = New System.Windows.Forms.TextBox()
        Me.vorlagenDropbox = New System.Windows.Forms.ComboBox()
        Me.DateTimeStart = New System.Windows.Forms.DateTimePicker()
        Me.dauerUnverändert = New System.Windows.Forms.CheckBox()
        Me.DateTimeEnde = New System.Windows.Forms.DateTimePicker()
        Me.lbl_Referenz1 = New System.Windows.Forms.Label()
        Me.startMilestoneDropbox = New System.Windows.Forms.ComboBox()
        Me.lbl_Referenz2 = New System.Windows.Forms.Label()
        Me.endMilestoneDropbox = New System.Windows.Forms.ComboBox()
        Me.lbl_Laufzeit = New System.Windows.Forms.Label()
        Me.lblProfitField = New System.Windows.Forms.Label()
        Me.profitAskedFor = New System.Windows.Forms.TextBox()
        Me.lbl_Description = New System.Windows.Forms.Label()
        Me.txtbx_description = New System.Windows.Forms.TextBox()
        Me.txtbx_pNr = New System.Windows.Forms.TextBox()
        Me.lbl_Number = New System.Windows.Forms.Label()
        Me.SuspendLayout()
        '
        'OKButton
        '
        Me.OKButton.DialogResult = System.Windows.Forms.DialogResult.OK
        Me.OKButton.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.OKButton.Location = New System.Drawing.Point(114, 289)
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
        Me.AbbrButton.Location = New System.Drawing.Point(672, 289)
        Me.AbbrButton.Margin = New System.Windows.Forms.Padding(2, 3, 2, 3)
        Me.AbbrButton.Name = "AbbrButton"
        Me.AbbrButton.Size = New System.Drawing.Size(119, 22)
        Me.AbbrButton.TabIndex = 10
        Me.AbbrButton.Text = "Abbrechen"
        Me.AbbrButton.UseVisualStyleBackColor = True
        '
        'lblBudget
        '
        Me.lblBudget.AutoSize = True
        Me.lblBudget.Enabled = False
        Me.lblBudget.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblBudget.Location = New System.Drawing.Point(396, 68)
        Me.lblBudget.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.lblBudget.Name = "lblBudget"
        Me.lblBudget.Size = New System.Drawing.Size(78, 16)
        Me.lblBudget.TabIndex = 5
        Me.lblBudget.Text = "Budget (T€)"
        '
        'lblVorlage
        '
        Me.lblVorlage.AutoSize = True
        Me.lblVorlage.Enabled = False
        Me.lblVorlage.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblVorlage.Location = New System.Drawing.Point(396, 28)
        Me.lblVorlage.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.lblVorlage.Name = "lblVorlage"
        Me.lblVorlage.Size = New System.Drawing.Size(56, 16)
        Me.lblVorlage.TabIndex = 14
        Me.lblVorlage.Text = "Vorlage"
        Me.lblVorlage.TextAlign = System.Drawing.ContentAlignment.TopRight
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
        'lbl_pName
        '
        Me.lbl_pName.AutoSize = True
        Me.lbl_pName.Enabled = False
        Me.lbl_pName.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lbl_pName.Location = New System.Drawing.Point(11, 28)
        Me.lbl_pName.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.lbl_pName.Name = "lbl_pName"
        Me.lbl_pName.Size = New System.Drawing.Size(45, 16)
        Me.lbl_pName.TabIndex = 16
        Me.lbl_pName.Text = "Name"
        '
        'Erloes
        '
        Me.Erloes.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Erloes.Location = New System.Drawing.Point(482, 62)
        Me.Erloes.Margin = New System.Windows.Forms.Padding(2)
        Me.Erloes.Name = "Erloes"
        Me.Erloes.Size = New System.Drawing.Size(88, 22)
        Me.Erloes.TabIndex = 20
        '
        'vorlagenDropbox
        '
        Me.vorlagenDropbox.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.vorlagenDropbox.FormattingEnabled = True
        Me.vorlagenDropbox.Location = New System.Drawing.Point(482, 25)
        Me.vorlagenDropbox.Margin = New System.Windows.Forms.Padding(2)
        Me.vorlagenDropbox.Name = "vorlagenDropbox"
        Me.vorlagenDropbox.Size = New System.Drawing.Size(309, 24)
        Me.vorlagenDropbox.TabIndex = 23
        '
        'DateTimeStart
        '
        Me.DateTimeStart.CustomFormat = ""
        Me.DateTimeStart.Format = System.Windows.Forms.DateTimePickerFormat.[Short]
        Me.DateTimeStart.Location = New System.Drawing.Point(672, 217)
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
        Me.dauerUnverändert.Location = New System.Drawing.Point(114, 185)
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
        Me.DateTimeEnde.Location = New System.Drawing.Point(672, 248)
        Me.DateTimeEnde.Margin = New System.Windows.Forms.Padding(2)
        Me.DateTimeEnde.Name = "DateTimeEnde"
        Me.DateTimeEnde.Size = New System.Drawing.Size(119, 22)
        Me.DateTimeEnde.TabIndex = 29
        '
        'lbl_Referenz1
        '
        Me.lbl_Referenz1.AutoSize = True
        Me.lbl_Referenz1.Enabled = False
        Me.lbl_Referenz1.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lbl_Referenz1.Location = New System.Drawing.Point(13, 219)
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
        Me.startMilestoneDropbox.Location = New System.Drawing.Point(114, 215)
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
        Me.lbl_Referenz2.Location = New System.Drawing.Point(13, 250)
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
        Me.endMilestoneDropbox.Location = New System.Drawing.Point(114, 246)
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
        Me.lbl_Laufzeit.Location = New System.Drawing.Point(304, 186)
        Me.lbl_Laufzeit.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.lbl_Laufzeit.Name = "lbl_Laufzeit"
        Me.lbl_Laufzeit.Size = New System.Drawing.Size(53, 16)
        Me.lbl_Laufzeit.TabIndex = 39
        Me.lbl_Laufzeit.Text = "Laufzeit"
        '
        'lblProfitField
        '
        Me.lblProfitField.AutoSize = True
        Me.lblProfitField.Enabled = False
        Me.lblProfitField.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblProfitField.Location = New System.Drawing.Point(631, 68)
        Me.lblProfitField.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.lblProfitField.Name = "lblProfitField"
        Me.lblProfitField.Size = New System.Drawing.Size(78, 16)
        Me.lblProfitField.TabIndex = 41
        Me.lblProfitField.Text = "Rendite (%)"
        '
        'profitAskedFor
        '
        Me.profitAskedFor.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.profitAskedFor.Location = New System.Drawing.Point(737, 65)
        Me.profitAskedFor.Margin = New System.Windows.Forms.Padding(2)
        Me.profitAskedFor.Name = "profitAskedFor"
        Me.profitAskedFor.Size = New System.Drawing.Size(54, 22)
        Me.profitAskedFor.TabIndex = 42
        '
        'lbl_Description
        '
        Me.lbl_Description.AutoSize = True
        Me.lbl_Description.Enabled = False
        Me.lbl_Description.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lbl_Description.Location = New System.Drawing.Point(13, 113)
        Me.lbl_Description.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.lbl_Description.Name = "lbl_Description"
        Me.lbl_Description.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.lbl_Description.Size = New System.Drawing.Size(38, 16)
        Me.lbl_Description.TabIndex = 43
        Me.lbl_Description.Text = "Ziele"
        '
        'txtbx_description
        '
        Me.txtbx_description.AcceptsReturn = True
        Me.txtbx_description.Location = New System.Drawing.Point(114, 105)
        Me.txtbx_description.Multiline = True
        Me.txtbx_description.Name = "txtbx_description"
        Me.txtbx_description.ScrollBars = System.Windows.Forms.ScrollBars.Vertical
        Me.txtbx_description.Size = New System.Drawing.Size(677, 65)
        Me.txtbx_description.TabIndex = 44
        '
        'txtbx_pNr
        '
        Me.txtbx_pNr.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtbx_pNr.Location = New System.Drawing.Point(114, 62)
        Me.txtbx_pNr.Margin = New System.Windows.Forms.Padding(2)
        Me.txtbx_pNr.Name = "txtbx_pNr"
        Me.txtbx_pNr.Size = New System.Drawing.Size(254, 22)
        Me.txtbx_pNr.TabIndex = 45
        '
        'lbl_Number
        '
        Me.lbl_Number.AutoSize = True
        Me.lbl_Number.Enabled = False
        Me.lbl_Number.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lbl_Number.Location = New System.Drawing.Point(11, 62)
        Me.lbl_Number.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.lbl_Number.Name = "lbl_Number"
        Me.lbl_Number.Size = New System.Drawing.Size(59, 16)
        Me.lbl_Number.TabIndex = 46
        Me.lbl_Number.Text = "Nummer"
        '
        'frmProjektEingabe1
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(96.0!, 96.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Dpi
        Me.BackColor = System.Drawing.SystemColors.Window
        Me.ClientSize = New System.Drawing.Size(827, 327)
        Me.Controls.Add(Me.lbl_Number)
        Me.Controls.Add(Me.txtbx_pNr)
        Me.Controls.Add(Me.txtbx_description)
        Me.Controls.Add(Me.lbl_Description)
        Me.Controls.Add(Me.profitAskedFor)
        Me.Controls.Add(Me.lblProfitField)
        Me.Controls.Add(Me.lbl_Laufzeit)
        Me.Controls.Add(Me.endMilestoneDropbox)
        Me.Controls.Add(Me.lbl_Referenz2)
        Me.Controls.Add(Me.startMilestoneDropbox)
        Me.Controls.Add(Me.lbl_Referenz1)
        Me.Controls.Add(Me.DateTimeEnde)
        Me.Controls.Add(Me.dauerUnverändert)
        Me.Controls.Add(Me.DateTimeStart)
        Me.Controls.Add(Me.vorlagenDropbox)
        Me.Controls.Add(Me.Erloes)
        Me.Controls.Add(Me.lbl_pName)
        Me.Controls.Add(Me.projectName)
        Me.Controls.Add(Me.lblVorlage)
        Me.Controls.Add(Me.lblBudget)
        Me.Controls.Add(Me.AbbrButton)
        Me.Controls.Add(Me.OKButton)
        Me.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.Margin = New System.Windows.Forms.Padding(2, 3, 2, 3)
        Me.Name = "frmProjektEingabe1"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.Manual
        Me.Text = "Neues Projekt anlegen"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Public WithEvents OKButton As System.Windows.Forms.Button
    Public WithEvents AbbrButton As System.Windows.Forms.Button
    Public WithEvents lblBudget As System.Windows.Forms.Label
    Public WithEvents lblVorlage As System.Windows.Forms.Label
    Public WithEvents projectName As System.Windows.Forms.TextBox
    Public WithEvents lbl_pName As System.Windows.Forms.Label
    Public WithEvents Erloes As System.Windows.Forms.TextBox
    Public WithEvents vorlagenDropbox As System.Windows.Forms.ComboBox
    Public WithEvents DateTimeStart As System.Windows.Forms.DateTimePicker
    Friend WithEvents dauerUnverändert As System.Windows.Forms.CheckBox
    Friend WithEvents DateTimeEnde As System.Windows.Forms.DateTimePicker
    Public WithEvents lbl_Referenz1 As System.Windows.Forms.Label
    Public WithEvents startMilestoneDropbox As System.Windows.Forms.ComboBox
    Public WithEvents lbl_Referenz2 As System.Windows.Forms.Label
    Public WithEvents endMilestoneDropbox As System.Windows.Forms.ComboBox
    Public WithEvents lbl_Laufzeit As System.Windows.Forms.Label
    Public WithEvents lblProfitField As System.Windows.Forms.Label
    Public WithEvents profitAskedFor As System.Windows.Forms.TextBox
    Public WithEvents lbl_Description As Windows.Forms.Label
    Public WithEvents txtbx_description As Windows.Forms.TextBox
    Public WithEvents txtbx_pNr As Windows.Forms.TextBox
    Public WithEvents lbl_Number As Windows.Forms.Label
End Class
