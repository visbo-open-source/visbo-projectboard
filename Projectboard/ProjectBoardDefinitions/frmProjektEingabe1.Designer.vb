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
        Me.Panel1 = New System.Windows.Forms.Panel()
        Me.lbl_selProjName = New System.Windows.Forms.Label()
        Me.lbl_Number = New System.Windows.Forms.Label()
        Me.txtbx_pNr = New System.Windows.Forms.TextBox()
        Me.txtbx_description = New System.Windows.Forms.TextBox()
        Me.lbl_Description = New System.Windows.Forms.Label()
        Me.profitAskedFor = New System.Windows.Forms.TextBox()
        Me.lblProfitField = New System.Windows.Forms.Label()
        Me.lbl_Laufzeit = New System.Windows.Forms.Label()
        Me.endMilestoneDropbox = New System.Windows.Forms.ComboBox()
        Me.lbl_Referenz2 = New System.Windows.Forms.Label()
        Me.startMilestoneDropbox = New System.Windows.Forms.ComboBox()
        Me.lbl_Referenz1 = New System.Windows.Forms.Label()
        Me.DateTimeEnde = New System.Windows.Forms.DateTimePicker()
        Me.dauerUnverändert = New System.Windows.Forms.CheckBox()
        Me.DateTimeStart = New System.Windows.Forms.DateTimePicker()
        Me.vorlagenDropbox = New System.Windows.Forms.ComboBox()
        Me.Erloes = New System.Windows.Forms.TextBox()
        Me.lbl_pName = New System.Windows.Forms.Label()
        Me.projectName = New System.Windows.Forms.TextBox()
        Me.lblVorlage = New System.Windows.Forms.Label()
        Me.lblBudget = New System.Windows.Forms.Label()
        Me.AbbrButton = New System.Windows.Forms.Button()
        Me.OKButton = New System.Windows.Forms.Button()
        Me.Panel1.SuspendLayout()
        Me.SuspendLayout()
        '
        'Panel1
        '
        Me.Panel1.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Panel1.Controls.Add(Me.lbl_selProjName)
        Me.Panel1.Controls.Add(Me.lbl_Number)
        Me.Panel1.Controls.Add(Me.txtbx_pNr)
        Me.Panel1.Controls.Add(Me.txtbx_description)
        Me.Panel1.Controls.Add(Me.lbl_Description)
        Me.Panel1.Controls.Add(Me.profitAskedFor)
        Me.Panel1.Controls.Add(Me.lblProfitField)
        Me.Panel1.Controls.Add(Me.lbl_Laufzeit)
        Me.Panel1.Controls.Add(Me.endMilestoneDropbox)
        Me.Panel1.Controls.Add(Me.lbl_Referenz2)
        Me.Panel1.Controls.Add(Me.startMilestoneDropbox)
        Me.Panel1.Controls.Add(Me.lbl_Referenz1)
        Me.Panel1.Controls.Add(Me.DateTimeEnde)
        Me.Panel1.Controls.Add(Me.dauerUnverändert)
        Me.Panel1.Controls.Add(Me.DateTimeStart)
        Me.Panel1.Controls.Add(Me.vorlagenDropbox)
        Me.Panel1.Controls.Add(Me.Erloes)
        Me.Panel1.Controls.Add(Me.lbl_pName)
        Me.Panel1.Controls.Add(Me.projectName)
        Me.Panel1.Controls.Add(Me.lblVorlage)
        Me.Panel1.Controls.Add(Me.lblBudget)
        Me.Panel1.Controls.Add(Me.AbbrButton)
        Me.Panel1.Controls.Add(Me.OKButton)
        Me.Panel1.Location = New System.Drawing.Point(2, 2)
        Me.Panel1.Name = "Panel1"
        Me.Panel1.Size = New System.Drawing.Size(884, 350)
        Me.Panel1.TabIndex = 0
        '
        'lbl_selProjName
        '
        Me.lbl_selProjName.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.lbl_selProjName.AutoSize = True
        Me.lbl_selProjName.Location = New System.Drawing.Point(535, 9)
        Me.lbl_selProjName.Name = "lbl_selProjName"
        Me.lbl_selProjName.Size = New System.Drawing.Size(174, 16)
        Me.lbl_selProjName.TabIndex = 70
        Me.lbl_selProjName.Text = "<Projekt-Name selektieren>"
        '
        'lbl_Number
        '
        Me.lbl_Number.AutoSize = True
        Me.lbl_Number.Enabled = False
        Me.lbl_Number.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lbl_Number.Location = New System.Drawing.Point(21, 64)
        Me.lbl_Number.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.lbl_Number.Name = "lbl_Number"
        Me.lbl_Number.Size = New System.Drawing.Size(59, 16)
        Me.lbl_Number.TabIndex = 52
        Me.lbl_Number.Text = "Nummer"
        '
        'txtbx_pNr
        '
        Me.txtbx_pNr.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtbx_pNr.Location = New System.Drawing.Point(124, 64)
        Me.txtbx_pNr.Margin = New System.Windows.Forms.Padding(2)
        Me.txtbx_pNr.Name = "txtbx_pNr"
        Me.txtbx_pNr.Size = New System.Drawing.Size(254, 22)
        Me.txtbx_pNr.TabIndex = 53
        '
        'txtbx_description
        '
        Me.txtbx_description.AcceptsReturn = True
        Me.txtbx_description.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtbx_description.Location = New System.Drawing.Point(124, 107)
        Me.txtbx_description.Multiline = True
        Me.txtbx_description.Name = "txtbx_description"
        Me.txtbx_description.ScrollBars = System.Windows.Forms.ScrollBars.Vertical
        Me.txtbx_description.Size = New System.Drawing.Size(723, 65)
        Me.txtbx_description.TabIndex = 59
        '
        'lbl_Description
        '
        Me.lbl_Description.AutoSize = True
        Me.lbl_Description.Enabled = False
        Me.lbl_Description.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lbl_Description.Location = New System.Drawing.Point(23, 115)
        Me.lbl_Description.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.lbl_Description.Name = "lbl_Description"
        Me.lbl_Description.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.lbl_Description.Size = New System.Drawing.Size(38, 16)
        Me.lbl_Description.TabIndex = 58
        Me.lbl_Description.Text = "Ziele"
        '
        'profitAskedFor
        '
        Me.profitAskedFor.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.profitAskedFor.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.profitAskedFor.Location = New System.Drawing.Point(791, 67)
        Me.profitAskedFor.Margin = New System.Windows.Forms.Padding(2)
        Me.profitAskedFor.Name = "profitAskedFor"
        Me.profitAskedFor.Size = New System.Drawing.Size(54, 22)
        Me.profitAskedFor.TabIndex = 57
        '
        'lblProfitField
        '
        Me.lblProfitField.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.lblProfitField.AutoSize = True
        Me.lblProfitField.Enabled = False
        Me.lblProfitField.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblProfitField.Location = New System.Drawing.Point(685, 70)
        Me.lblProfitField.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.lblProfitField.Name = "lblProfitField"
        Me.lblProfitField.Size = New System.Drawing.Size(78, 16)
        Me.lblProfitField.TabIndex = 56
        Me.lblProfitField.Text = "Rendite (%)"
        '
        'lbl_Laufzeit
        '
        Me.lbl_Laufzeit.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.lbl_Laufzeit.AutoSize = True
        Me.lbl_Laufzeit.Enabled = False
        Me.lbl_Laufzeit.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lbl_Laufzeit.Location = New System.Drawing.Point(314, 199)
        Me.lbl_Laufzeit.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.lbl_Laufzeit.Name = "lbl_Laufzeit"
        Me.lbl_Laufzeit.Size = New System.Drawing.Size(53, 16)
        Me.lbl_Laufzeit.TabIndex = 61
        Me.lbl_Laufzeit.Text = "Laufzeit"
        '
        'endMilestoneDropbox
        '
        Me.endMilestoneDropbox.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.endMilestoneDropbox.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.endMilestoneDropbox.FormattingEnabled = True
        Me.endMilestoneDropbox.Location = New System.Drawing.Point(124, 259)
        Me.endMilestoneDropbox.Margin = New System.Windows.Forms.Padding(2)
        Me.endMilestoneDropbox.Name = "endMilestoneDropbox"
        Me.endMilestoneDropbox.Size = New System.Drawing.Size(532, 24)
        Me.endMilestoneDropbox.TabIndex = 68
        '
        'lbl_Referenz2
        '
        Me.lbl_Referenz2.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.lbl_Referenz2.AutoSize = True
        Me.lbl_Referenz2.Enabled = False
        Me.lbl_Referenz2.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lbl_Referenz2.Location = New System.Drawing.Point(23, 263)
        Me.lbl_Referenz2.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.lbl_Referenz2.Name = "lbl_Referenz2"
        Me.lbl_Referenz2.Size = New System.Drawing.Size(86, 16)
        Me.lbl_Referenz2.TabIndex = 67
        Me.lbl_Referenz2.Text = "Meilenstein 2"
        '
        'startMilestoneDropbox
        '
        Me.startMilestoneDropbox.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.startMilestoneDropbox.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.startMilestoneDropbox.FormattingEnabled = True
        Me.startMilestoneDropbox.Location = New System.Drawing.Point(124, 228)
        Me.startMilestoneDropbox.Margin = New System.Windows.Forms.Padding(2)
        Me.startMilestoneDropbox.Name = "startMilestoneDropbox"
        Me.startMilestoneDropbox.Size = New System.Drawing.Size(532, 24)
        Me.startMilestoneDropbox.TabIndex = 66
        '
        'lbl_Referenz1
        '
        Me.lbl_Referenz1.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.lbl_Referenz1.AutoSize = True
        Me.lbl_Referenz1.Enabled = False
        Me.lbl_Referenz1.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lbl_Referenz1.Location = New System.Drawing.Point(23, 232)
        Me.lbl_Referenz1.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.lbl_Referenz1.Name = "lbl_Referenz1"
        Me.lbl_Referenz1.Size = New System.Drawing.Size(86, 16)
        Me.lbl_Referenz1.TabIndex = 65
        Me.lbl_Referenz1.Text = "Meilenstein 1"
        '
        'DateTimeEnde
        '
        Me.DateTimeEnde.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.DateTimeEnde.Format = System.Windows.Forms.DateTimePickerFormat.[Short]
        Me.DateTimeEnde.Location = New System.Drawing.Point(725, 261)
        Me.DateTimeEnde.Margin = New System.Windows.Forms.Padding(2)
        Me.DateTimeEnde.Name = "DateTimeEnde"
        Me.DateTimeEnde.Size = New System.Drawing.Size(119, 22)
        Me.DateTimeEnde.TabIndex = 63
        '
        'dauerUnverändert
        '
        Me.dauerUnverändert.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.dauerUnverändert.AutoSize = True
        Me.dauerUnverändert.Checked = True
        Me.dauerUnverändert.CheckState = System.Windows.Forms.CheckState.Checked
        Me.dauerUnverändert.Location = New System.Drawing.Point(124, 198)
        Me.dauerUnverändert.Margin = New System.Windows.Forms.Padding(2)
        Me.dauerUnverändert.Name = "dauerUnverändert"
        Me.dauerUnverändert.Size = New System.Drawing.Size(138, 20)
        Me.dauerUnverändert.TabIndex = 60
        Me.dauerUnverändert.Text = "Dauer wie Vorlage"
        Me.dauerUnverändert.UseVisualStyleBackColor = True
        '
        'DateTimeStart
        '
        Me.DateTimeStart.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.DateTimeStart.CustomFormat = ""
        Me.DateTimeStart.Format = System.Windows.Forms.DateTimePickerFormat.[Short]
        Me.DateTimeStart.Location = New System.Drawing.Point(725, 230)
        Me.DateTimeStart.Margin = New System.Windows.Forms.Padding(2)
        Me.DateTimeStart.Name = "DateTimeStart"
        Me.DateTimeStart.Size = New System.Drawing.Size(119, 22)
        Me.DateTimeStart.TabIndex = 62
        '
        'vorlagenDropbox
        '
        Me.vorlagenDropbox.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.vorlagenDropbox.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.vorlagenDropbox.FormattingEnabled = True
        Me.vorlagenDropbox.Location = New System.Drawing.Point(535, 27)
        Me.vorlagenDropbox.Margin = New System.Windows.Forms.Padding(2)
        Me.vorlagenDropbox.Name = "vorlagenDropbox"
        Me.vorlagenDropbox.Size = New System.Drawing.Size(309, 24)
        Me.vorlagenDropbox.TabIndex = 51
        '
        'Erloes
        '
        Me.Erloes.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Erloes.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Erloes.Location = New System.Drawing.Point(536, 64)
        Me.Erloes.Margin = New System.Windows.Forms.Padding(2)
        Me.Erloes.Name = "Erloes"
        Me.Erloes.Size = New System.Drawing.Size(88, 22)
        Me.Erloes.TabIndex = 55
        '
        'lbl_pName
        '
        Me.lbl_pName.AutoSize = True
        Me.lbl_pName.Enabled = False
        Me.lbl_pName.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lbl_pName.Location = New System.Drawing.Point(21, 30)
        Me.lbl_pName.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.lbl_pName.Name = "lbl_pName"
        Me.lbl_pName.Size = New System.Drawing.Size(45, 16)
        Me.lbl_pName.TabIndex = 48
        Me.lbl_pName.Text = "Name"
        '
        'projectName
        '
        Me.projectName.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.projectName.Location = New System.Drawing.Point(124, 27)
        Me.projectName.Margin = New System.Windows.Forms.Padding(2)
        Me.projectName.Name = "projectName"
        Me.projectName.Size = New System.Drawing.Size(254, 22)
        Me.projectName.TabIndex = 49
        '
        'lblVorlage
        '
        Me.lblVorlage.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.lblVorlage.AutoSize = True
        Me.lblVorlage.Enabled = False
        Me.lblVorlage.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblVorlage.Location = New System.Drawing.Point(449, 30)
        Me.lblVorlage.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.lblVorlage.Name = "lblVorlage"
        Me.lblVorlage.Size = New System.Drawing.Size(56, 16)
        Me.lblVorlage.TabIndex = 50
        Me.lblVorlage.Text = "Vorlage"
        Me.lblVorlage.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        'lblBudget
        '
        Me.lblBudget.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.lblBudget.AutoSize = True
        Me.lblBudget.Enabled = False
        Me.lblBudget.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblBudget.Location = New System.Drawing.Point(450, 70)
        Me.lblBudget.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.lblBudget.Name = "lblBudget"
        Me.lblBudget.Size = New System.Drawing.Size(78, 16)
        Me.lblBudget.TabIndex = 54
        Me.lblBudget.Text = "Budget (T€)"
        '
        'AbbrButton
        '
        Me.AbbrButton.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.AbbrButton.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.AbbrButton.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.AbbrButton.Location = New System.Drawing.Point(725, 302)
        Me.AbbrButton.Margin = New System.Windows.Forms.Padding(2, 3, 2, 3)
        Me.AbbrButton.Name = "AbbrButton"
        Me.AbbrButton.Size = New System.Drawing.Size(119, 22)
        Me.AbbrButton.TabIndex = 69
        Me.AbbrButton.Text = "Abbrechen"
        Me.AbbrButton.UseVisualStyleBackColor = True
        '
        'OKButton
        '
        Me.OKButton.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.OKButton.DialogResult = System.Windows.Forms.DialogResult.OK
        Me.OKButton.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.OKButton.Location = New System.Drawing.Point(124, 302)
        Me.OKButton.Margin = New System.Windows.Forms.Padding(2, 3, 2, 3)
        Me.OKButton.Name = "OKButton"
        Me.OKButton.Size = New System.Drawing.Size(119, 22)
        Me.OKButton.TabIndex = 64
        Me.OKButton.Text = "OK"
        Me.OKButton.UseVisualStyleBackColor = True
        '
        'frmProjektEingabe1
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(96.0!, 96.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Dpi
        Me.BackColor = System.Drawing.SystemColors.Window
        Me.ClientSize = New System.Drawing.Size(889, 354)
        Me.Controls.Add(Me.Panel1)
        Me.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.Margin = New System.Windows.Forms.Padding(2, 3, 2, 3)
        Me.Name = "frmProjektEingabe1"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Neues Projekt anlegen"
        Me.Panel1.ResumeLayout(False)
        Me.Panel1.PerformLayout()
        Me.ResumeLayout(False)

    End Sub

    Friend WithEvents Panel1 As Windows.Forms.Panel
    Friend WithEvents lbl_selProjName As Windows.Forms.Label
    Public WithEvents lbl_Number As Windows.Forms.Label
    Public WithEvents txtbx_pNr As Windows.Forms.TextBox
    Public WithEvents txtbx_description As Windows.Forms.TextBox
    Public WithEvents lbl_Description As Windows.Forms.Label
    Public WithEvents profitAskedFor As Windows.Forms.TextBox
    Public WithEvents lblProfitField As Windows.Forms.Label
    Public WithEvents lbl_Laufzeit As Windows.Forms.Label
    Public WithEvents endMilestoneDropbox As Windows.Forms.ComboBox
    Public WithEvents lbl_Referenz2 As Windows.Forms.Label
    Public WithEvents startMilestoneDropbox As Windows.Forms.ComboBox
    Public WithEvents lbl_Referenz1 As Windows.Forms.Label
    Friend WithEvents DateTimeEnde As Windows.Forms.DateTimePicker
    Friend WithEvents dauerUnverändert As Windows.Forms.CheckBox
    Public WithEvents DateTimeStart As Windows.Forms.DateTimePicker
    Public WithEvents vorlagenDropbox As Windows.Forms.ComboBox
    Public WithEvents Erloes As Windows.Forms.TextBox
    Public WithEvents lbl_pName As Windows.Forms.Label
    Public WithEvents projectName As Windows.Forms.TextBox
    Public WithEvents lblVorlage As Windows.Forms.Label
    Public WithEvents lblBudget As Windows.Forms.Label
    Public WithEvents AbbrButton As Windows.Forms.Button
    Public WithEvents OKButton As Windows.Forms.Button
End Class
