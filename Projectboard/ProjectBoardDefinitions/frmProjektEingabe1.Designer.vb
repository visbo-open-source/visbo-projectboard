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
        Me.kennzeichnungDate = New System.Windows.Forms.Label()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.projectName = New System.Windows.Forms.TextBox()
        Me.pName = New System.Windows.Forms.Label()
        Me.selectedMonth = New System.Windows.Forms.NumericUpDown()
        Me.calcMonth = New System.Windows.Forms.Label()
        Me.Erloes = New System.Windows.Forms.TextBox()
        Me.sFit = New System.Windows.Forms.TextBox()
        Me.risiko = New System.Windows.Forms.TextBox()
        Me.vorlagenDropbox = New System.Windows.Forms.ComboBox()
        Me.Label5 = New System.Windows.Forms.Label()
        Me.volume = New System.Windows.Forms.TextBox()
        Me.DateTimeProject = New System.Windows.Forms.DateTimePicker()
        Me.dauerUnverändert = New System.Windows.Forms.CheckBox()
        Me.kennzeichnungEnde = New System.Windows.Forms.Label()
        Me.DateTimeEnde = New System.Windows.Forms.DateTimePicker()
        CType(Me.selectedMonth, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'OKButton
        '
        Me.OKButton.DialogResult = System.Windows.Forms.DialogResult.OK
        Me.OKButton.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.OKButton.Location = New System.Drawing.Point(45, 351)
        Me.OKButton.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.OKButton.Name = "OKButton"
        Me.OKButton.Size = New System.Drawing.Size(87, 28)
        Me.OKButton.TabIndex = 1
        Me.OKButton.Text = "OK"
        Me.OKButton.UseVisualStyleBackColor = True
        '
        'AbbrButton
        '
        Me.AbbrButton.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.AbbrButton.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.AbbrButton.Location = New System.Drawing.Point(367, 351)
        Me.AbbrButton.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.AbbrButton.Name = "AbbrButton"
        Me.AbbrButton.Size = New System.Drawing.Size(101, 28)
        Me.AbbrButton.TabIndex = 10
        Me.AbbrButton.Text = "Abbrechen"
        Me.AbbrButton.UseVisualStyleBackColor = True
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Enabled = False
        Me.Label1.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label1.Location = New System.Drawing.Point(41, 143)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(98, 20)
        Me.Label1.TabIndex = 5
        Me.Label1.Text = "Budget (T€)"
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Enabled = False
        Me.Label2.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label2.Location = New System.Drawing.Point(41, 178)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(133, 20)
        Me.Label2.TabIndex = 6
        Me.Label2.Text = "Strategischer Fit"
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Enabled = False
        Me.Label3.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label3.Location = New System.Drawing.Point(41, 213)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(155, 20)
        Me.Label3.TabIndex = 7
        Me.Label3.Text = "Umsetzungs-Risiko"
        '
        'kennzeichnungDate
        '
        Me.kennzeichnungDate.AutoSize = True
        Me.kennzeichnungDate.Enabled = False
        Me.kennzeichnungDate.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.kennzeichnungDate.Location = New System.Drawing.Point(41, 248)
        Me.kennzeichnungDate.Name = "kennzeichnungDate"
        Me.kennzeichnungDate.Size = New System.Drawing.Size(50, 20)
        Me.kennzeichnungDate.TabIndex = 9
        Me.kennzeichnungDate.Text = "Start "
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Enabled = False
        Me.Label4.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label4.Location = New System.Drawing.Point(41, 71)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(94, 20)
        Me.Label4.TabIndex = 14
        Me.Label4.Text = "Projekt-Typ"
        '
        'projectName
        '
        Me.projectName.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.projectName.Location = New System.Drawing.Point(167, 33)
        Me.projectName.Name = "projectName"
        Me.projectName.Size = New System.Drawing.Size(259, 26)
        Me.projectName.TabIndex = 0
        '
        'pName
        '
        Me.pName.AutoSize = True
        Me.pName.Enabled = False
        Me.pName.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.pName.Location = New System.Drawing.Point(41, 36)
        Me.pName.Name = "pName"
        Me.pName.Size = New System.Drawing.Size(111, 20)
        Me.pName.TabIndex = 16
        Me.pName.Text = "Projekt-Name"
        '
        'selectedMonth
        '
        Me.selectedMonth.Location = New System.Drawing.Point(256, 353)
        Me.selectedMonth.Maximum = New Decimal(New Integer() {120, 0, 0, 0})
        Me.selectedMonth.Name = "selectedMonth"
        Me.selectedMonth.Size = New System.Drawing.Size(17, 27)
        Me.selectedMonth.TabIndex = 6
        Me.selectedMonth.Visible = False
        '
        'calcMonth
        '
        Me.calcMonth.AutoSize = True
        Me.calcMonth.Enabled = False
        Me.calcMonth.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.calcMonth.Location = New System.Drawing.Point(184, 355)
        Me.calcMonth.Name = "calcMonth"
        Me.calcMonth.Size = New System.Drawing.Size(66, 20)
        Me.calcMonth.TabIndex = 19
        Me.calcMonth.Text = "Mon YY"
        Me.calcMonth.Visible = False
        '
        'Erloes
        '
        Me.Erloes.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Erloes.Location = New System.Drawing.Point(199, 143)
        Me.Erloes.Name = "Erloes"
        Me.Erloes.Size = New System.Drawing.Size(74, 26)
        Me.Erloes.TabIndex = 20
        '
        'sFit
        '
        Me.sFit.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.sFit.Location = New System.Drawing.Point(199, 175)
        Me.sFit.Name = "sFit"
        Me.sFit.Size = New System.Drawing.Size(74, 26)
        Me.sFit.TabIndex = 21
        '
        'risiko
        '
        Me.risiko.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.risiko.Location = New System.Drawing.Point(199, 214)
        Me.risiko.Name = "risiko"
        Me.risiko.Size = New System.Drawing.Size(74, 26)
        Me.risiko.TabIndex = 22
        '
        'vorlagenDropbox
        '
        Me.vorlagenDropbox.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.vorlagenDropbox.FormattingEnabled = True
        Me.vorlagenDropbox.Location = New System.Drawing.Point(167, 65)
        Me.vorlagenDropbox.Name = "vorlagenDropbox"
        Me.vorlagenDropbox.Size = New System.Drawing.Size(259, 28)
        Me.vorlagenDropbox.TabIndex = 23
        '
        'Label5
        '
        Me.Label5.AutoEllipsis = True
        Me.Label5.AutoSize = True
        Me.Label5.Enabled = False
        Me.Label5.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label5.Location = New System.Drawing.Point(297, 178)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(74, 20)
        Me.Label5.TabIndex = 24
        Me.Label5.Text = "Volumen"
        '
        'volume
        '
        Me.volume.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.volume.Location = New System.Drawing.Point(394, 172)
        Me.volume.Name = "volume"
        Me.volume.Size = New System.Drawing.Size(74, 26)
        Me.volume.TabIndex = 25
        '
        'DateTimeProject
        '
        Me.DateTimeProject.Location = New System.Drawing.Point(209, 248)
        Me.DateTimeProject.Name = "DateTimeProject"
        Me.DateTimeProject.Size = New System.Drawing.Size(259, 27)
        Me.DateTimeProject.TabIndex = 26
        '
        'dauerUnverändert
        '
        Me.dauerUnverändert.AutoSize = True
        Me.dauerUnverändert.Checked = True
        Me.dauerUnverändert.CheckState = System.Windows.Forms.CheckState.Checked
        Me.dauerUnverändert.Location = New System.Drawing.Point(301, 213)
        Me.dauerUnverändert.Name = "dauerUnverändert"
        Me.dauerUnverändert.Size = New System.Drawing.Size(169, 25)
        Me.dauerUnverändert.TabIndex = 27
        Me.dauerUnverändert.Text = "Dauer wie Vorlage"
        Me.dauerUnverändert.UseVisualStyleBackColor = True
        '
        'kennzeichnungEnde
        '
        Me.kennzeichnungEnde.AutoSize = True
        Me.kennzeichnungEnde.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.kennzeichnungEnde.ForeColor = System.Drawing.SystemColors.ControlText
        Me.kennzeichnungEnde.Location = New System.Drawing.Point(42, 285)
        Me.kennzeichnungEnde.Name = "kennzeichnungEnde"
        Me.kennzeichnungEnde.Size = New System.Drawing.Size(47, 20)
        Me.kennzeichnungEnde.TabIndex = 28
        Me.kennzeichnungEnde.Text = "Ende"
        '
        'DateTimeEnde
        '
        Me.DateTimeEnde.Location = New System.Drawing.Point(209, 285)
        Me.DateTimeEnde.Name = "DateTimeEnde"
        Me.DateTimeEnde.Size = New System.Drawing.Size(259, 27)
        Me.DateTimeEnde.TabIndex = 29
        '
        'frmProjektEingabe1
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(9.0!, 19.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.SystemColors.Window
        Me.ClientSize = New System.Drawing.Size(557, 401)
        Me.Controls.Add(Me.DateTimeEnde)
        Me.Controls.Add(Me.kennzeichnungEnde)
        Me.Controls.Add(Me.dauerUnverändert)
        Me.Controls.Add(Me.DateTimeProject)
        Me.Controls.Add(Me.volume)
        Me.Controls.Add(Me.Label5)
        Me.Controls.Add(Me.vorlagenDropbox)
        Me.Controls.Add(Me.risiko)
        Me.Controls.Add(Me.sFit)
        Me.Controls.Add(Me.Erloes)
        Me.Controls.Add(Me.calcMonth)
        Me.Controls.Add(Me.selectedMonth)
        Me.Controls.Add(Me.pName)
        Me.Controls.Add(Me.projectName)
        Me.Controls.Add(Me.Label4)
        Me.Controls.Add(Me.kennzeichnungDate)
        Me.Controls.Add(Me.Label3)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.AbbrButton)
        Me.Controls.Add(Me.OKButton)
        Me.Font = New System.Drawing.Font("Tahoma", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.Name = "frmProjektEingabe1"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.Manual
        Me.Text = "Neues Projekt anlegen"
        CType(Me.selectedMonth, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Public WithEvents OKButton As System.Windows.Forms.Button
    Public WithEvents AbbrButton As System.Windows.Forms.Button
    Public WithEvents Label1 As System.Windows.Forms.Label
    Public WithEvents Label2 As System.Windows.Forms.Label
    Public WithEvents Label3 As System.Windows.Forms.Label
    Public WithEvents kennzeichnungDate As System.Windows.Forms.Label
    Public WithEvents Label4 As System.Windows.Forms.Label
    Public WithEvents projectName As System.Windows.Forms.TextBox
    Public WithEvents pName As System.Windows.Forms.Label
    Public WithEvents selectedMonth As System.Windows.Forms.NumericUpDown
    Public WithEvents calcMonth As System.Windows.Forms.Label
    Public WithEvents Erloes As System.Windows.Forms.TextBox
    Public WithEvents sFit As System.Windows.Forms.TextBox
    Public WithEvents risiko As System.Windows.Forms.TextBox
    Public WithEvents vorlagenDropbox As System.Windows.Forms.ComboBox
    Public WithEvents Label5 As System.Windows.Forms.Label
    Public WithEvents volume As System.Windows.Forms.TextBox
    Public WithEvents DateTimeProject As System.Windows.Forms.DateTimePicker
    Friend WithEvents dauerUnverändert As System.Windows.Forms.CheckBox
    Friend WithEvents DateTimeEnde As System.Windows.Forms.DateTimePicker
    Public WithEvents kennzeichnungEnde As System.Windows.Forms.Label
End Class
