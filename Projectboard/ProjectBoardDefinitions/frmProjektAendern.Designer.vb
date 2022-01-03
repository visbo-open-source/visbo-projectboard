<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmProjektAendern
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmProjektAendern))
        Me.pName = New System.Windows.Forms.Label()
        Me.projectName = New System.Windows.Forms.TextBox()
        Me.risiko = New System.Windows.Forms.TextBox()
        Me.sFit = New System.Windows.Forms.TextBox()
        Me.Erloes = New System.Windows.Forms.TextBox()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.ruleEngine = New System.Windows.Forms.LinkLabel()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.AbbrButton = New System.Windows.Forms.Button()
        Me.OKButton = New System.Windows.Forms.Button()
        Me.vorlagenName = New System.Windows.Forms.TextBox()
        Me.Label5 = New System.Windows.Forms.Label()
        Me.businessUnit = New System.Windows.Forms.ComboBox()
        Me.SuspendLayout()
        '
        'pName
        '
        Me.pName.AutoSize = True
        Me.pName.Enabled = False
        Me.pName.Location = New System.Drawing.Point(30, 31)
        Me.pName.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.pName.Name = "pName"
        Me.pName.Size = New System.Drawing.Size(86, 16)
        Me.pName.TabIndex = 18
        Me.pName.Text = "Projekt-Name"
        '
        'projectName
        '
        Me.projectName.Location = New System.Drawing.Point(149, 28)
        Me.projectName.Margin = New System.Windows.Forms.Padding(2, 3, 2, 3)
        Me.projectName.Name = "projectName"
        Me.projectName.Size = New System.Drawing.Size(217, 23)
        Me.projectName.TabIndex = 17
        '
        'risiko
        '
        Me.risiko.Location = New System.Drawing.Point(306, 175)
        Me.risiko.Margin = New System.Windows.Forms.Padding(2)
        Me.risiko.Name = "risiko"
        Me.risiko.Size = New System.Drawing.Size(60, 23)
        Me.risiko.TabIndex = 35
        '
        'sFit
        '
        Me.sFit.Location = New System.Drawing.Point(306, 144)
        Me.sFit.Margin = New System.Windows.Forms.Padding(2)
        Me.sFit.Name = "sFit"
        Me.sFit.Size = New System.Drawing.Size(60, 23)
        Me.sFit.TabIndex = 34
        '
        'Erloes
        '
        Me.Erloes.Location = New System.Drawing.Point(306, 115)
        Me.Erloes.Margin = New System.Windows.Forms.Padding(2)
        Me.Erloes.Name = "Erloes"
        Me.Erloes.Size = New System.Drawing.Size(60, 23)
        Me.Erloes.TabIndex = 33
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Enabled = False
        Me.Label4.Font = New System.Drawing.Font("Tahoma", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label4.Location = New System.Drawing.Point(30, 59)
        Me.Label4.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(74, 16)
        Me.Label4.TabIndex = 32
        Me.Label4.Text = "Projekt-Typ"
        '
        'ruleEngine
        '
        Me.ruleEngine.AutoSize = True
        Me.ruleEngine.Location = New System.Drawing.Point(30, 222)
        Me.ruleEngine.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.ruleEngine.Name = "ruleEngine"
        Me.ruleEngine.Size = New System.Drawing.Size(135, 16)
        Me.ruleEngine.TabIndex = 30
        Me.ruleEngine.TabStop = True
        Me.ruleEngine.Text = "Regeln und Prämissen"
        Me.ruleEngine.Visible = False
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Enabled = False
        Me.Label3.Location = New System.Drawing.Point(30, 178)
        Me.Label3.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(115, 16)
        Me.Label3.TabIndex = 27
        Me.Label3.Text = "Umsetzungs-Risiko"
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Enabled = False
        Me.Label2.Location = New System.Drawing.Point(30, 147)
        Me.Label2.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(102, 16)
        Me.Label2.TabIndex = 26
        Me.Label2.Text = "Strategischer Fit"
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Enabled = False
        Me.Label1.Location = New System.Drawing.Point(30, 118)
        Me.Label1.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(76, 16)
        Me.Label1.TabIndex = 25
        Me.Label1.Text = "Budget (T€)"
        '
        'AbbrButton
        '
        Me.AbbrButton.Location = New System.Drawing.Point(254, 251)
        Me.AbbrButton.Margin = New System.Windows.Forms.Padding(2, 3, 2, 3)
        Me.AbbrButton.Name = "AbbrButton"
        Me.AbbrButton.Size = New System.Drawing.Size(113, 25)
        Me.AbbrButton.TabIndex = 31
        Me.AbbrButton.Text = "Abbrechen"
        Me.AbbrButton.UseVisualStyleBackColor = True
        '
        'OKButton
        '
        Me.OKButton.Location = New System.Drawing.Point(33, 251)
        Me.OKButton.Margin = New System.Windows.Forms.Padding(2, 3, 2, 3)
        Me.OKButton.Name = "OKButton"
        Me.OKButton.Size = New System.Drawing.Size(70, 25)
        Me.OKButton.TabIndex = 23
        Me.OKButton.Text = "OK"
        Me.OKButton.UseVisualStyleBackColor = True
        '
        'vorlagenName
        '
        Me.vorlagenName.Enabled = False
        Me.vorlagenName.Location = New System.Drawing.Point(149, 56)
        Me.vorlagenName.Margin = New System.Windows.Forms.Padding(2, 3, 2, 3)
        Me.vorlagenName.Name = "vorlagenName"
        Me.vorlagenName.Size = New System.Drawing.Size(217, 23)
        Me.vorlagenName.TabIndex = 36
        '
        'Label5
        '
        Me.Label5.AutoSize = True
        Me.Label5.Location = New System.Drawing.Point(30, 87)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(83, 16)
        Me.Label5.TabIndex = 37
        Me.Label5.Text = "Business Unit"
        '
        'businessUnit
        '
        Me.businessUnit.FormattingEnabled = True
        Me.businessUnit.Location = New System.Drawing.Point(150, 84)
        Me.businessUnit.Name = "businessUnit"
        Me.businessUnit.Size = New System.Drawing.Size(216, 24)
        Me.businessUnit.TabIndex = 38
        '
        'frmProjektAendern
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(96.0!, 96.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Dpi
        Me.ClientSize = New System.Drawing.Size(392, 288)
        Me.Controls.Add(Me.businessUnit)
        Me.Controls.Add(Me.Label5)
        Me.Controls.Add(Me.vorlagenName)
        Me.Controls.Add(Me.risiko)
        Me.Controls.Add(Me.sFit)
        Me.Controls.Add(Me.Erloes)
        Me.Controls.Add(Me.Label4)
        Me.Controls.Add(Me.ruleEngine)
        Me.Controls.Add(Me.Label3)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.AbbrButton)
        Me.Controls.Add(Me.OKButton)
        Me.Controls.Add(Me.pName)
        Me.Controls.Add(Me.projectName)
        Me.Font = New System.Drawing.Font("Tahoma", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.Margin = New System.Windows.Forms.Padding(2, 3, 2, 3)
        Me.Name = "frmProjektAendern"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Daten ändern"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Public WithEvents pName As System.Windows.Forms.Label
    Public WithEvents projectName As System.Windows.Forms.TextBox
    Public WithEvents risiko As System.Windows.Forms.TextBox
    Public WithEvents sFit As System.Windows.Forms.TextBox
    Public WithEvents Erloes As System.Windows.Forms.TextBox
    Public WithEvents Label4 As System.Windows.Forms.Label
    Public WithEvents ruleEngine As System.Windows.Forms.LinkLabel
    Public WithEvents Label3 As System.Windows.Forms.Label
    Public WithEvents Label2 As System.Windows.Forms.Label
    Public WithEvents Label1 As System.Windows.Forms.Label
    Public WithEvents AbbrButton As System.Windows.Forms.Button
    Public WithEvents OKButton As System.Windows.Forms.Button
    Public WithEvents vorlagenName As System.Windows.Forms.TextBox
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Public WithEvents businessUnit As System.Windows.Forms.ComboBox
End Class
