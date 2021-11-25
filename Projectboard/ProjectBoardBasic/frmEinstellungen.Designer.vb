<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmEinstellungen
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmEinstellungen))
        Me.chkboxPropAnpass = New System.Windows.Forms.CheckBox()
        Me.chkboxAmpel = New System.Windows.Forms.CheckBox()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.SprachAusw = New System.Windows.Forms.ComboBox()
        Me.statusLabel = New System.Windows.Forms.Label()
        Me.rdbFirst = New System.Windows.Forms.RadioButton()
        Me.rdbLast = New System.Windows.Forms.RadioButton()
        Me.GroupBox1 = New System.Windows.Forms.GroupBox()
        Me.loadPFV = New System.Windows.Forms.CheckBox()
        Me.chkbxPhasesAnteilig = New System.Windows.Forms.CheckBox()
        Me.chkbxInvoices = New System.Windows.Forms.CheckBox()
        Me.chkbx_KUG_active = New System.Windows.Forms.CheckBox()
        Me.chkbx_TakeCapaFromOldOrga = New System.Windows.Forms.CheckBox()
        Me.chkbx_autoSetActualDataDate = New System.Windows.Forms.CheckBox()
        Me.GroupBox1.SuspendLayout()
        Me.SuspendLayout()
        '
        'chkboxPropAnpass
        '
        Me.chkboxPropAnpass.AutoSize = True
        Me.chkboxPropAnpass.Location = New System.Drawing.Point(16, 93)
        Me.chkboxPropAnpass.Name = "chkboxPropAnpass"
        Me.chkboxPropAnpass.Size = New System.Drawing.Size(230, 17)
        Me.chkboxPropAnpass.TabIndex = 1
        Me.chkboxPropAnpass.Text = "Ressourcen-Bedarfe proportional anpassen"
        Me.chkboxPropAnpass.UseVisualStyleBackColor = True
        '
        'chkboxAmpel
        '
        Me.chkboxAmpel.AutoSize = True
        Me.chkboxAmpel.Location = New System.Drawing.Point(16, 137)
        Me.chkboxAmpel.Name = "chkboxAmpel"
        Me.chkboxAmpel.Size = New System.Drawing.Size(101, 17)
        Me.chkboxAmpel.TabIndex = 2
        Me.chkboxAmpel.Text = "Ampel anzeigen"
        Me.chkboxAmpel.UseVisualStyleBackColor = True
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!)
        Me.Label1.Location = New System.Drawing.Point(13, 268)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(116, 15)
        Me.Label1.TabIndex = 4
        Me.Label1.Text = "Sprache für Reports"
        '
        'SprachAusw
        '
        Me.SprachAusw.AutoCompleteMode = System.Windows.Forms.AutoCompleteMode.Suggest
        Me.SprachAusw.FormattingEnabled = True
        Me.SprachAusw.Location = New System.Drawing.Point(146, 265)
        Me.SprachAusw.MaxDropDownItems = 4
        Me.SprachAusw.Name = "SprachAusw"
        Me.SprachAusw.Size = New System.Drawing.Size(158, 21)
        Me.SprachAusw.TabIndex = 5
        '
        'statusLabel
        '
        Me.statusLabel.AutoSize = True
        Me.statusLabel.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.statusLabel.Location = New System.Drawing.Point(13, 297)
        Me.statusLabel.Name = "statusLabel"
        Me.statusLabel.Size = New System.Drawing.Size(45, 15)
        Me.statusLabel.TabIndex = 45
        Me.statusLabel.Text = "Label1"
        Me.statusLabel.TextAlign = System.Drawing.ContentAlignment.BottomLeft
        Me.statusLabel.Visible = False
        '
        'rdbFirst
        '
        Me.rdbFirst.AutoSize = True
        Me.rdbFirst.Location = New System.Drawing.Point(13, 19)
        Me.rdbFirst.Name = "rdbFirst"
        Me.rdbFirst.Size = New System.Drawing.Size(52, 17)
        Me.rdbFirst.TabIndex = 46
        Me.rdbFirst.TabStop = True
        Me.rdbFirst.Text = "Erster"
        Me.rdbFirst.UseVisualStyleBackColor = True
        '
        'rdbLast
        '
        Me.rdbLast.AutoSize = True
        Me.rdbLast.Location = New System.Drawing.Point(132, 19)
        Me.rdbLast.Name = "rdbLast"
        Me.rdbLast.Size = New System.Drawing.Size(53, 17)
        Me.rdbLast.TabIndex = 47
        Me.rdbLast.TabStop = True
        Me.rdbLast.Text = "letzter"
        Me.rdbLast.UseVisualStyleBackColor = True
        '
        'GroupBox1
        '
        Me.GroupBox1.Controls.Add(Me.rdbLast)
        Me.GroupBox1.Controls.Add(Me.rdbFirst)
        Me.GroupBox1.Location = New System.Drawing.Point(14, 9)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(309, 51)
        Me.GroupBox1.TabIndex = 48
        Me.GroupBox1.TabStop = False
        Me.GroupBox1.Text = "Vergleich mit welcher Version"
        '
        'loadPFV
        '
        Me.loadPFV.AutoSize = True
        Me.loadPFV.Location = New System.Drawing.Point(16, 71)
        Me.loadPFV.Name = "loadPFV"
        Me.loadPFV.Size = New System.Drawing.Size(125, 17)
        Me.loadPFV.TabIndex = 49
        Me.loadPFV.Text = "immer Vorgabe laden"
        Me.loadPFV.UseVisualStyleBackColor = True
        '
        'chkbxPhasesAnteilig
        '
        Me.chkbxPhasesAnteilig.AutoSize = True
        Me.chkbxPhasesAnteilig.Location = New System.Drawing.Point(16, 115)
        Me.chkbxPhasesAnteilig.Name = "chkbxPhasesAnteilig"
        Me.chkbxPhasesAnteilig.Size = New System.Drawing.Size(314, 17)
        Me.chkbxPhasesAnteilig.TabIndex = 50
        Me.chkbxPhasesAnteilig.Text = "Phasen in Monats-Häufigkeitsdiagrammen anteilig berechnen"
        Me.chkbxPhasesAnteilig.UseVisualStyleBackColor = True
        '
        'chkbxInvoices
        '
        Me.chkbxInvoices.AutoSize = True
        Me.chkbxInvoices.Location = New System.Drawing.Point(16, 159)
        Me.chkbxInvoices.Name = "chkbxInvoices"
        Me.chkbxInvoices.Size = New System.Drawing.Size(195, 17)
        Me.chkbxInvoices.TabIndex = 51
        Me.chkbxInvoices.Text = "Rechnungen / Penalties bearbeiten"
        Me.chkbxInvoices.UseVisualStyleBackColor = True
        '
        'chkbx_KUG_active
        '
        Me.chkbx_KUG_active.AutoSize = True
        Me.chkbx_KUG_active.Location = New System.Drawing.Point(16, 182)
        Me.chkbx_KUG_active.Name = "chkbx_KUG_active"
        Me.chkbx_KUG_active.Size = New System.Drawing.Size(125, 17)
        Me.chkbx_KUG_active.TabIndex = 52
        Me.chkbx_KUG_active.Text = "Kurzarbeit ist möglich"
        Me.chkbx_KUG_active.UseVisualStyleBackColor = True
        '
        'chkbx_TakeCapaFromOldOrga
        '
        Me.chkbx_TakeCapaFromOldOrga.AutoSize = True
        Me.chkbx_TakeCapaFromOldOrga.Location = New System.Drawing.Point(16, 205)
        Me.chkbx_TakeCapaFromOldOrga.Name = "chkbx_TakeCapaFromOldOrga"
        Me.chkbx_TakeCapaFromOldOrga.Size = New System.Drawing.Size(277, 17)
        Me.chkbx_TakeCapaFromOldOrga.TabIndex = 53
        Me.chkbx_TakeCapaFromOldOrga.Text = "Kapazitäten aus bisheriger Organisation übernehmen "
        Me.chkbx_TakeCapaFromOldOrga.UseVisualStyleBackColor = True
        '
        'chkbx_autoSetActualDataDate
        '
        Me.chkbx_autoSetActualDataDate.AutoSize = True
        Me.chkbx_autoSetActualDataDate.Location = New System.Drawing.Point(16, 229)
        Me.chkbx_autoSetActualDataDate.Name = "chkbx_autoSetActualDataDate"
        Me.chkbx_autoSetActualDataDate.Size = New System.Drawing.Size(295, 17)
        Me.chkbx_autoSetActualDataDate.TabIndex = 54
        Me.chkbx_autoSetActualDataDate.Text = "Daten aus Vergangenheit explizit als Ist-Daten bestätigen"
        Me.chkbx_autoSetActualDataDate.UseVisualStyleBackColor = True
        '
        'frmEinstellungen
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(352, 327)
        Me.Controls.Add(Me.chkbx_autoSetActualDataDate)
        Me.Controls.Add(Me.chkbx_TakeCapaFromOldOrga)
        Me.Controls.Add(Me.chkbx_KUG_active)
        Me.Controls.Add(Me.chkbxInvoices)
        Me.Controls.Add(Me.chkbxPhasesAnteilig)
        Me.Controls.Add(Me.loadPFV)
        Me.Controls.Add(Me.GroupBox1)
        Me.Controls.Add(Me.statusLabel)
        Me.Controls.Add(Me.SprachAusw)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.chkboxAmpel)
        Me.Controls.Add(Me.chkboxPropAnpass)
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.Name = "frmEinstellungen"
        Me.Text = "Einstellungen"
        Me.GroupBox1.ResumeLayout(False)
        Me.GroupBox1.PerformLayout()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents chkboxPropAnpass As System.Windows.Forms.CheckBox
    Friend WithEvents chkboxAmpel As System.Windows.Forms.CheckBox
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents SprachAusw As System.Windows.Forms.ComboBox
    Friend WithEvents statusLabel As System.Windows.Forms.Label
    Friend WithEvents rdbFirst As Windows.Forms.RadioButton
    Friend WithEvents rdbLast As Windows.Forms.RadioButton
    Friend WithEvents GroupBox1 As Windows.Forms.GroupBox
    Friend WithEvents loadPFV As Windows.Forms.CheckBox
    Friend WithEvents chkbxPhasesAnteilig As Windows.Forms.CheckBox
    Friend WithEvents chkbxInvoices As Windows.Forms.CheckBox
    Friend WithEvents chkbx_KUG_active As Windows.Forms.CheckBox
    Friend WithEvents chkbx_TakeCapaFromOldOrga As Windows.Forms.CheckBox
    Friend WithEvents chkbx_autoSetActualDataDate As Windows.Forms.CheckBox
End Class
