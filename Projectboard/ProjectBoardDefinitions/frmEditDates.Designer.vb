<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmEditDates
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmEditDates))
        Me.startdatePicker = New System.Windows.Forms.DateTimePicker()
        Me.enddatePicker = New System.Windows.Forms.DateTimePicker()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.lblElemName = New System.Windows.Forms.Label()
        Me.btn_OK = New System.Windows.Forms.Button()
        Me.btn_Cancel = New System.Windows.Forms.Button()
        Me.chkbx_adjustChilds = New System.Windows.Forms.CheckBox()
        Me.chkbxAutoDistr = New System.Windows.Forms.CheckBox()
        Me.SuspendLayout()
        '
        'startdatePicker
        '
        Me.startdatePicker.Location = New System.Drawing.Point(19, 48)
        Me.startdatePicker.Name = "startdatePicker"
        Me.startdatePicker.Size = New System.Drawing.Size(200, 20)
        Me.startdatePicker.TabIndex = 0
        '
        'enddatePicker
        '
        Me.enddatePicker.Location = New System.Drawing.Point(246, 48)
        Me.enddatePicker.Name = "enddatePicker"
        Me.enddatePicker.Size = New System.Drawing.Size(200, 20)
        Me.enddatePicker.TabIndex = 1
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(224, 52)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(16, 13)
        Me.Label1.TabIndex = 2
        Me.Label1.Text = " - "
        '
        'lblElemName
        '
        Me.lblElemName.AutoSize = True
        Me.lblElemName.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblElemName.Location = New System.Drawing.Point(19, 15)
        Me.lblElemName.Name = "lblElemName"
        Me.lblElemName.Size = New System.Drawing.Size(49, 16)
        Me.lblElemName.TabIndex = 3
        Me.lblElemName.Text = "Label2"
        '
        'btn_OK
        '
        Me.btn_OK.Location = New System.Drawing.Point(120, 99)
        Me.btn_OK.Name = "btn_OK"
        Me.btn_OK.Size = New System.Drawing.Size(99, 23)
        Me.btn_OK.TabIndex = 4
        Me.btn_OK.Text = "OK"
        Me.btn_OK.UseVisualStyleBackColor = True
        '
        'btn_Cancel
        '
        Me.btn_Cancel.Location = New System.Drawing.Point(246, 99)
        Me.btn_Cancel.Name = "btn_Cancel"
        Me.btn_Cancel.Size = New System.Drawing.Size(99, 23)
        Me.btn_Cancel.TabIndex = 5
        Me.btn_Cancel.Text = "Abbruch"
        Me.btn_Cancel.UseVisualStyleBackColor = True
        '
        'chkbx_adjustChilds
        '
        Me.chkbx_adjustChilds.AutoSize = True
        Me.chkbx_adjustChilds.Enabled = False
        Me.chkbx_adjustChilds.Location = New System.Drawing.Point(246, 76)
        Me.chkbx_adjustChilds.Name = "chkbx_adjustChilds"
        Me.chkbx_adjustChilds.Size = New System.Drawing.Size(175, 17)
        Me.chkbx_adjustChilds.TabIndex = 6
        Me.chkbx_adjustChilds.Text = """Kinder"" automatisch anpassen"
        Me.chkbx_adjustChilds.UseVisualStyleBackColor = True
        Me.chkbx_adjustChilds.Visible = False
        '
        'chkbxAutoDistr
        '
        Me.chkbxAutoDistr.AutoSize = True
        Me.chkbxAutoDistr.Location = New System.Drawing.Point(22, 76)
        Me.chkbxAutoDistr.Name = "chkbxAutoDistr"
        Me.chkbxAutoDistr.Size = New System.Drawing.Size(191, 17)
        Me.chkbxAutoDistr.TabIndex = 7
        Me.chkbxAutoDistr.Text = "Ress.- und Kosten autom. verteilen"
        Me.chkbxAutoDistr.UseVisualStyleBackColor = True
        Me.chkbxAutoDistr.Visible = False
        '
        'frmEditDates
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(473, 145)
        Me.Controls.Add(Me.chkbxAutoDistr)
        Me.Controls.Add(Me.chkbx_adjustChilds)
        Me.Controls.Add(Me.btn_Cancel)
        Me.Controls.Add(Me.btn_OK)
        Me.Controls.Add(Me.lblElemName)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.enddatePicker)
        Me.Controls.Add(Me.startdatePicker)
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.Name = "frmEditDates"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "edit Dates"
        Me.TopMost = True
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents Label1 As Windows.Forms.Label
    Friend WithEvents btn_OK As Windows.Forms.Button
    Friend WithEvents btn_Cancel As Windows.Forms.Button
    Public WithEvents startdatePicker As Windows.Forms.DateTimePicker
    Public WithEvents enddatePicker As Windows.Forms.DateTimePicker
    Public WithEvents lblElemName As Windows.Forms.Label
    Public WithEvents chkbx_adjustChilds As Windows.Forms.CheckBox
    Public WithEvents chkbxAutoDistr As Windows.Forms.CheckBox
End Class
