<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmSelectOneItem
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmSelectOneItem))
        Me.OKBtn = New System.Windows.Forms.Button()
        Me.itemList = New System.Windows.Forms.ListBox()
        Me.SuspendLayout()
        '
        'OKBtn
        '
        Me.OKBtn.DialogResult = System.Windows.Forms.DialogResult.OK
        Me.OKBtn.Location = New System.Drawing.Point(114, 203)
        Me.OKBtn.Name = "OKBtn"
        Me.OKBtn.Size = New System.Drawing.Size(108, 23)
        Me.OKBtn.TabIndex = 0
        Me.OKBtn.Text = "OK"
        Me.OKBtn.UseVisualStyleBackColor = True
        '
        'itemList
        '
        Me.itemList.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.itemList.FormattingEnabled = True
        Me.itemList.ItemHeight = 16
        Me.itemList.Location = New System.Drawing.Point(12, 16)
        Me.itemList.Name = "itemList"
        Me.itemList.Size = New System.Drawing.Size(308, 164)
        Me.itemList.Sorted = True
        Me.itemList.TabIndex = 1
        '
        'frmSelectOneItem
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(336, 248)
        Me.Controls.Add(Me.itemList)
        Me.Controls.Add(Me.OKBtn)
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.Name = "frmSelectOneItem"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Wählen Sie ein VISBO Project Center"
        Me.ResumeLayout(False)

    End Sub

    Friend WithEvents OKBtn As Windows.Forms.Button
    Public WithEvents itemList As Windows.Forms.ListBox
End Class
