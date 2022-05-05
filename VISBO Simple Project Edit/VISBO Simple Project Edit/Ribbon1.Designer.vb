Partial Class Ribbon1
    Inherits Microsoft.Office.Tools.Ribbon.RibbonBase

    <System.Diagnostics.DebuggerNonUserCode()> _
    Public Sub New(ByVal container As System.ComponentModel.IContainer)
        MyClass.New()

        'Erforderlich für die Unterstützung des Windows.Forms-Klassenkompositions-Designers
        If (container IsNot Nothing) Then
            container.Add(Me)
        End If

    End Sub

    <System.Diagnostics.DebuggerNonUserCode()> _
    Public Sub New()
        MyBase.New(Globals.Factory.GetRibbonFactory())

        'Dieser Aufruf ist für den Komponenten-Designer erforderlich.
        InitializeComponent()

    End Sub

    'Die Komponente überschreibt den Löschvorgang zum Bereinigen der Komponentenliste.
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

    'Wird vom Komponenten-Designer benötigt.
    Private components As System.ComponentModel.IContainer

    'Hinweis: Die folgende Prozedur ist für den Komponenten-Designer erforderlich.
    'Das Bearbeiten ist mit dem Komponenten-Designer möglich.
    'Das Bearbeiten mit dem Code-Editor ist nicht möglich.
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Me.Tab1 = Me.Factory.CreateRibbonTab
        Me.Group1 = Me.Factory.CreateRibbonGroup
        Me.VISBO = Me.Factory.CreateRibbonTab
        Me.SPEdit = Me.Factory.CreateRibbonGroup
        Me.Laden = Me.Factory.CreateRibbonButton
        Me.Speichern = Me.Factory.CreateRibbonButton
        Me.Delete = Me.Factory.CreateRibbonButton
        Me.Separator1 = Me.Factory.CreateRibbonSeparator
        Me.Cost = Me.Factory.CreateRibbonButton
        Me.Time = Me.Factory.CreateRibbonButton
        Me.Resources = Me.Factory.CreateRibbonButton
        Me.Tab1.SuspendLayout()
        Me.VISBO.SuspendLayout()
        Me.SPEdit.SuspendLayout()
        Me.SuspendLayout()
        '
        'Tab1
        '
        Me.Tab1.ControlId.ControlIdType = Microsoft.Office.Tools.Ribbon.RibbonControlIdType.Office
        Me.Tab1.Groups.Add(Me.Group1)
        Me.Tab1.Label = "TabAddIns"
        Me.Tab1.Name = "Tab1"
        '
        'Group1
        '
        Me.Group1.Label = "Group1"
        Me.Group1.Name = "Group1"
        '
        'VISBO
        '
        Me.VISBO.Groups.Add(Me.SPEdit)
        Me.VISBO.Label = "VISBO"
        Me.VISBO.Name = "VISBO"
        '
        'SPEdit
        '
        Me.SPEdit.Items.Add(Me.Laden)
        Me.SPEdit.Items.Add(Me.Speichern)
        Me.SPEdit.Items.Add(Me.Delete)
        Me.SPEdit.Items.Add(Me.Separator1)
        Me.SPEdit.Items.Add(Me.Cost)
        Me.SPEdit.Items.Add(Me.Time)
        Me.SPEdit.Items.Add(Me.Resources)
        Me.SPEdit.Label = "Simple Project Edit"
        Me.SPEdit.Name = "SPEdit"
        '
        'Laden
        '
        Me.Laden.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge
        Me.Laden.Label = "Laden"
        Me.Laden.Name = "Laden"
        Me.Laden.ShowImage = True
        '
        'Speichern
        '
        Me.Speichern.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge
        Me.Speichern.Label = "Speichern"
        Me.Speichern.Name = "Speichern"
        Me.Speichern.ShowImage = True
        '
        'Delete
        '
        Me.Delete.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge
        Me.Delete.Label = "Löschen"
        Me.Delete.Name = "Delete"
        Me.Delete.ShowImage = True
        '
        'Separator1
        '
        Me.Separator1.Name = "Separator1"
        '
        'Cost
        '
        Me.Cost.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge
        Me.Cost.Image = Global.VISBO_Simple_Project_Edit.My.Resources.Resources.noun_money_1061584
        Me.Cost.Label = "Cost"
        Me.Cost.Name = "Cost"
        Me.Cost.ShowImage = True
        '
        'Time
        '
        Me.Time.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge
        Me.Time.Image = Global.VISBO_Simple_Project_Edit.My.Resources.Resources.noun_stop_watch_2010575
        Me.Time.Label = "Time"
        Me.Time.Name = "Time"
        Me.Time.ShowImage = True
        '
        'Resources
        '
        Me.Resources.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge
        Me.Resources.Image = Global.VISBO_Simple_Project_Edit.My.Resources.Resources.noun_bottleneck_366717
        Me.Resources.Label = "Resources"
        Me.Resources.Name = "Resources"
        Me.Resources.ShowImage = True
        '
        'Ribbon1
        '
        Me.Name = "Ribbon1"
        Me.RibbonType = "Microsoft.Excel.Workbook"
        Me.Tabs.Add(Me.Tab1)
        Me.Tabs.Add(Me.VISBO)
        Me.Tab1.ResumeLayout(False)
        Me.Tab1.PerformLayout()
        Me.VISBO.ResumeLayout(False)
        Me.VISBO.PerformLayout()
        Me.SPEdit.ResumeLayout(False)
        Me.SPEdit.PerformLayout()
        Me.ResumeLayout(False)

    End Sub

    Friend WithEvents Tab1 As Microsoft.Office.Tools.Ribbon.RibbonTab
    Friend WithEvents Group1 As Microsoft.Office.Tools.Ribbon.RibbonGroup
    Friend WithEvents VISBO As Microsoft.Office.Tools.Ribbon.RibbonTab
    Friend WithEvents SPEdit As Microsoft.Office.Tools.Ribbon.RibbonGroup
    Friend WithEvents Laden As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents Speichern As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents Delete As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents Separator1 As Microsoft.Office.Tools.Ribbon.RibbonSeparator
    Friend WithEvents Cost As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents Time As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents Resources As Microsoft.Office.Tools.Ribbon.RibbonButton
End Class

Partial Class ThisRibbonCollection

    <System.Diagnostics.DebuggerNonUserCode()> _
    Friend ReadOnly Property Ribbon1() As Ribbon1
        Get
            Return Me.GetRibbon(Of Ribbon1)()
        End Get
    End Property
End Class
