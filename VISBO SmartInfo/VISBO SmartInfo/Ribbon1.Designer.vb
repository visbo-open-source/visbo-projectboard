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
    'Nehmen Sie keine Änderungen mit dem Code-Editor vor.
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Me.Tab1 = Me.Factory.CreateRibbonTab
        Me.Group1 = Me.Factory.CreateRibbonGroup
        Me.Tab2 = Me.Factory.CreateRibbonTab
        Me.SmartInfo = Me.Factory.CreateRibbonGroup
        Me.activateTab = Me.Factory.CreateRibbonButton
        Me.timeMachineTab = Me.Factory.CreateRibbonButton
        Me.settingsTab = Me.Factory.CreateRibbonButton
        Me.Tab1.SuspendLayout()
        Me.Tab2.SuspendLayout()
        Me.SmartInfo.SuspendLayout()
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
        'Tab2
        '
        Me.Tab2.Groups.Add(Me.SmartInfo)
        Me.Tab2.Label = "VISBO"
        Me.Tab2.Name = "Tab2"
        '
        'SmartInfo
        '
        Me.SmartInfo.Items.Add(Me.activateTab)
        Me.SmartInfo.Items.Add(Me.timeMachineTab)
        Me.SmartInfo.Items.Add(Me.settingsTab)
        Me.SmartInfo.Label = "Smart-Info"
        Me.SmartInfo.Name = "SmartInfo"
        '
        'activateTab
        '
        Me.activateTab.Label = "Activate"
        Me.activateTab.Name = "activateTab"
        '
        'timeMachineTab
        '
        Me.timeMachineTab.Label = "Time-Machine"
        Me.timeMachineTab.Name = "timeMachineTab"
        '
        'settingsTab
        '
        Me.settingsTab.Label = "Settings"
        Me.settingsTab.Name = "settingsTab"
        '
        'Ribbon1
        '
        Me.Name = "Ribbon1"
        Me.RibbonType = "Microsoft.PowerPoint.Presentation"
        Me.Tabs.Add(Me.Tab1)
        Me.Tabs.Add(Me.Tab2)
        Me.Tab1.ResumeLayout(False)
        Me.Tab1.PerformLayout()
        Me.Tab2.ResumeLayout(False)
        Me.Tab2.PerformLayout()
        Me.SmartInfo.ResumeLayout(False)
        Me.SmartInfo.PerformLayout()

    End Sub

    Friend WithEvents Tab1 As Microsoft.Office.Tools.Ribbon.RibbonTab
    Friend WithEvents Group1 As Microsoft.Office.Tools.Ribbon.RibbonGroup
    Friend WithEvents Tab2 As Microsoft.Office.Tools.Ribbon.RibbonTab
    Friend WithEvents SmartInfo As Microsoft.Office.Tools.Ribbon.RibbonGroup
    Friend WithEvents activateTab As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents settingsTab As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents timeMachineTab As Microsoft.Office.Tools.Ribbon.RibbonButton
End Class

Partial Class ThisRibbonCollection

    <System.Diagnostics.DebuggerNonUserCode()> _
    Friend ReadOnly Property Ribbon1() As Ribbon1
        Get
            Return Me.GetRibbon(Of Ribbon1)()
        End Get
    End Property
End Class
