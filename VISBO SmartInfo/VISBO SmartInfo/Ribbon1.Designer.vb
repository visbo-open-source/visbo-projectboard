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
        Me.activateInfo = Me.Factory.CreateRibbonButton
        Me.timeMachineTab = Me.Factory.CreateRibbonButton
        Me.variantTab_Click = Me.Factory.CreateRibbonButton
        Me.settingsTab = Me.Factory.CreateRibbonButton
        Me.Group2 = Me.Factory.CreateRibbonGroup
        Me.Button1 = Me.Factory.CreateRibbonButton
        Me.Group3 = Me.Factory.CreateRibbonGroup
        Me.Button2 = Me.Factory.CreateRibbonButton
        Me.Button3 = Me.Factory.CreateRibbonButton
        Me.Button4 = Me.Factory.CreateRibbonButton
        Me.Button5 = Me.Factory.CreateRibbonButton
        Me.Button6 = Me.Factory.CreateRibbonButton
        Me.Group4 = Me.Factory.CreateRibbonGroup
        Me.Tab1.SuspendLayout()
        Me.Tab2.SuspendLayout()
        Me.SmartInfo.SuspendLayout()
        Me.Group2.SuspendLayout()
        Me.Group3.SuspendLayout()
        Me.Group4.SuspendLayout()
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
        'Tab2
        '
        Me.Tab2.Groups.Add(Me.Group2)
        Me.Tab2.Groups.Add(Me.Group3)
        Me.Tab2.Groups.Add(Me.Group4)
        Me.Tab2.Groups.Add(Me.SmartInfo)
        Me.Tab2.Label = "VISBO"
        Me.Tab2.Name = "Tab2"
        '
        'SmartInfo
        '
        Me.SmartInfo.Items.Add(Me.timeMachineTab)
        Me.SmartInfo.Items.Add(Me.variantTab_Click)
        Me.SmartInfo.Items.Add(Me.settingsTab)
        Me.SmartInfo.Label = "Smart-Info (alt)"
        Me.SmartInfo.Name = "SmartInfo"
        '
        'activateTab
        '
        Me.activateTab.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge
        Me.activateTab.Image = Global.VISBO_SmartInfo.My.Resources.Resources.view
        Me.activateTab.Label = "Search"
        Me.activateTab.Name = "activateTab"
        Me.activateTab.ShowImage = True
        '
        'activateInfo
        '
        Me.activateInfo.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge
        Me.activateInfo.Label = "Eigenschaften"
        Me.activateInfo.Name = "activateInfo"
        Me.activateInfo.ShowImage = True
        '
        'timeMachineTab
        '
        Me.timeMachineTab.Label = "Time-Machine"
        Me.timeMachineTab.Name = "timeMachineTab"
        '
        'variantTab_Click
        '
        Me.variantTab_Click.Label = "Varianten"
        Me.variantTab_Click.Name = "variantTab_Click"
        '
        'settingsTab
        '
        Me.settingsTab.Label = "Settings"
        Me.settingsTab.Name = "settingsTab"
        '
        'Group2
        '
        Me.Group2.Items.Add(Me.Button1)
        Me.Group2.Label = "Aktualisieren"
        Me.Group2.Name = "Group2"
        '
        'Button1
        '
        Me.Button1.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge
        Me.Button1.Label = "Aktuellste Version"
        Me.Button1.Name = "Button1"
        Me.Button1.ShowImage = True
        '
        'Group3
        '
        Me.Group3.Items.Add(Me.Button2)
        Me.Group3.Items.Add(Me.Button3)
        Me.Group3.Items.Add(Me.Button4)
        Me.Group3.Items.Add(Me.Button5)
        Me.Group3.Items.Add(Me.Button6)
        Me.Group3.Label = "Time Machine"
        Me.Group3.Name = "Group3"
        '
        'Button2
        '
        Me.Button2.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge
        Me.Button2.Label = "Ursprüngl. Version"
        Me.Button2.Name = "Button2"
        Me.Button2.ShowImage = True
        '
        'Button3
        '
        Me.Button3.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge
        Me.Button3.Label = "Vorherige Version"
        Me.Button3.Name = "Button3"
        Me.Button3.ShowImage = True
        '
        'Button4
        '
        Me.Button4.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge
        Me.Button4.Label = "Differenz zu vorheriger Version"
        Me.Button4.Name = "Button4"
        Me.Button4.ShowImage = True
        '
        'Button5
        '
        Me.Button5.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge
        Me.Button5.Label = "Folgende Version"
        Me.Button5.Name = "Button5"
        Me.Button5.ShowImage = True
        '
        'Button6
        '
        Me.Button6.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge
        Me.Button6.Label = "Aktuellste Version"
        Me.Button6.Name = "Button6"
        Me.Button6.ShowImage = True
        '
        'Group4
        '
        Me.Group4.Items.Add(Me.activateTab)
        Me.Group4.Items.Add(Me.activateInfo)
        Me.Group4.Label = "Aktionsbereiche"
        Me.Group4.Name = "Group4"
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
        Me.Group2.ResumeLayout(False)
        Me.Group2.PerformLayout()
        Me.Group3.ResumeLayout(False)
        Me.Group3.PerformLayout()
        Me.Group4.ResumeLayout(False)
        Me.Group4.PerformLayout()
        Me.ResumeLayout(False)

    End Sub

    Friend WithEvents Tab1 As Microsoft.Office.Tools.Ribbon.RibbonTab
    Friend WithEvents Group1 As Microsoft.Office.Tools.Ribbon.RibbonGroup
    Friend WithEvents Tab2 As Microsoft.Office.Tools.Ribbon.RibbonTab
    Friend WithEvents SmartInfo As Microsoft.Office.Tools.Ribbon.RibbonGroup
    Friend WithEvents activateTab As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents settingsTab As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents timeMachineTab As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents variantTab_Click As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents activateInfo As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents Group2 As Microsoft.Office.Tools.Ribbon.RibbonGroup
    Friend WithEvents Button1 As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents Group3 As Microsoft.Office.Tools.Ribbon.RibbonGroup
    Friend WithEvents Button2 As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents Button3 As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents Button4 As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents Button5 As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents Button6 As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents Group4 As Microsoft.Office.Tools.Ribbon.RibbonGroup
End Class

Partial Class ThisRibbonCollection

    <System.Diagnostics.DebuggerNonUserCode()> _
    Friend ReadOnly Property Ribbon1() As Ribbon1
        Get
            Return Me.GetRibbon(Of Ribbon1)()
        End Get
    End Property
End Class
