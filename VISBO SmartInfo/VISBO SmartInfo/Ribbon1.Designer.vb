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
        Me.Group2 = Me.Factory.CreateRibbonGroup
        Me.btnUpdate = Me.Factory.CreateRibbonButton
        Me.Group3 = Me.Factory.CreateRibbonGroup
        Me.btnStart = Me.Factory.CreateRibbonButton
        Me.btnFastBack = Me.Factory.CreateRibbonButton
        Me.Button4 = Me.Factory.CreateRibbonButton
        Me.btnFastForward = Me.Factory.CreateRibbonButton
        Me.btnEnd2 = Me.Factory.CreateRibbonButton
        Me.Group4 = Me.Factory.CreateRibbonGroup
        Me.activateTab = Me.Factory.CreateRibbonButton
        Me.activateInfo = Me.Factory.CreateRibbonButton
        Me.SmartInfo = Me.Factory.CreateRibbonGroup
        Me.timeMachineTab = Me.Factory.CreateRibbonButton
        Me.variantTab_Click = Me.Factory.CreateRibbonButton
        Me.settingsTab = Me.Factory.CreateRibbonButton
        Me.Tab1.SuspendLayout()
        Me.Tab2.SuspendLayout()
        Me.Group2.SuspendLayout()
        Me.Group3.SuspendLayout()
        Me.Group4.SuspendLayout()
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
        Me.Tab2.Groups.Add(Me.Group2)
        Me.Tab2.Groups.Add(Me.Group3)
        Me.Tab2.Groups.Add(Me.Group4)
        Me.Tab2.Groups.Add(Me.SmartInfo)
        Me.Tab2.Label = "VISBO"
        Me.Tab2.Name = "Tab2"
        '
        'Group2
        '
        Me.Group2.Items.Add(Me.btnUpdate)
        Me.Group2.Label = "Aktualisieren"
        Me.Group2.Name = "Group2"
        '
        'btnUpdate
        '
        Me.btnUpdate.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge
        Me.btnUpdate.Image = Global.VISBO_SmartInfo.My.Resources.Resources.Visbo_update_Button
        Me.btnUpdate.Label = "Aktuellste Version"
        Me.btnUpdate.Name = "btnUpdate"
        Me.btnUpdate.ShowImage = True
        '
        'Group3
        '
        Me.Group3.Items.Add(Me.btnStart)
        Me.Group3.Items.Add(Me.btnFastBack)
        Me.Group3.Items.Add(Me.Button4)
        Me.Group3.Items.Add(Me.btnFastForward)
        Me.Group3.Items.Add(Me.btnEnd2)
        Me.Group3.Label = "Time Machine"
        Me.Group3.Name = "Group3"
        '
        'btnStart
        '
        Me.btnStart.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge
        Me.btnStart.Image = Global.VISBO_SmartInfo.My.Resources.Resources.navigate_beginning1
        Me.btnStart.Label = "Ursprüngl. Version"
        Me.btnStart.Name = "btnStart"
        Me.btnStart.ShowImage = True
        '
        'btnFastBack
        '
        Me.btnFastBack.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge
        Me.btnFastBack.Image = Global.VISBO_SmartInfo.My.Resources.Resources.navigate_left
        Me.btnFastBack.Label = "Vorherige Version"
        Me.btnFastBack.Name = "btnFastBack"
        Me.btnFastBack.ShowImage = True
        '
        'Button4
        '
        Me.Button4.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge
        Me.Button4.Image = Global.VISBO_SmartInfo.My.Resources.Resources.arrow_down_blue
        Me.Button4.Label = "Differenz zu vorheriger Version"
        Me.Button4.Name = "Button4"
        Me.Button4.ShowImage = True
        '
        'btnFastForward
        '
        Me.btnFastForward.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge
        Me.btnFastForward.Image = Global.VISBO_SmartInfo.My.Resources.Resources.navigate_right
        Me.btnFastForward.Label = "Folgende Version"
        Me.btnFastForward.Name = "btnFastForward"
        Me.btnFastForward.ShowImage = True
        '
        'btnEnd2
        '
        Me.btnEnd2.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge
        Me.btnEnd2.Image = Global.VISBO_SmartInfo.My.Resources.Resources.navigate_end
        Me.btnEnd2.Label = "Aktuellste Version"
        Me.btnEnd2.Name = "btnEnd2"
        Me.btnEnd2.ScreenTip = " fösadlkfsödlkf"
        Me.btnEnd2.ShowImage = True
        '
        'Group4
        '
        Me.Group4.Items.Add(Me.activateTab)
        Me.Group4.Items.Add(Me.activateInfo)
        Me.Group4.Label = "Aktionsbereiche"
        Me.Group4.Name = "Group4"
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
        Me.activateInfo.Image = Global.VISBO_SmartInfo.My.Resources.Resources.layout_center
        Me.activateInfo.Label = "Eigenschaften"
        Me.activateInfo.Name = "activateInfo"
        Me.activateInfo.ShowImage = True
        '
        'SmartInfo
        '
        Me.SmartInfo.Items.Add(Me.timeMachineTab)
        Me.SmartInfo.Items.Add(Me.variantTab_Click)
        Me.SmartInfo.Items.Add(Me.settingsTab)
        Me.SmartInfo.Label = "Smart-Info (alt)"
        Me.SmartInfo.Name = "SmartInfo"
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
        Me.Group2.ResumeLayout(False)
        Me.Group2.PerformLayout()
        Me.Group3.ResumeLayout(False)
        Me.Group3.PerformLayout()
        Me.Group4.ResumeLayout(False)
        Me.Group4.PerformLayout()
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
    Friend WithEvents variantTab_Click As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents activateInfo As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents Group2 As Microsoft.Office.Tools.Ribbon.RibbonGroup
    Friend WithEvents btnUpdate As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents Group3 As Microsoft.Office.Tools.Ribbon.RibbonGroup
    Friend WithEvents btnStart As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents btnFastBack As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents Button4 As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents btnFastForward As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents btnEnd2 As Microsoft.Office.Tools.Ribbon.RibbonButton
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
